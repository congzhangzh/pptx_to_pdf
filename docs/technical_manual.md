# PPTX to PDF 转换服务 — 技术手册

> 简单就是美。本文档记录本项目的核心技术决策、关键知识点和踩坑经验。

---

## 目录

1. [COM 自动化基础](#1-com-自动化基础)
2. [COM 线程安全模型](#2-com-线程安全模型)
3. [FastAPI 线程模型](#3-fastapi-线程模型)
4. [跨线程 Future 桥接](#4-跨线程-future-桥接)
5. [FileResponse 与临时文件生命周期](#5-fileresponse-与临时文件生命周期)
6. [PowerPoint COM 调用注意事项](#6-powerpoint-com-调用注意事项)
7. [架构设计决策](#7-架构设计决策)

---

## 1. COM 自动化基础

### 什么是 COM？

COM (Component Object Model) 是 Microsoft 的二进制接口标准，允许不同语言编写的组件相互通信。PowerPoint 通过 COM 暴露了完整的自动化接口，Python 可以通过 `pywin32`（`win32com.client`）来调用。

### Python 中的两个 COM 库

| 库 | 包名 | 特点 |
|---|---|---|
| `win32com.client` | `pywin32` | 标准选择，成熟稳定，基于 `IDispatch` 接口 |
| `comtypes` | `comtypes` | 适合需要访问高级 COM 接口的场景，一般自动化任务中属于"杀鸡用牛刀" |

**本项目选择 `pywin32`**，因为 PowerPoint 自动化只需要 `IDispatch`，不需要更底层的接口。

### 核心转换代码

```python
import win32com.client

powerpoint = win32com.client.Dispatch("PowerPoint.Application")
presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
presentation.SaveAs(output_path, 32)  # 32 = ppSaveAsPDF
presentation.Close()
powerpoint.Quit()
```

`32` 是 PowerPoint 的 `ppSaveAsPDF` 常量，代表导出为 PDF 格式。

---

## 2. COM 线程安全模型

### 核心问题

> **COM 对象不是线程安全的。**

COM 使用"套间"（Apartment）模型来管理线程安全：

| 模型 | 名称 | 含义 |
|---|---|---|
| STA | Single-Threaded Apartment | 对象只能在创建它的线程中使用 |
| MTA | Multi-Threaded Apartment | 对象可以在任意线程中使用（但仍需同步） |

PowerPoint COM 对象使用 **STA 模型**，这意味着：

1. 每个线程必须调用 `CoInitialize()` 或 `CoInitializeEx()` 来初始化 COM
2. COM 对象**只能在创建它的线程中使用**
3. 跨线程传递 COM 对象会导致 `RPC_E_WRONG_THREAD` 错误或直接崩溃

### 常见错误模式

```python
# ❌ 错误：在主线程创建，在其他线程使用
powerpoint = win32com.client.Dispatch("PowerPoint.Application")

def worker():
    # 这里使用 powerpoint 对象会崩溃！
    presentation = powerpoint.Presentations.Open(...)

threading.Thread(target=worker).start()
```

```python
# ❌ 错误：在 FastAPI 多线程中共享
@app.post("/convert")
def convert(file: UploadFile):
    # FastAPI 的同步端点在线程池中运行
    # 多个请求 = 多个线程同时访问同一个 COM 对象 → 崩溃
    presentation = shared_powerpoint.Presentations.Open(...)
```

### 正确做法：专用线程

```python
# ✅ 正确：所有 COM 操作在同一线程中完成
def com_worker_loop():
    pythoncom.CoInitialize()  # 初始化 COM
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")

    while True:
        task = queue.get()  # 从队列取任务
        # 在同一线程中执行所有 COM 操作
        presentation = powerpoint.Presentations.Open(...)
        presentation.SaveAs(...)
        presentation.Close()

    powerpoint.Quit()
    pythoncom.CoUninitialize()
```

---

## 3. FastAPI 线程模型

### async def vs def

FastAPI 对两种端点定义有完全不同的执行策略：

```
┌─────────────────────────────────────────┐
│            FastAPI 服务器               │
│                                         │
│  ┌─────────────────────────────────┐    │
│  │   asyncio 事件循环（主线程）     │    │
│  │                                 │    │
│  │  async def endpoint() ← 在此运行 │    │
│  └─────────────────────────────────┘    │
│                                         │
│  ┌─────────────────────────────────┐    │
│  │   线程池 (ThreadPoolExecutor)    │    │
│  │                                 │    │
│  │  def endpoint() ← 在此运行      │    │
│  │  Thread-1: 请求A                │    │
│  │  Thread-2: 请求B                │    │
│  │  Thread-3: 请求C                │    │
│  └─────────────────────────────────┘    │
└─────────────────────────────────────────┘
```

| 定义方式 | 运行位置 | 适用场景 |
|---------|---------|---------|
| `async def` | asyncio 事件循环（主线程） | I/O 密集型、`await` 异步操作 |
| `def` | 线程池 | CPU 密集型、阻塞 I/O |

### 关键推论

1. **`async def` 端点中不能执行阻塞操作**：因为它直接运行在事件循环上，阻塞 = 阻塞整个服务
2. **`def` 端点是多线程并发的**：多个请求会在不同线程中同时运行，共享状态需要加锁
3. **本项目使用 `async def`**：因为我们通过 `await` 等待 COM 线程完成，完全非阻塞

---

## 4. 跨线程 Future 桥接

### 问题

我们的架构中存在两个世界：

- **asyncio 世界**：FastAPI 端点，运行在事件循环上
- **线程世界**：COM Worker，运行在专用线程中

需要一种机制让 async 端点等待线程中的结果。

### 方案对比

#### ❌ 方案 A：`run_in_executor` + `future.result()`

```python
loop = asyncio.get_event_loop()
await loop.run_in_executor(None, future.result)
```

工作原理：
1. 把 `future.result()`（阻塞调用）扔到线程池
2. 线程池中的一个线程被阻塞等待结果
3. 结果到达后，通过 asyncio 通知事件循环

**缺点**：浪费一个线程池线程，它什么也不做，只是阻塞等待。

#### ✅ 方案 B：`asyncio.wrap_future()`

```python
await asyncio.wrap_future(future)
```

工作原理：
1. `wrap_future` 把 `concurrent.futures.Future` 包装为 `asyncio.Future`
2. 当 COM 线程调用 `future.set_result()` 时，内部自动调用 `loop.call_soon_threadsafe()` 通知事件循环
3. **零线程消耗**，纯事件驱动

```
方案 A                           方案 B
┌──────────┐                    ┌──────────┐
│ 事件循环  │                    │ 事件循环  │
│ await ──────→ 线程池线程       │ await ←─────── call_soon_threadsafe
│          │   (阻塞等待)       │          │
└──────────┘   ↑                └──────────┘     ↑
               │                                 │
          set_result()                      set_result()
               │                                 │
         ┌─────┴──────┐                    ┌─────┴──────┐
         │ COM 线程    │                    │ COM 线程    │
         └────────────┘                    └────────────┘

     多一个线程被浪费               零额外线程消耗
```

### `concurrent.futures.Future` 的线程安全性

`concurrent.futures.Future` **天生就是线程安全的**，它的 `set_result()`、`set_exception()`、`result()` 等方法内部都有锁保护。这是标准库的设计保证，可以放心跨线程使用。

> 注意：`asyncio.Future` 则**不是**线程安全的，它只能在事件循环所在线程中操作。
> `wrap_future` 正是负责在两者之间做安全桥接。

### 带超时的完整模式

```python
await asyncio.wait_for(
    asyncio.wrap_future(future),
    timeout=120,
)
```

`wait_for` 提供超时保护。如果 COM 转换卡死（比如 PowerPoint 弹出对话框），不会让请求无限等待。

---

## 5. FileResponse 与临时文件生命周期

### 问题

转换流程中产生临时文件（上传的 PPTX 和生成的 PDF），需要在响应完成后清理。但清理时机很关键：

```python
# ❌ 错误：文件还没发送完就被删了
try:
    return FileResponse(path=output_path)
finally:
    shutil.rmtree(work_dir)  # FileResponse 是流式的，这里删除太早！
```

`FileResponse` 不是立即读取文件内容的，它是**流式发送**。在 `return` 之后，文件还需要保持存在直到响应发送完毕。

### 解决方案：BackgroundTask

Starlette 的 `BackgroundTask` 在**响应完全发送到客户端之后**才执行：

```python
from starlette.background import BackgroundTask

return FileResponse(
    path=output_path,
    filename=pdf_name,
    media_type="application/pdf",
    background=BackgroundTask(_cleanup, work_dir),  # 发送完毕后才清理
)
```

时序图：

```
客户端请求 → 生成 PDF → 开始流式发送 → 发送完毕 → BackgroundTask 清理临时文件
                                                   ↑ 这里才安全
```

### 错误路径的清理

如果转换失败（超时、异常），`FileResponse` 不会被返回，所以需要在 `except` 块中手动清理：

```python
except asyncio.TimeoutError:
    _cleanup(work_dir)       # 手动清理
    raise HTTPException(...)

except Exception as e:
    _cleanup(work_dir)       # 手动清理
    raise HTTPException(...)
```

---

## 6. PowerPoint COM 调用注意事项

### 必须使用绝对路径

COM 的当前工作目录与 Python 的 `os.getcwd()` **不是同一个东西**。传递相对路径会导致"文件未找到"错误：

```python
# ❌ 相对路径 → COM 找不到文件
presentation = powerpoint.Presentations.Open("temp/input.pptx")

# ✅ 绝对路径
presentation = powerpoint.Presentations.Open(r"C:\temp\input.pptx")
```

本项目在 `ComWorker.convert()` 中自动转换：

```python
task = ConvertTask(
    input_path=os.path.abspath(input_path),
    output_path=os.path.abspath(output_path),
    ...
)
```

### 后台模式运行

```python
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
powerpoint.Visible = False  # 不显示窗口
```

如果不设置 `Visible = False`，PowerPoint 窗口会弹出，在服务器环境中可能导致各种问题。

### 打开文件时的参数

```python
presentation = powerpoint.Presentations.Open(
    input_path,
    ReadOnly=True,      # 只读打开，避免锁定文件
    Untitled=False,     # 不作为"无标题"文档
    WithWindow=False,   # 不创建窗口（进一步减少 GUI 交互）
)
```

### 确保关闭资源

COM 对象如果不正确关闭，PowerPoint 进程会残留在后台：

```python
# ✅ 总是在 finally 中关闭
presentation = None
try:
    presentation = powerpoint.Presentations.Open(...)
    presentation.SaveAs(...)
finally:
    if presentation:
        presentation.Close()
```

服务关闭时同样要确保退出 PowerPoint：

```python
# 在 shutdown 中
powerpoint.Quit()
pythoncom.CoUninitialize()
```

### PowerPoint 弹框问题

在某些情况下（文件损坏、宏安全提示等），PowerPoint 可能会弹出对话框，导致自动化卡住。缓解措施：

1. **超时保护**：`asyncio.wait_for(..., timeout=120)` 确保不会无限等待
2. **后台运行**：`Visible = False` + `WithWindow = False` 减少 GUI 交互
3. **只读打开**：`ReadOnly = True` 避免"是否保存"提示

---

## 7. 架构设计决策

### 为什么选择单文件架构？

```
pptx_to_pdf/
├── main.py             # 所有逻辑
├── pyproject.toml      # 依赖
├── test.sh             # 测试
└── readme.md           # 文档
```

本服务功能单一（上传 PPTX → 返回 PDF），代码量约 300 行。拆分成多个模块反而增加了认知负担，不符合"简单就是美"的原则。

### 为什么复用 PowerPoint 实例？

| 策略 | 启动耗时 | 内存 | 稳定性 |
|------|---------|------|--------|
| 每次请求创建/销毁 | 高（2-5 秒启动） | 波动大 | 低（可能残留进程） |
| **复用单实例** | **首次启动后零开销** | **稳定** | **高** |
| 多实例池 | 高初始成本 | 高 | 中等（需要复杂管理） |

复用单实例是性能和复杂度的最佳平衡点。

### 为什么使用队列串行处理？

PowerPoint 即使在同一线程中，同时打开多个文件进行 `SaveAs` 也可能出现不稳定行为。串行处理：

- **可靠**：一次只做一件事
- **简单**：不需要信号量、锁或池管理
- **可预测**：每个请求的资源消耗恒定

如果需要更高吞吐量，可以启动多个服务实例（进程级并行），每个进程有独立的 PowerPoint 实例。

### 为什么用 `uv` 管理环境？

| 特性 | pip + venv | uv |
|------|-----------|-----|
| 依赖解析速度 | 慢 | 极快（Rust 实现） |
| 锁文件 | 无原生支持 | `uv.lock` 自动生成 |
| 环境管理 | 手动 `python -m venv` | `uv sync` 一步完成 |
| 可复现性 | 依赖 `requirements.txt` 手动维护 | 锁文件保证精确复现 |

---

## 附录：完整数据流

```
客户端                  FastAPI (事件循环)              COM Worker (专用线程)
  │                         │                              │
  │── POST /convert ──────→│                              │
  │   (上传 .pptx)          │                              │
  │                         │── 保存到临时目录              │
  │                         │── 创建 Future                │
  │                         │── 放入队列 ────────────────→ │
  │                         │                              │── 取出任务
  │                         │   await wrap_future(future)  │── 打开 PPTX
  │                         │   (非阻塞等待)               │── SaveAs PDF
  │                         │                              │── 关闭 PPTX
  │                         │                              │── set_result()
  │                         │   ←── call_soon_threadsafe ──│
  │                         │── Future 完成                │
  │                         │── FileResponse (流式发送)    │
  │←── PDF 文件 ────────────│                              │
  │                         │── BackgroundTask: 清理临时文件│
  │                         │                              │
```
