# PPTX to PDF 转换服务

基于 COM 技术在 Windows 上运行的 FastAPI 服务，完成 PPTX 到 PDF 的转换。  
简单就是美。

## 前置条件

- Windows 系统
- 已安装 Microsoft PowerPoint
- Python 3.12+
- [uv](https://docs.astral.sh/uv/)

## 快速开始

```bash
# 安装依赖
uv sync

# 启动服务
uv run python main.py
```

服务启动后访问 http://localhost:8000/docs 查看 API 文档。

## API

### 健康检查

```bash
curl http://localhost:8000/health
```

### 转换文件

```bash
curl -X POST http://localhost:8000/convert \
     -F "file=@演示文稿.pptx" \
     -o 输出.pdf
```

或使用测试脚本：

```bash
bash test.sh 演示文稿.pptx
```

## 架构

```
POST /convert → FastAPI → 任务队列 → COM 工作线程 (PowerPoint) → PDF
```

- **单 PowerPoint 实例**：在专用线程中复用，避免反复启动/关闭
- **队列串行处理**：确保 COM 线程安全，不会因并发导致崩溃
- **优雅退出**：服务关闭时自动退出 PowerPoint

## 注意事项

- COM 要求使用绝对路径，程序已自动处理
- PowerPoint 以后台模式运行（Visible = False）
- 默认转换超时 120 秒
- 串行处理，适合中低并发场景
