"""
PPTX to PDF 转换服务

基于 COM 技术，使用 FastAPI 提供 HTTP 接口，将 PPTX 文件转换为 PDF。
需要 Windows 环境 + 已安装 Microsoft PowerPoint。

用法：
    python main.py
    或
    uvicorn main:app --host 0.0.0.0 --port 8000
"""

import asyncio
import logging
import os
import tempfile
import threading
import uuid
from concurrent.futures import Future
from contextlib import asynccontextmanager
from dataclasses import dataclass
from queue import Queue
from typing import Optional

import psutil
from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask

# ---------------------------------------------------------------------------
# 日志配置
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("main.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# PDF 格式常量
PP_SAVE_AS_PDF = 32

# 转换超时（秒）
CONVERT_TIMEOUT = 120

# 内存限制（MB）：Python 进程 + POWERPNT.EXE 之和超限后主动退出
MEMORY_LIMIT_MB = int(os.environ.get("MEMORY_LIMIT_MB", "2048"))
SHOW_PPT = os.environ.get("SHOW_PPT", "0").strip() in ("1", "true", "True")


# ---------------------------------------------------------------------------
# 内存监控
# ---------------------------------------------------------------------------
def _total_memory_mb() -> float:
    """Return combined RSS of this process and any running POWERPNT.EXE."""
    py_mb = psutil.Process().memory_info().rss / 1024 / 1024
    ppt_mb = sum(
        p.info["memory_info"].rss / 1024 / 1024
        for p in psutil.process_iter(["name", "memory_info"])
        if p.info["name"] and "POWERPNT" in p.info["name"].upper()
        and p.info["memory_info"] is not None
    )
    return py_mb + ppt_mb


def _check_memory_and_exit() -> None:
    """Called as a background task after each response; triggers SIGTERM if over limit."""
    import signal
    mb = _total_memory_mb()
    logger.info(f"memory {mb:.1f}MB / limit {MEMORY_LIMIT_MB}MB")
    if mb > MEMORY_LIMIT_MB:
        logger.warning(f"memory limit exceeded ({mb:.1f}MB > {MEMORY_LIMIT_MB}MB), triggering exit")
        os.kill(os.getpid(), signal.SIGTERM)


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------
@dataclass
class ConvertTask:
    """一个转换任务"""
    input_path: str
    output_path: str
    future: Future


# ---------------------------------------------------------------------------
# COM Worker — 专用线程管理 PowerPoint 生命周期
# ---------------------------------------------------------------------------
class ComWorker:
    """
    在专用线程中运行 COM 操作，确保线程安全。

    PowerPoint COM 对象只能在同一线程中使用，
    本类通过队列将转换任务派发到专用线程执行。
    """

    _SENTINEL = object()  # 停止信号

    def __init__(self) -> None:
        self._queue: Queue = Queue()
        self._thread: Optional[threading.Thread] = None
        self._ready = threading.Event()

    def start(self) -> None:
        """启动 COM 工作线程"""
        self._thread = threading.Thread(
            target=self._worker_loop,
            name="com-worker",
            daemon=True,
        )
        self._thread.start()
        # 等待线程初始化完成
        self._ready.wait(timeout=30)
        logger.info("COM Worker 已启动")

    def _worker_loop(self) -> None:
        """工作线程主循环"""
        import pythoncom
        import win32com.client

        powerpoint = None
        try:
            # 初始化 COM
            pythoncom.CoInitialize()
            logger.info("COM 已初始化")

            # 创建 PowerPoint 实例
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            if SHOW_PPT:
                powerpoint.Visible = True
            else:
                powerpoint.Visible = False
            logger.info(f"PowerPoint 实例已创建 (Visible={SHOW_PPT})")

            self._ready.set()

            # 循环处理任务
            while True:
                task = self._queue.get()

                # 收到停止信号
                if task is self._SENTINEL:
                    logger.info("收到停止信号，COM Worker 退出中...")
                    break

                self._handle_task(powerpoint, task)

        except Exception as e:
            logger.error(f"COM Worker 异常: {e}")
            self._ready.set()  # 即使失败也要释放等待
        finally:
            if powerpoint is not None:
                try:
                    powerpoint.Quit()
                    logger.info("PowerPoint 已退出")
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _handle_task(self, powerpoint, task: ConvertTask) -> None:
        """处理单个转换任务"""
        presentation = None
        try:
            logger.info(f"开始转换: {os.path.basename(task.input_path)}")

            presentation = powerpoint.Presentations.Open(
                task.input_path,
                ReadOnly=True,
                Untitled=False,
                WithWindow=False,
            )

            presentation.SaveAs(task.output_path, PP_SAVE_AS_PDF)

            logger.info(f"转换完成: {os.path.basename(task.output_path)}")
            task.future.set_result(True)

        except Exception as e:
            logger.error(f"转换失败: {e}")
            task.future.set_exception(e)

        finally:
            if presentation is not None:
                try:
                    presentation.Close()
                except Exception:
                    pass

    def convert(self, input_path: str, output_path: str) -> Future:
        """
        提交转换任务，返回 Future。

        Args:
            input_path: PPTX 文件的绝对路径
            output_path: 输出 PDF 文件的绝对路径

        Returns:
            Future 对象，可用于等待结果
        """
        future = Future()
        task = ConvertTask(
            input_path=os.path.abspath(input_path),
            output_path=os.path.abspath(output_path),
            future=future,
        )
        self._queue.put(task)
        return future

    def shutdown(self) -> None:
        """停止工作线程"""
        self._queue.put(self._SENTINEL)
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=30)
            logger.info("COM Worker 已停止")


# ---------------------------------------------------------------------------
# FastAPI 应用
# ---------------------------------------------------------------------------

# 全局 COM Worker 实例
com_worker: Optional[ComWorker] = None


@asynccontextmanager
async def lifespan(app: FastAPI):
    """管理应用生命周期"""
    global com_worker
    com_worker = ComWorker()
    com_worker.start()
    logger.info("服务已启动")

    yield

    com_worker.shutdown()
    logger.info("服务已关闭")


app = FastAPI(
    title="PPTX to PDF 转换服务",
    description="基于 COM 技术，将 PowerPoint 文件转换为 PDF",
    version="1.0.0",
    lifespan=lifespan,
)


@app.get("/health")
async def health():
    """健康检查，附带内存信息"""
    return {"status": "ok", "memory_mb": round(_total_memory_mb(), 1)}


@app.post("/convert")
async def convert(file: UploadFile):
    """
    上传 PPTX 文件，返回转换后的 PDF。

    - **file**: PowerPoint 文件 (.pptx / .ppt)
    """
    # 验证文件类型
    filename = file.filename or "upload.pptx"
    ext = os.path.splitext(filename)[1].lower()
    if ext not in (".pptx", ".ppt"):
        raise HTTPException(
            status_code=400,
            detail=f"不支持的文件类型: {ext}，仅支持 .pptx 和 .ppt",
        )

    if com_worker is None:
        raise HTTPException(status_code=503, detail="服务未就绪")

    # 创建临时目录
    work_dir = tempfile.mkdtemp(prefix="pptx2pdf_")
    task_id = uuid.uuid4().hex[:8]
    input_path = os.path.join(work_dir, f"{task_id}{ext}")
    output_path = os.path.join(work_dir, f"{task_id}.pdf")
    pdf_name = os.path.splitext(filename)[0] + ".pdf"

    try:
        # 保存上传文件
        content = await file.read()
        with open(input_path, "wb") as f:
            f.write(content)
        logger.info(f"收到文件: {filename} ({len(content)} bytes)")

        # 提交转换任务并等待结果
        future = com_worker.convert(input_path, output_path)

        # 等待 COM 线程完成转换
        # wrap_future 将 concurrent.futures.Future 桥接为 asyncio.Future
        # COM 线程调用 set_result() 时，会通过 call_soon_threadsafe 通知事件循环
        await asyncio.wait_for(
            asyncio.wrap_future(future),
            timeout=CONVERT_TIMEOUT,
        )

        # 返回 PDF 文件；响应发出后清理临时文件并检查内存
        return FileResponse(
            path=output_path,
            filename=pdf_name,
            media_type="application/pdf",
            background=BackgroundTask(_cleanup_and_check, work_dir),
        )

    except asyncio.TimeoutError:
        # 清理临时文件
        _cleanup(work_dir)
        raise HTTPException(status_code=504, detail="转换超时")

    except HTTPException:
        _cleanup(work_dir)
        raise

    except Exception as e:
        _cleanup(work_dir)
        logger.error(f"转换出错: {e}")
        raise HTTPException(status_code=500, detail=f"转换失败: {e}")


def _cleanup(path: str) -> None:
    """安全删除临时目录"""
    import shutil
    try:
        shutil.rmtree(path, ignore_errors=True)
    except Exception:
        pass


def _cleanup_and_check(path: str) -> None:
    """清理临时目录后检查内存，超限则触发退出"""
    _cleanup(path)
    _check_memory_and_exit()


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
