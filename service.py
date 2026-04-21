"""
PPTX to PDF + OCR forwarding service

Receives a PPTX file, converts it to PDF via PowerPoint COM, then forwards
the PDF to an external OCR engine and returns the OCR JSON result.

Requires Windows + Microsoft PowerPoint installed.

Usage:
    uv run python service.py

Environment variables:
    OCR_URL          URL of the OCR endpoint (default: http://localhost:9000/ocr)
    MEMORY_LIMIT_MB  Combined Python + POWERPNT.EXE RSS limit in MB (default: 2048)
    PORT             Listening port (default: 8000)
"""

import asyncio
import logging
import os
import tempfile
import threading
import uuid
from itertools import count

from concurrent.futures import Future
from contextlib import asynccontextmanager
from dataclasses import dataclass
from queue import Queue
from typing import Optional

import httpx
import psutil
from fastapi import FastAPI, HTTPException, UploadFile
from starlette.background import BackgroundTask

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO, 
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("service.log", encoding="utf-8"),
        # logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

PP_SAVE_AS_PDF = 32
CONVERT_TIMEOUT = 120
OCR_TIMEOUT = 60

OCR_URL = os.environ.get("OCR_URL", "http://localhost:9000/ocr")
MEMORY_LIMIT_MB = int(os.environ.get("MEMORY_LIMIT_MB", "2048"))
PORT = int(os.environ.get("PORT", "8000"))
# SHOW_PPT = os.environ.get("SHOW_PPT", "0").strip() in ("1", "true", "True")


# ---------------------------------------------------------------------------
# Memory monitoring
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
    """Trigger SIGTERM if combined memory exceeds the configured limit."""
    import signal
    import subprocess
    mb = _total_memory_mb()
    logger.info(f"memory {mb:.1f}MB / limit {MEMORY_LIMIT_MB}MB")
    if mb > MEMORY_LIMIT_MB:
        logger.warning(f"memory limit exceeded ({mb:.1f}MB > {MEMORY_LIMIT_MB}MB), triggering exit")
        # 先温和关闭 COM worker
        if com_worker:
            try:
                com_worker.shutdown()
            except Exception as e:
                logger.error(f"COM worker shutdown error: {e}")
        # 强制杀 PowerPoint
        try:
            subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], 
                         capture_output=True, timeout=5)
        except Exception:
            pass
        # 最后杀自己
        os.kill(os.getpid(), signal.SIGTERM)


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------
@dataclass
class ConvertTask:
    input_path: str
    output_path: str
    future: Future


# ---------------------------------------------------------------------------
# COM Worker — dedicated thread owns the PowerPoint lifecycle
# ---------------------------------------------------------------------------
class ComWorker:
    _SENTINEL = object()

    def __init__(self) -> None:
        self._queue: Queue = Queue()
        self._thread: Optional[threading.Thread] = None
        self._ready = threading.Event()

    def start(self) -> None:
        self._thread = threading.Thread(target=self._worker_loop, name="com-worker", daemon=True)
        self._thread.start()
        self._ready.wait(timeout=30)
        logger.info("COM worker started")

    def _worker_loop(self) -> None:
        import pythoncom
        import win32com.client

        powerpoint = None
        try:
            pythoncom.CoInitialize()
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            # if SHOW_PPT:
            #     powerpoint.Visible = True
            # else:
            #     powerpoint.Visible = False
            # logger.info(f"PowerPoint instance ready (Visible={SHOW_PPT})")
            self._ready.set()
            save_as_pdf_counter=0
            while True:
                task = self._queue.get()
                if task is self._SENTINEL:
                    logger.info("COM worker stopping")
                    break
                save_as_pdf_counter += 1
                logger.info(f"--begin-- [worker] save_as_pdf_counter={save_as_pdf_counter}")
                self._handle_task(powerpoint, task)
                logger.info(f"--end-- [worker] save_as_pdf_counter={save_as_pdf_counter}")

        except Exception as e:
            logger.error(f"COM worker error: {e}")
            self._ready.set()
        finally:
            if powerpoint is not None:
                try:
                    powerpoint.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            # 强制杀死可能残留的 PowerPoint 进程
            import subprocess
            try:
                subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], 
                             capture_output=True, timeout=5)
            except Exception:
                pass

    def _handle_task(self, powerpoint, task: ConvertTask) -> None:
        presentation = None
        try:
            logger.info(f"converting {os.path.basename(task.input_path)}")
            presentation = powerpoint.Presentations.Open(
                task.input_path, ReadOnly=True, Untitled=False, WithWindow=False
            )
            presentation.SaveAs(task.output_path, PP_SAVE_AS_PDF)
            logger.info(f"converted -> {os.path.basename(task.output_path)}")
            task.future.set_result(True)
        except Exception as e:
            logger.error(f"conversion failed: {e}")
            task.future.set_exception(e)
        finally:
            if presentation is not None:
                try:
                    presentation.Close()
                except Exception:
                    pass

    def convert(self, input_path: str, output_path: str) -> Future:
        future: Future = Future()
        self._queue.put(ConvertTask(
            input_path=os.path.abspath(input_path),
            output_path=os.path.abspath(output_path),
            future=future,
        ))
        return future

    def shutdown(self) -> None:
        self._queue.put(self._SENTINEL)
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=30)


# ---------------------------------------------------------------------------
# FastAPI
# ---------------------------------------------------------------------------
com_worker: Optional[ComWorker] = None


@asynccontextmanager
async def lifespan(app: FastAPI):
    global com_worker
    com_worker = ComWorker()
    com_worker.start()
    logger.info(f"service started, OCR_URL={OCR_URL}")
    yield
    com_worker.shutdown()
    logger.info("service stopped")


app = FastAPI(
    title="PPTX to PDF + OCR Service",
    description="Converts PPTX to PDF via COM, then forwards to an OCR engine.",
    version="1.0.0",
    lifespan=lifespan,
)


@app.get("/health")
async def health():
    return {"status": "ok", "memory_mb": round(_total_memory_mb(), 1), "ocr_url": OCR_URL}

api_ppt_counter_gen = count(1)
@app.post("/ppt")
async def convert_ppt(file: UploadFile):
    """
    Upload a PPT/PPTX file; returns OCR JSON result.
    """
    local_api_ppt_counter = next(api_ppt_counter_gen)
    try:
        print(f"--begin-- [api] api_ppt_counter={local_api_ppt_counter}")
        filename = file.filename or "upload.pptx"
        ext = os.path.splitext(filename)[1].lower()
        if ext not in (".pptx", ".ppt"):
            raise HTTPException(status_code=400, detail=f"unsupported file type: {ext}")

        if com_worker is None:
            raise HTTPException(status_code=503, detail="service not ready")

        work_dir = tempfile.mkdtemp(prefix="pptx2pdf_")
        task_id = uuid.uuid4().hex[:8]
        input_path = os.path.join(work_dir, f"{task_id}{ext}")
        output_path = os.path.join(work_dir, f"{task_id}.pdf")
        pdf_name = os.path.splitext(filename)[0] + ".pdf"

        try:
            content = await file.read()
            with open(input_path, "wb") as f:
                f.write(content)
            logger.info(f"received {filename} ({len(content)} bytes)")

            # PPTX -> PDF (COM worker thread)
            future = com_worker.convert(input_path, output_path)
            await asyncio.wait_for(asyncio.wrap_future(future), timeout=CONVERT_TIMEOUT)
            logger.info(f"--end-- [api] push task {input_path} -> {output_path}")

            logger.info(f"--begin-- [api] do ocr for {pdf_name}:{output_path}")
            # PDF -> OCR
            async with httpx.AsyncClient() as client:
                with open(output_path, "rb") as pdf_file:
                    resp = await client.post(
                        OCR_URL,
                        files={"file": (pdf_name, pdf_file, "application/pdf")},
                        timeout=OCR_TIMEOUT,
                    )
            logger.info(f"--end-- [api] do ocr for {pdf_name}:{output_path}")

            if resp.status_code != 200:
                raise HTTPException(status_code=502, detail=f"OCR service error: {resp.text}")

            return resp.json()

        except asyncio.TimeoutError:
            raise HTTPException(status_code=504, detail="conversion timeout")
        except HTTPException:
            raise
        except Exception as e:
            logger.error(f"error: {e}")
            raise HTTPException(status_code=500, detail=str(e))
        finally:
            _cleanup(work_dir)
            _check_memory_and_exit()
    finally:
        print(f"--end-- [api] api_ppt_counter={local_api_ppt_counter}")

def _cleanup(path: str) -> None:
    import shutil
    try:
        shutil.rmtree(path, ignore_errors=True)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=PORT)
