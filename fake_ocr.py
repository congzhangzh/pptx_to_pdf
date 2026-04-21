"""
Fake OCR stub for local development.
Returns a fixed response so the full PPTX -> PDF -> OCR pipeline
can be exercised without a real OCR engine.

Usage:
    uv run python fake_ocr.py
"""

import logging

from itertools import count

from fastapi import FastAPI, UploadFile
from fastapi.responses import JSONResponse

logging.basicConfig(
    level=logging.INFO, 
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("fake_ocr.log", encoding="utf-8"),
        # logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = FastAPI(title="Fake OCR Service")


api_ocr_counter_gen = count(1)
@app.post("/ocr")
async def ocr(file: UploadFile) -> JSONResponse:
    counter = next(api_ocr_counter_gen)
    try:
        logger.info(f"--begin-- request {counter}")
        print(f"--begin-- request {counter} — received file: {file.filename}")
        content = await file.read()
        logger.info(f"received file: {file.filename} ({len(content)} bytes) — returning stub response")
        return JSONResponse(
            content={
                "status": "ok",
                "filename": file.filename,
                "pages": 1,
                "text": "fake ocr text — replace with real OCR engine",
            }
        )
    finally:
        logger.info(f"--end-- request {counter}")
        print(f"--end-- request {counter} -- processed file: {file.filename}")


@app.get("/health")
async def health():
    return {"status": "ok"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=9000)
