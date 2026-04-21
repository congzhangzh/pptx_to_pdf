import random
import re
from fastapi import FastAPI, UploadFile, File, HTTPException
import os
import tempfile
import requests
import win32com.client
import pythoncom
from fastapi.responses import JSONResponse

app = FastAPI()
# PARSE_API_URL = "http://xxxx/ocr/"
PARSE_API_URL = "http://yyyy/ocr/"
SERVICE = "paddle_ocr"

@app.post("/ppt")
async def convert_ppt_to_pdf(file: UploadFile = File(...)):
    # 检查文件扩展名
    if not file.filename.lower().endswith(('.ppt', '.pptx')):
        raise HTTPException(status_code=400, detail="仅支持ppt,pptx格式")
    
    try:
        # 创建临时目录保存文件
        with tempfile.TemporaryDirectory() as temp_dir:
            # 保存上传的PPT文件
            ppt_path = os.path.join(temp_dir, file.filename)
            with open(ppt_path, "wb") as f:
                f.write(await file.read())

            pdf_path = os.path.join(temp_dir, str(random.randint(1000, 9999)) + ".pdf")
            
            # PPT转PDF
            pythoncom.CoInitialize()
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                
            # 打开PPT文件
            presentation = powerpoint.Presentations.Open(ppt_path)
                
            # 保存为PDF
            presentation.SaveAs(pdf_path, 32)  # 32是PDF格式的枚举值
                
            # 关闭演示文稿和PowerPoint
            presentation.Close()
            powerpoint.Quit()
            pdf_filename = f"{os.path.splitext(file.filename)[0]}.pdf"
            # 上传PDF到解析接口
            with open(pdf_path, 'rb') as f:
                files = {'file': (pdf_filename, f, "application/octet-stream")}
                #data = {"service": service}
                response = requests.post(PARSE_API_URL, files=files)
                
            if response.status_code != 200:
                raise HTTPException(status_code=500, detail=f'解析失败: {response.text}')
                    
            # 返回解析结果
            return JSONResponse(content=response.json())
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        pythoncom.CoUninitialize()

if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=2334)
