# 准备环境

```bash
uv sync
```

# 切换真实OCR服务

## 方法1，使用环境变量

```bash
set OCR_URL=http://localhost:9000/ocr
uv run python service.py
```

## 方法2，改代码

```bash
# find and replace '''OCR_URL = os.environ.get("OCR_URL", "http://localhost:9000/ocr")'''
```

# 测试

## Linux

```bash
# from WSL2 Debian
#URL=http://172.20.144.1:8000/ppt ./batch_process_test.sh  ppt169_禅意风_金刚经第一品研究.pptx
API_URL=http://172.18.208.1:8000/ppt ./batch_process_test.sh  ppt169_chanyifeng_jingangjingdiyipinyanjiu.pptx
```

## Windows

```powershell
# from WSL2 Debian
#URL=http://172.20.144.1:8000/ppt ./batch_process_test.sh  ppt169_禅意风_金刚经第一品研究.pptx
# $env:API_URL = "http://172.18.208.1:8000/ppt"
./batch_process_test.ps1  ppt169_chanyifeng_jingangjingdiyipinyanjiu.pptx
```
