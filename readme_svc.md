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
