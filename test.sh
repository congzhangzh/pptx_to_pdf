#!/bin/bash
# PPTX to PDF 转换服务 — 测试脚本
# 用法: bash test.sh [pptx文件路径]

BASE_URL="http://localhost:8000"

echo "=== 健康检查 ==="
curl -s "$BASE_URL/health"
echo ""
echo ""

if [ -z "$1" ]; then
    echo "用法: bash test.sh <pptx文件路径>"
    echo ""
    echo "示例:"
    echo "  bash test.sh 演示文稿.pptx"
    exit 0
fi

INPUT="$1"
OUTPUT="${INPUT%.*}.pdf"

echo "=== 转换测试 ==="
echo "输入: $INPUT"
echo "输出: $OUTPUT"
echo ""

curl -X POST "$BASE_URL/convert" \
     -F "file=@$INPUT" \
     -o "$OUTPUT" \
     -w "HTTP状态码: %{http_code}\n文件大小: %{size_download} bytes\n耗时: %{time_total}s\n" \
     --progress-bar

echo ""
if [ -f "$OUTPUT" ] && [ -s "$OUTPUT" ]; then
    echo "✅ 转换成功！已保存: $OUTPUT"
else
    echo "❌ 转换失败"
fi
