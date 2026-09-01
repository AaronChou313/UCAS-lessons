#!/usr/bin/env bash
# 用源 Excel 重新生成选课页面（需在 build/ 目录下运行，或直接执行本脚本）
set -euo pipefail
cd "$(dirname "$0")"
echo "== 1/2 解析课表数据 -> courses_merged.json"
python3 build_data.py
echo "== 2/2 生成页面 -> ../index.html"
python3 inject.py ../index.html
echo "完成"
