#!/usr/bin/env bash
# 重新生成课程数据与页面（在 build/ 目录下运行，或直接执行本脚本）
set -euo pipefail
cd "$(dirname "$0")"
echo "== 1/3 合并课程数据 -> courses_merged.json"
python3 build_data.py
echo "== 2/3 生成预设数据 -> ../presets/*.json + manifest.json"
python3 build_presets.py
echo "== 3/3 生成页面 -> ../index.html（工具 + 数据分离，数据在 presets/）"
cp template.html ../index.html
echo "完成"
