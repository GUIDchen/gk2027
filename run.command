#!/bin/bash
cd "$(dirname "$0")"
if [ ! -x ".venv/bin/python" ]; then
  echo "首次使用：正在创建 .venv 并安装依赖..."
  ./setup.sh || exit 1
fi
exec .venv/bin/python gk2027.py "$@"
