#!/bin/bash
cd "$(dirname "$0")"

# 检查虚拟环境
if [ ! -d "venv" ]; then
  echo "正在创建虚拟环境..."
  python3 -m venv venv
fi

# 激活虚拟环境
source venv/bin/activate

# 安装依赖
echo "正在检查依赖..."
pip install -q -r requirements.txt

# 启动服务
echo ""
echo "=========================================="
echo "  厦门大学论文格式工具 已启动"
echo "  请在浏览器打开: http://localhost:5001"
echo "  按 Ctrl+C 停止服务"
echo "=========================================="
echo ""

open http://localhost:5001
python3 app.py
