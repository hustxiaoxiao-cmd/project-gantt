#!/bin/bash
# 启动本地 HTTP 服务器，用于甘特图数据驱动版
# 使用方法: cd project-gantt && ./serve.sh
cd "$(dirname "$0")"
PORT=${1:-8080}
echo "🚀 甘特图服务已启动"
echo "📂 数据目录: $(pwd)/data/"
echo "🌐 浏览器打开: http://localhost:$PORT"
echo ""
echo "修改 data/ 下的 JSON 文件后，刷新浏览器即可看到变化"
echo "按 Ctrl+C 停止服务"
echo ""
python3 -m http.server $PORT
