#!/bin/zsh
cd "/Users/shoufu/Documents/小狗/投资追踪" || exit 1
if /usr/bin/curl -s -o /dev/null --max-time 2 http://127.0.0.1:5050/; then
  echo "服务已在运行，无需重复启动。"
  exit 0
fi
while true; do
  /Library/Developer/CommandLineTools/Library/Frameworks/Python3.framework/Versions/3.9/Resources/Python.app/Contents/MacOS/Python server.py
  echo "服务退出，5 秒后自动重启…"
  sleep 5
done
