#!/bin/bash
# Mở website tạo giấy khen qua máy chủ cục bộ (chỉ chạy trên máy này, không ra Internet)
cd "$(dirname "$0")"
PORT=8000
while lsof -iTCP:$PORT -sTCP:LISTEN >/dev/null 2>&1; do PORT=$((PORT+1)); done
echo "Website đang chạy tại: http://localhost:$PORT"
echo "Đóng cửa sổ này hoặc nhấn Ctrl + C để dừng."
(sleep 1; open "http://localhost:$PORT") &
python3 -m http.server $PORT --bind 127.0.0.1
