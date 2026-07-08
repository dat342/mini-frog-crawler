#!/bin/bash
# Nhấp đúp file này trong Finder để chạy tool (Mac)
# Lần đầu chạy sẽ tự cài môi trường, các lần sau mở thẳng.
cd "$(dirname "$0")"

if [ ! -x ".venv/bin/python" ]; then
  echo "🔧 Lần đầu chạy: đang cài đặt môi trường (1-2 phút, cần mạng)..."
  python3 -m venv .venv
  if [ $? -ne 0 ]; then
    echo "❌ Không tìm thấy Python 3. macOS sẽ tự đề nghị cài Command Line Tools — bấm Install rồi chạy lại file này."
    read -n 1 -s -p "Nhấn phím bất kỳ để đóng..."
    exit 1
  fi
  .venv/bin/pip install -q --upgrade pip
  .venv/bin/pip install -q -r requirements.txt
  if [ $? -ne 0 ]; then
    echo "❌ Cài thư viện thất bại. Kiểm tra kết nối mạng rồi chạy lại."
    read -n 1 -s -p "Nhấn phím bất kỳ để đóng..."
    exit 1
  fi
  echo "✓ Cài đặt xong!"
fi

echo "Đang khởi động Mini Frog..."
.venv/bin/python -m uvicorn app:app --port 8765 &
SERVER_PID=$!
sleep 2
open http://localhost:8765
echo ""
echo "✓ Tool đang chạy tại http://localhost:8765"
echo "  Đóng cửa sổ Terminal này để tắt tool."
wait $SERVER_PID
