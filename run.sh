#!/usr/bin/env bash
cd "$(dirname "$0")"
if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 is not installed."
  exit 1
fi
if [ ! -x .venv/bin/python ]; then
  python3 -m venv .venv
  .venv/bin/python -m pip install -r requirements.txt
fi
echo "Open http://127.0.0.1:8000"
( sleep 1 && (xdg-open "http://127.0.0.1:8000" || open "http://127.0.0.1:8000") >/dev/null 2>&1 ) &
exec .venv/bin/python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
