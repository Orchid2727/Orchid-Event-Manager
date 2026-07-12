#!/bin/bash
cd "$(dirname "$0")"
PYTHON="/Library/Frameworks/Python.framework/Versions/3.14/bin/python3"
if [ ! -x "$PYTHON" ]; then
  PYTHON="python3"
fi
"$PYTHON" app.py
