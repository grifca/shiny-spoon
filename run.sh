#!/bin/bash
# Campaign Performance Reporter — Launcher
# Usage: ./run.sh

cd "$(dirname "$0")"

# Install deps if needed
if ! python3 -c "import streamlit" 2>/dev/null; then
  echo "Installing dependencies..."
  pip3 install -r requirements.txt
fi

echo "Starting Campaign Performance Reporter..."
echo "Open: http://localhost:8501"
python3 -m streamlit run app.py --server.port 8501 --browser.gatherUsageStats false
