#!/usr/bin/env bash
set -e

# Load .env if it exists
if [ -f .env ]; then
  export $(grep -v '^#' .env | xargs)
fi

# Install deps if needed
if ! python3 -c "import flask" 2>/dev/null; then
  pip install -r requirements.txt -q
fi

echo "Starting VideoGen AI on http://localhost:5000"
python3 app.py
