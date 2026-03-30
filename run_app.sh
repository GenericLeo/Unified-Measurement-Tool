#!/bin/bash
# Launcher script for Unified Measurement Tool

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Navigate to the script directory
cd "$SCRIPT_DIR"

# Check if virtual environment exists
if [ ! -d ".venv" ]; then
    echo "Virtual environment not found. Creating one..."
    python3 -m venv .venv
    
    echo "Installing dependencies..."
    source .venv/bin/activate
    pip install -r requirements.txt
else
    # Activate virtual environment
    source .venv/bin/activate
fi

# Run the application
echo "Starting Unified Measurement Tool..."
python main_app.py

# Deactivate virtual environment when done
deactivate
