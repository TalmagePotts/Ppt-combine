#!/bin/bash

# PowerPoint Combiner - Mac Launcher
# This script sets up the environment and launches the GUI

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Name of virtual environment
VENV_DIR="venv"

# Check if virtual environment exists, if not create it
if [ ! -d "$VENV_DIR" ]; then
    echo "Creating virtual environment..."
    python3 -m venv "$VENV_DIR"

    if [ $? -ne 0 ]; then
        osascript -e 'display alert "Error" message "Failed to create virtual environment. Please ensure Python 3 is installed."'
        exit 1
    fi
fi

# Activate virtual environment
source "$VENV_DIR/bin/activate"

# Install/update requirements
echo "Checking dependencies..."
pip install -q -r requirements.txt

if [ $? -ne 0 ]; then
    osascript -e 'display alert "Error" message "Failed to install dependencies."'
    exit 1
fi

# Launch the GUI
echo "Launching PowerPoint Combiner..."
python combine_powerpoints_gui.py

# Deactivate virtual environment when done
deactivate
