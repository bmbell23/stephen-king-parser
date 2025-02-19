#!/bin/bash

# Exit on error
set -e

# Check if required packages are installed
if ! dpkg -l | grep -q "python3-venv\|python3-full"; then
    echo "Required packages not found. Installing python3-venv and python3-full..."
    sudo apt update
    sudo apt install -y python3-venv python3-full
fi

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
source venv/bin/activate

# Upgrade pip
pip install --upgrade pip

# Install package in development mode
echo "Installing package..."
pip install -e .

echo "Setup complete! You can now run:"
echo "stephen-king-parser --output ."
echo "or"
echo "python -m stephen_king_parser --output ."