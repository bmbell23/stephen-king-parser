#!/bin/bash

# Exit on any error
set -e

echo "Setting up Python virtual environment..."

# Check if required packages are installed
if ! dpkg -l | grep -q "python3-venv\|python3-full"; then
    echo "Required packages not found. Installing python3-venv and python3-full..."
    sudo apt update
    sudo apt install -y python3-venv python3-full
fi

# Check if Python 3 is installed
if ! command -v python3 &> /dev/null; then
    echo "Python 3 is not installed. Please install it first."
    exit 1
fi

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "Creating new virtual environment..."
    python3 -m venv venv
else
    echo "Virtual environment already exists."
fi

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Verify we're in the virtual environment
if [ -z "$VIRTUAL_ENV" ]; then
    echo "Failed to activate virtual environment."
    exit 1
fi

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip

# Install required packages
echo "Installing required packages..."
# pip install beautifulsoup4  # This includes bs4
sudo apt install python3-bs4
pip install requests
# pip install pandas
sudo apt install python3-pandas
pip install lxml

# Print Python version and installed packages
echo -e "\nPython version:"
python --version
echo -e "\nInstalled packages:"
pip list

echo -e "\nSetup complete! Virtual environment is now active."
echo "To deactivate the virtual environment, run: deactivate"
echo "To activate it again later, run: source venv/bin/activate"
