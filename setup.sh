#!/bin/bash

# Exit on error
set -e

# Project root directory
PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

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

# Create or update the activation script
cat > activate_venv.sh << 'EOF'
#!/bin/bash

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Activate virtual environment
source "${SCRIPT_DIR}/venv/bin/activate"

# Add project root to PYTHONPATH
export PYTHONPATH="${SCRIPT_DIR}/src:${PYTHONPATH}"

# Print status
echo "Virtual environment activated!"
echo "Python version: $(python --version)"
echo "Using pip: $(which pip)"
echo "PYTHONPATH: ${PYTHONPATH}"

# Optional: cd to project root
cd "${SCRIPT_DIR}"
EOF

chmod +x activate_venv.sh

# Activate virtual environment
source venv/bin/activate

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip

# Install package in development mode
echo "Installing package and dependencies..."
pip install -e .

# Install development dependencies
echo "Installing development dependencies..."
pip install pytest pytest-cov black isort mypy

echo -e "\nSetup complete! To activate the virtual environment:"
echo "1. Run: source activate_venv.sh"
echo "2. To run the parser: python src/main.py"
echo "3. To deactivate when done: deactivate"
