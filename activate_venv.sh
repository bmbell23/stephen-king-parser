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
