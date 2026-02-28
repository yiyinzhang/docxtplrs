#!/bin/bash
# Build script for docxtplrs

set -e

echo "Building docxtplrs..."

# Check if uv is available
if command -v uv &> /dev/null; then
    echo "Using uv for environment management"
    
    # Create virtual environment if it doesn't exist
    if [ ! -d ".venv" ]; then
        uv venv
    fi
    
    # Activate virtual environment
    source .venv/bin/activate
    
    # Install maturin if needed
    if ! command -v maturin &> /dev/null; then
        uv pip install maturin
    fi
    
    # Build and install
    maturin develop --uv
else
    echo "Using pip for environment management"
    
    # Create virtual environment if it doesn't exist
    if [ ! -d ".venv" ]; then
        python3 -m venv .venv
    fi
    
    # Activate virtual environment
    source .venv/bin/activate
    
    # Install maturin if needed
    pip install maturin
    
    # Build and install
    maturin develop
fi

echo "Build complete!"
echo ""
echo "Run tests with: python3 -m pytest tests/"
echo "Run basic example with: python3 examples/basic_example.py"
