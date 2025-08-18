#!/usr/bin/env bash
# exit on error
set -o errexit

# --- Add this section to install system dependencies ---
echo "Installing system dependencies for image processing..."
apt-get update && apt-get install -y libjpeg-dev zlib1g-dev --no-install-recommends

# --- Your original build command ---
echo "Installing Python dependencies..."
pip install -r requirements.txt

echo "Build complete."
