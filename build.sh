#!/usr/bin/env bash
# exit on error
set -o errexit

# Install system dependencies
apt-get update && apt-get install -y libjpeg-dev zlib1g-dev --no-install-recommends

# Install Python dependencies
pip install -r requirements.txt