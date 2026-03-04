#!/usr/bin/env bash
# Exit on error
set -o errexit

# Linux dependencies for converting DOCX to PDF
apt-get update && apt-get install -y libreoffice

# Python dependencies
pip install -r requirements.txt
