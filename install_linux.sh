#!/bin/sh
set -eu
cd "$(dirname "$0")"
python3 -m venv .venv
. .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
mkdir -p uploads portal_updates portal_backups
echo "Installation completed. Configure environment variables before start_linux.sh."
