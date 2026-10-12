#!/bin/sh
set -eu
cd "$(dirname "$0")"
. .venv/bin/activate
exec python app.py
