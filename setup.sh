#!/bin/bash
set -e
cd "$(dirname "$0")"
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
./.venv/bin/python gk2027.py doctor
