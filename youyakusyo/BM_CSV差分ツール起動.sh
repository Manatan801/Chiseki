#!/usr/bin/env bash
cd "$(dirname "$0")" || exit 1
python3 diff_bm_csv_web.py
