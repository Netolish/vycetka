#!/usr/bin/env bash

set -e
INST_DIR="$HOME/.config/libreoffice/4/user/Scripts/python"

if [ ! -d "$INST_DIR" ]; then
	mkdir -p "$INST_DIR"
fi

cp src/Vycetka.py "$INST_DIR"
