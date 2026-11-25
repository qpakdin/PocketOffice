#!/bin/sh
# Installs everything the converter + editors need on Debian/Ubuntu (also inside Termux's Debian proot).
# Usage:  sh install_deps.sh
set -e
if [ "$(id -u)" != "0" ] && command -v sudo >/dev/null 2>&1; then SUDO=sudo; else SUDO=; fi
$SUDO apt update
$SUDO apt install -y python3 python3-lxml poppler-utils \
  libreoffice-writer libreoffice-calc libreoffice-impress \
  fonts-crosextra-carlito fonts-crosextra-caladea fonts-liberation fonts-noto-core
echo
python3 "$(dirname "$0")/po_doctor.py"
