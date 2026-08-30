#!/bin/bash
# Build DocMerge.app - a real Mac application with a Dock icon.
# Double-click to run. Takes 2-3 minutes the first time.
set -e
cd "$(dirname "$0")"

echo "=============================================="
echo " DocMerge 3.1 - Build Mac App"
echo "=============================================="

# 1. Python check
if ! command -v python3 &> /dev/null; then
  echo "Python 3 not found. Install via: xcode-select --install"
  read -p "Press Enter to close."; exit 1
fi

# 2. Virtual environment + dependencies
echo ""
echo "[1/4] Installing dependencies..."
PYVER=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
echo "  Using Python $PYVER"
python3 -m venv .venv
source .venv/bin/activate
pip install --quiet --upgrade pip
pip install --quiet python-docx pypdf pyinstaller

# 3. Convert icon.png to icon.icns using built-in macOS tools
echo "[2/4] Building app icon..."
if [ -f icon.png ]; then
  rm -rf icon.iconset icon.icns
  mkdir icon.iconset
  for size in 16 32 64 128 256 512; do
    sips -z $size $size icon.png --out icon.iconset/icon_${size}x${size}.png >/dev/null
    dbl=$((size * 2))
    sips -z $dbl $dbl icon.png --out icon.iconset/icon_${size}x${size}@2x.png >/dev/null
  done
  iconutil -c icns icon.iconset -o icon.icns
  rm -rf icon.iconset
  ICON_FLAG="--icon icon.icns"
else
  echo "  (icon.png missing, building without custom icon)"
  ICON_FLAG=""
fi

# 4. Build the .app with PyInstaller
echo "[3/4] Building DocMerge.app (this is the slow bit)..."
rm -rf build dist DocMerge.spec
pyinstaller --windowed --noconfirm --name DocMerge $ICON_FLAG \
  --osx-bundle-identifier com.stevederviniotis.docmerge \
  docmerge.py >/dev/null 2>build_log.txt || {
    echo "Build failed. See build_log.txt for details."
    read -p "Press Enter to close."; exit 1
  }

# 5. Install to Applications
echo "[4/4] Installing..."
if [ -d "/Applications/DocMerge.app" ]; then
  rm -rf "/Applications/DocMerge.app"
fi
cp -R "dist/DocMerge.app" /Applications/
rm -rf build dist DocMerge.spec build_log.txt

echo ""
echo "=============================================="
echo " Done. DocMerge is now in your Applications"
echo " folder. Open it from Launchpad or Spotlight,"
echo " and drag it to your Dock to keep it there."
echo ""
echo " Because it was built on this Mac, Gatekeeper"
echo " will not block it."
echo "=============================================="
read -p "Press Enter to close."
