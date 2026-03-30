#!/bin/bash
# ============================================
# Build Liora Fiche Fournisseur — macOS App
# ============================================
set -e

echo "🔧 Build Liora-Fournisseur.app pour macOS..."
echo ""

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 non trouvé. Installez-le depuis https://www.python.org/downloads/"
    exit 1
fi

echo "📦 Installation des dépendances..."
pip3 install -r requirements.txt

# Convert PNG to ICNS for macOS icon
echo "🎨 Création de l'icône..."
if [ -f "Liora_Logo_Orange_alpha.png" ]; then
    mkdir -p icon.iconset
    sips -z 16 16     Liora_Logo_Orange_alpha.png --out icon.iconset/icon_16x16.png 2>/dev/null || true
    sips -z 32 32     Liora_Logo_Orange_alpha.png --out icon.iconset/icon_16x16@2x.png 2>/dev/null || true
    sips -z 32 32     Liora_Logo_Orange_alpha.png --out icon.iconset/icon_32x32.png 2>/dev/null || true
    sips -z 64 64     Liora_Logo_Orange_alpha.png --out icon.iconset/icon_32x32@2x.png 2>/dev/null || true
    sips -z 128 128   Liora_Logo_Orange_alpha.png --out icon.iconset/icon_128x128.png 2>/dev/null || true
    sips -z 256 256   Liora_Logo_Orange_alpha.png --out icon.iconset/icon_128x128@2x.png 2>/dev/null || true
    sips -z 256 256   Liora_Logo_Orange_alpha.png --out icon.iconset/icon_256x256.png 2>/dev/null || true
    sips -z 512 512   Liora_Logo_Orange_alpha.png --out icon.iconset/icon_256x256@2x.png 2>/dev/null || true
    sips -z 512 512   Liora_Logo_Orange_alpha.png --out icon.iconset/icon_512x512.png 2>/dev/null || true
    sips -z 1024 1024 Liora_Logo_Orange_alpha.png --out icon.iconset/icon_512x512@2x.png 2>/dev/null || true
    iconutil -c icns icon.iconset -o liora.icns 2>/dev/null || true
    rm -rf icon.iconset
    ICON_FLAG="--icon=liora.icns"
else
    ICON_FLAG=""
fi

echo "🏗️ Build avec PyInstaller..."
pyinstaller \
    --name "Liora-Fournisseur" \
    --onefile \
    --windowed \
    $ICON_FLAG \
    --add-data "index.html:." \
    --add-data "style.css:." \
    --add-data "app-utils.js:." \
    --add-data "app-data.js:." \
    --add-data "app-api.js:." \
    --add-data "app-file.js:." \
    --add-data "app-ui.js:." \
    --add-data "app.js:." \
    --add-data "company_data.json:." \
    --add-data "Liora_Logo_Orange_alpha.png:." \
    --hidden-import=openpyxl \
    --hidden-import=pypdf \
    --hidden-import=reportlab \
    --hidden-import=flask \
    --hidden-import=flask_cors \
    --collect-submodules=reportlab \
    server.py

echo ""
echo "✅ Build terminé !"
echo "📁 L'application se trouve dans : dist/Liora-Fournisseur.app"
echo ""
echo "Pour distribuer : copiez dist/Liora-Fournisseur.app sur les postes cibles."
