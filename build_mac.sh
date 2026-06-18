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

# List all web files to include
WEB_FILES=(
    index.html
    style.css
    app.js
    app-api.js
    app-data.js
    app-file.js
    app-ui.js
    app-utils.js
    company_data.json
)

# Build --add-data flags (macOS uses : as separator)
ADD_DATA=""
for f in "${WEB_FILES[@]}"; do
    if [ -f "$f" ]; then
        ADD_DATA="$ADD_DATA --add-data $f:."
    else
        echo "⚠️  Fichier manquant: $f"
    fi
done

# Add logo if present
if [ -f "Liora_Logo_Orange_alpha.png" ]; then
    ADD_DATA="$ADD_DATA --add-data Liora_Logo_Orange_alpha.png:."
fi

pyinstaller \
    --name "Liora-Fournisseur" \
    --onefile \
    --windowed \
    --noconfirm \
    $ICON_FLAG \
    $ADD_DATA \
    --hidden-import=openpyxl \
    --hidden-import=pypdf \
    --hidden-import=reportlab \
    --hidden-import=flask \
    --hidden-import=flask_cors \
    --hidden-import=PIL \
    --collect-submodules=reportlab \
    --collect-submodules=openpyxl \
    server.py

echo ""
echo "✅ Build terminé !"
echo "📁 L'application se trouve dans : dist/Liora-Fournisseur.app"
echo ""
echo "Pour distribuer : copiez dist/Liora-Fournisseur.app sur les postes cibles."
