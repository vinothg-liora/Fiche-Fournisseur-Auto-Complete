@echo off
REM ============================================
REM Build Liora Fiche Fournisseur — Windows EXE
REM ============================================

echo.
echo  Build Liora-Fournisseur.exe pour Windows...
echo.

REM Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  Python n'est pas installe.
    echo  Telechargez-le ici : https://www.python.org/downloads/
    echo  IMPORTANT : cochez "Add Python to PATH" lors de l'installation.
    echo.
    pause
    exit /b 1
)

echo  Installation des dependances...
pip install -r requirements.txt

REM Convert PNG to ICO for Windows icon
echo  Creation de l'icone...
python -c "from PIL import Image; img = Image.open('Liora_Logo_Orange_alpha.png'); img.save('liora.ico', format='ICO', sizes=[(256,256),(128,128),(64,64),(48,48),(32,32),(16,16)])" 2>nul
if exist liora.ico (
    set ICON_FLAG=--icon=liora.ico
) else (
    set ICON_FLAG=
)

echo  Build avec PyInstaller...
pyinstaller ^
    --name "Liora-Fournisseur" ^
    --onefile ^
    --windowed ^
    %ICON_FLAG% ^
    --add-data "index.html;." ^
    --add-data "style.css;." ^
    --add-data "app-utils.js;." ^
    --add-data "app-data.js;." ^
    --add-data "app-api.js;." ^
    --add-data "app-file.js;." ^
    --add-data "app-ui.js;." ^
    --add-data "app.js;." ^
    --add-data "company_data.json;." ^
    --add-data "Liora_Logo_Orange_alpha.png;." ^
    --hidden-import=openpyxl ^
    --hidden-import=pypdf ^
    --hidden-import=reportlab ^
    --hidden-import=flask ^
    --hidden-import=flask_cors ^
    --collect-submodules=reportlab ^
    server.py

echo.
echo  Build termine !
echo  L'application se trouve dans : dist\Liora-Fournisseur.exe
echo.
echo  Pour distribuer : copiez dist\Liora-Fournisseur.exe sur les postes cibles.
echo.
pause
