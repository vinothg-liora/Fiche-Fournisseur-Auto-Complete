# Liora — Fiche Fournisseur Auto-Complete

## Comment générer l'exécutable

### Prérequis
- **Python 3.10+** installé ([python.org/downloads](https://www.python.org/downloads/))
- Sur Windows : cocher **"Add Python to PATH"** lors de l'installation

### Build macOS
```bash
chmod +x build_mac.sh
./build_mac.sh
```
Résultat : `dist/Liora-Fournisseur.app`

### Build Windows
```
build_windows.bat
```
(Double-clic sur le fichier)

Résultat : `dist\Liora-Fournisseur.exe`

## Distribution

Copiez simplement le fichier `.app` (macOS) ou `.exe` (Windows) sur les postes de l'équipe ADV. **Aucune installation requise** pour l'utilisateur final.

## Utilisation

1. Double-cliquez sur **Liora-Fournisseur**
2. Le navigateur s'ouvre automatiquement
3. Configurez votre clé API Anthropic dans **Paramètres** (une seule fois)
4. Glissez-déposez votre fichier client (Excel ou PDF)
5. Vérifiez et complétez les champs
6. Téléchargez le document complété

## Architecture

```
Utilisateur → double-clic sur l'app
                ↓
        Serveur Flask local (port 5000)
                ↓
        index.html s'ouvre dans le navigateur
                ↓
        Upload fichier → API Claude analyse les champs
                ↓
        Vérification/validation par l'utilisateur
                ↓
        Serveur Flask complète le fichier avec openpyxl
                ↓
        Téléchargement du fichier complété (formatage intact)
```
