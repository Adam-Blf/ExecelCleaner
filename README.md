[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/adambeloucif/) ![Visitor Badge](https://visitor-badge.laobi.icu/badge?page_id=Adam-Blf.ExecelCleaner)


![Dernier commit](https://img.shields.io/github/last-commit/Adam-Blf/ExecelCleaner?style=flat&logo=git&logoColor=white&color=0080ff&label=Dernier%20commit) ![Langage principal](https://img.shields.io/github/languages/top/Adam-Blf/ExecelCleaner?style=flat&logo=git&logoColor=white&color=0080ff&label=Langage%20principal) ![Nombre de langages](https://img.shields.io/github/languages/count/Adam-Blf/ExecelCleaner?style=flat&logo=git&logoColor=white&color=0080ff&label=Nombre%20de%20langages)

### Construit avec les outils et technologies : 

![Python](https://img.shields.io/badge/-Python-0080ff?style=flat) ![Batchfile](https://img.shields.io/badge/-Batchfile-0080ff?style=flat)

🇫🇷 Français | 🇬🇧 Anglais | 🇪🇸 Espagnol | 🇮🇹 Italien | 🇵🇹 Portugais | 🇷🇺 Russe | 🇩🇪 Allemand | 🇹🇷 Turc

# ExcelCleaner - Nettoyeur Excel / Excel Cleaner

[🇫🇷 Version Française](#version-française) | [🇬🇧 English Version](#english-version)

---

## <a name="version-française"></a>🇫🇷 Version Française

Application de bureau **Tkinter** pour nettoyer et normaliser des fichiers Excel/CSV. Supprimez des colonnes indésirables, standardisez les formats de date et exportez des données nettoyées via une interface graphique simple avec glisser-déposer.

### ✨ Fonctionnalités

- 📂 **Support Glisser-Déposer** : chargement de fichiers intuitif (nécessite `tkinterdnd2`)
- 🗑️ **Suppression de Colonnes** : effacement interactif des colonnes inutiles
- 📅 **Normalisation de Dates** : détection et conversion automatiques au format `YYYY-MM-DD`
- 💾 **Export Excel** : sauvegarde des données nettoyées en `*_clean.xlsx`
- 🖥️ **Interface Multiplateforme** : GUI Tkinter fonctionnant sur Windows, macOS et Linux
- 📦 **Exécutable Autonome** : script PyInstaller pour créer un `.exe` Windows

### 🛠️ Stack Technologique

| Composant | Technologie | Objectif |
|-----------|-------------|----------|
| **Framework GUI** | Tkinter | Bibliothèque UI native Python |
| **Traitement Données** | pandas 2.1+ | Opérations DataFrame et transformations |
| **Moteur Excel** | openpyxl 3.1+ | Lecture/écriture fichiers .xlsx |
| **Glisser-Déposer** | tkinterdnd2 (optionnel) | UX sélection fichiers améliorée |
| **Packaging** | PyInstaller 6.3+ | Génération exécutable Windows |
| **Langage** | Python 3.9+ | Logique applicative |

### 📁 Structure du Projet

```
ExecelCleaner/
├── main.py                  # Application GUI principale
├── requirements.txt         # Dépendances principales
├── requirements-dev.txt     # Dépendances développement/packaging
├── scripts/
│   ├── excel_cleaner.spec   # Configuration PyInstaller
│   └── build_windows.bat    # Script construction exécutable Windows
└── README.md
```

### 🚀 Démarrage Rapide

#### Prérequis

- Python 3.9 ou supérieur
- Gestionnaire de paquets pip

#### Installation

```bash
# Clonez ou téléchargez le dépôt
cd ExecelCleaner

# Créez un environnement virtuel (recommandé)
python -m venv .venv

# Activez l'environnement
# Windows PowerShell:
.\.venv\Scripts\Activate.ps1
# Windows CMD:
.venv\Scripts\activate.bat
# macOS/Linux:
source .venv/bin/activate

# Installez les dépendances
pip install -r requirements.txt

# (Optionnel) Support glisser-déposer
pip install tkinterdnd2

# Lancez l'application
python main.py
```

### 🎯 Utilisation

1. **Charger un Fichier** : bouton **"Browse"** ou glisser fichier `.xlsx`/`.csv`
2. **Supprimer Colonnes** : sélectionner dans liste → **"Remove Selected Columns"**
3. **Normaliser Dates** : clic **"Normalize Dates"** (détection automatique)
4. **Exporter** : **"Export Clean Excel"** → sauvegarde `*_clean.xlsx`

### 🗺️ Feuille de Route

- [ ] Export CSV
- [ ] Traitement par lots
- [ ] Rapports qualité données
- [ ] Filtres avancés (doublons, plages valeurs)
- [ ] Undo/Redo
- [ ] Mode CLI pour automatisation
- [ ] Intégration cloud (Google Sheets, OneDrive)
- [ ] Packaging macOS/Linux

---

## <a name="english-version"></a>🇬🇧 English Version

A **Tkinter**-based desktop application for cleaning and normalizing Excel/CSV files. Remove unwanted columns, standardize date formats, and export sanitized data through a simple drag-and-drop GUI.

### ✨ Features

- 📂 **Drag-and-Drop Support**: intuitive file loading (requires `tkinterdnd2`)
- 🗑️ **Column Removal**: interactively delete unnecessary columns
- 📅 **Date Normalization**: auto-detect and convert dates to `YYYY-MM-DD` format
- 💾 **Excel Export**: save cleaned data as `*_clean.xlsx`
- 🖥️ **Cross-Platform GUI**: Tkinter interface works on Windows, macOS, and Linux
- 📦 **Standalone Executable**: PyInstaller script to build Windows `.exe`

### 🛠️ Tech Stack

| Component | Technology | Purpose |
|-----------|------------|---------|
| **GUI Framework** | Tkinter | Native Python UI library |
| **Data Processing** | pandas 2.1+ | DataFrame operations and transformations |
| **Excel Engine** | openpyxl 3.1+ | .xlsx file read/write |
| **Drag-and-Drop** | tkinterdnd2 (optional) | Enhanced file selection UX |
| **Packaging** | PyInstaller 6.3+ | Windows executable generation |
| **Language** | Python 3.9+ | Core application logic |

### 📁 Project Structure

```
ExecelCleaner/
├── main.py                  # Main GUI application
├── requirements.txt         # Core dependencies
├── requirements-dev.txt     # Development/packaging dependencies
├── scripts/
│   ├── excel_cleaner.spec   # PyInstaller configuration
│   └── build_windows.bat    # Windows executable build script
└── README.md
```

### 🚀 Quick Start

#### Prerequisites

- Python 3.9 or higher
- pip package manager

#### Installation

```bash
# Clone or download the repository
cd ExecelCleaner

# Create virtual environment (recommended)
python -m venv .venv

# Activate environment
# Windows PowerShell:
.\.venv\Scripts\Activate.ps1
# Windows CMD:
.venv\Scripts\activate.bat
# macOS/Linux:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# (Optional) Drag-and-drop support
pip install tkinterdnd2

# Launch application
python main.py
```

### 🎯 Usage

1. **Load File**: click **"Browse"** or drag `.xlsx`/`.csv` file
2. **Remove Columns**: select from checklist → **"Remove Selected Columns"**
3. **Normalize Dates**: click **"Normalize Dates"** (auto-detection)
4. **Export**: **"Export Clean Excel"** → saves `*_clean.xlsx`

### 🗺️ Roadmap

- [ ] CSV export
- [ ] Batch processing
- [ ] Data quality reports
- [ ] Advanced filters (duplicates, value ranges)
- [ ] Undo/Redo
- [ ] CLI mode for automation
- [ ] Cloud integration (Google Sheets, OneDrive)
- [ ] macOS/Linux packaging

### 📄 License

This project is open source. See LICENSE file for details.

---

**Author**: Adam Beloucif  
**Repository**: [github.com/Adam-Blf/ExecelCleaner](https://github.com/Adam-Blf/ExecelCleaner)

For bug reports or feature requests, open an issue on GitHub.

---

<p align="center">
  <sub>Par <a href="https://adam.beloucif.com">Adam Beloucif</a> · Data Engineer & Fullstack Developer · <a href="https://github.com/Adam-Blf">GitHub</a> · <a href="https://www.linkedin.com/in/adambeloucif/">LinkedIn</a></sub>
</p>
