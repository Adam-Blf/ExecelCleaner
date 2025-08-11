# Excel Cleaner

Petit outil Python qui permet de nettoyer rapidement un fichier Excel ou CSV.

## Fonctionnalités
- Suppression des colonnes inutiles
- Normalisation des dates en `YYYY-MM-DD` (auto ou manuel)
- Export d’un fichier propre `*_clean.xlsx`
- Interface graphique Tkinter
- Glisser-déposer du fichier (si `tkinterdnd2` est installé)
- Script pour générer un `.exe` Windows avec PyInstaller

## Utilisation
### Lancer l’application
1. Cloner le projet  
2. Installer les dépendances :
   ```bash
   pip install -r requirements.txt
   ```
3. Lancer :
   ```bash
   python main.py
   ```

### Générer un .exe (Windows)
1. Lancer :
   ```bash
   build_windows.bat
   ```
2. L’exécutable sera dans `dist\ExcelCleaner\ExcelCleaner.exe`

## Stack utilisée
- Python
- Pandas
- OpenPyXL
- Tkinter
- (Optionnel) tkinterdnd2 pour le drag-and-drop
- PyInstaller pour la version .exe

## Pourquoi ce projet ?
C’est un petit script que j’ai codé pour m’entraîner à :
- Travailler avec **pandas** pour manipuler des données
- Créer une **UI simple avec Tkinter**
- Gérer la **conversion de fichiers Excel/CSV**
- Préparer une version distribuable (.exe)

## Améliorations possibles
- Sauvegarde aussi en CSV
- Barre de progression
- Rapports sur la qualité des données
- Choix de profils de nettoyage enregistrés

---
Licence : MIT
