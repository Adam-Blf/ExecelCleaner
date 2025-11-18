# ExcelCleaner

A **Tkinter**-based desktop application for cleaning and normalizing Excel/CSV files. Remove unwanted columns, standardize date formats, and export sanitized data—all through a simple drag-and-drop GUI.

## ✨ Features

- 📂 **Drag-and-Drop Support**: intuitive file loading (requires `tkinterdnd2`)
- 🗑️ **Column Removal**: interactively delete unnecessary columns
- 📅 **Date Normalization**: auto-detect and convert dates to `YYYY-MM-DD` format
- 💾 **Excel Export**: save cleaned data as `*_clean.xlsx`
- 🖥️ **Cross-Platform GUI**: Tkinter interface works on Windows, macOS, and Linux
- 📦 **Standalone Executable**: PyInstaller script to build Windows `.exe`

## 🛠️ Tech Stack

| Component | Technology | Purpose |
|-----------|------------|---------|
| **GUI Framework** | Tkinter | Native Python UI library |
| **Data Processing** | pandas 2.1+ | DataFrame operations and transformations |
| **Excel Engine** | openpyxl 3.1+ | .xlsx file read/write |
| **Drag-and-Drop** | tkinterdnd2 (optional) | Enhanced file selection UX |
| **Packaging** | PyInstaller 6.3+ | Windows executable generation |
| **Language** | Python 3.9+ | Core application logic |

## 📁 Project Structure

```text
ExecelCleaner/
├── main.py                # Main GUI application
├── requirements.txt       # Core dependencies
├── requirements-dev.txt   # Development/packaging dependencies
├── excel_cleaner.spec     # PyInstaller configuration
├── build_windows.bat      # Windows executable build script
└── README.md
```

## 🚀 Quick Start

### Prerequisites

- Python 3.9 or higher
- pip package manager

### Installation

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

# (Optional) Install drag-and-drop support
pip install tkinterdnd2

# Launch application
python main.py
```

### Usage

1. **Load File**
   - Click **"Browse"** button or drag `.xlsx`/`.csv` file into window (if `tkinterdnd2` installed)
   - File preview appears in table

2. **Remove Columns**
   - Select unwanted columns from checklist
   - Click **"Remove Selected Columns"**

3. **Normalize Dates**
   - Application auto-detects date columns
   - Click **"Normalize Dates"** to convert to `YYYY-MM-DD`
   - Manual date column specification available

4. **Export**
   - Click **"Export Clean Excel"**
   - File saves as `original_filename_clean.xlsx` in same directory

## 📋 Supported Formats

### Input

- **Excel**: `.xlsx` (via openpyxl)
- **CSV**: `.csv` (auto-detected encoding)

### Output

- **Excel**: `.xlsx` with cleaned data

### Date Format Detection

Recognizes common patterns:

- `YYYY-MM-DD`, `DD/MM/YYYY`, `MM-DD-YYYY`
- `DD MMM YYYY` (e.g., "15 Jan 2024")
- Timestamps (converted to date-only)

## ⚙️ Configuration

### Customize Date Format Output

Edit `main.py`:

```python
# Change output format (default: YYYY-MM-DD)
df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y')
```

### Adjust Column Removal Behavior

In the GUI layout section:

```python
# Allow multiple column selection
column_listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE)
```

### Set Default Export Path

Modify export logic:

```python
# Export to specific directory
output_path = os.path.join('/path/to/output', f'{base_name}_clean.xlsx')
```

## 🔒 Best Practices

- **Backup Original Files**: app overwrites export if filename exists
- **Date Validation**: review normalized dates; ambiguous formats (e.g., 01/02/03) may parse incorrectly
- **Large Files**: files >50MB may cause GUI lag; consider batch processing or CLI alternative
- **Encoding Issues**: CSV files with non-UTF-8 encoding may fail; pre-convert using `iconv` or similar tools

## 🧪 Testing

### Sample Test Case

Create `test_data.xlsx`:

| Name | Date | Value | Unused Column |
|------|------|-------|---------------|
| Alice | 2024-01-15 | 100 | junk |
| Bob | 15/01/2024 | 200 | more junk |

Expected outcome after cleaning:

- Remove "Unused Column"
- Normalize "Date" to `2024-01-15` format
- Export as `test_data_clean.xlsx`

### Automated Testing

```bash
# (Future enhancement—not yet implemented)
pytest tests/test_cleaner.py
```

## 📦 Building Executable (Windows)

### Prerequisites

```bash
pip install -r requirements-dev.txt
```

### Build Steps

```bash
# Run build script
build_windows.bat
```

Or manually:

```bash
pyinstaller excel_cleaner.spec
```

Output:

- `dist\ExcelCleaner\ExcelCleaner.exe` (main executable)
- `dist\ExcelCleaner\` (folder with all dependencies)

Distribute the entire `dist\ExcelCleaner\` folder to end users.

### Customize Executable

Edit `excel_cleaner.spec`:

```python
# Change app name
exe = EXE(..., name='MyCleanerApp', ...)

# Add icon
exe = EXE(..., icon='path/to/icon.ico', ...)

# Bundle as single file (larger startup time)
exe = EXE(..., onefile=True, ...)
```

## 🗺️ Roadmap

- [ ] **CSV Export**: add option to export as `.csv`
- [ ] **Batch Processing**: clean multiple files in one run
- [ ] **Data Quality Reports**: generate summary of changes (rows removed, columns modified)
- [ ] **Advanced Filters**: remove duplicates, filter by value ranges
- [ ] **Undo/Redo**: revert cleaning operations
- [ ] **CLI Mode**: headless operation for automation/scripts
- [ ] **Cloud Integration**: support Google Sheets or OneDrive uploads
- [ ] **macOS/Linux Packaging**: PyInstaller specs for non-Windows platforms

## 📄 License

This project is licensed under the **MIT License**. See [LICENSE](LICENSE) for details.

---

**Author**: Adam Beloucif  
**Repository**: [github.com/Adam-Blf/ExecelCleaner](https://github.com/Adam-Blf/ExecelCleaner)

For bug reports or feature requests, open an issue on GitHub.
