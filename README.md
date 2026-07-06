# Excel Merger Tool

A **Python utility** that merges dozens of per-part Excel files into a single consolidated workbook — extracting only the data block between two text markers, tagging each row with Part No. / OP / source filename, and automatically detecting and excluding any row that would corrupt the output file.

Built for shop-floor / CAM documentation workflows where each part or operation lives in its own Excel file and needs to be rolled up into one master sheet.

---

## 📌 Features

- **Merge many Excel files into one** — scans a source folder and combines every file into a single output workbook.
- **Marker-based extraction** — reads only the rows between a start marker (`APPENDIX_HEADER`, e.g. `"Tool No."`) and a stop marker (`STOP_MARKER`, e.g. `"CAM Programmer :"`), so headers/footers/signatures outside that block are ignored.
- **Filename-derived metadata** — parses each filename as `<OP><separator><PART_NUMBER>` (e.g. `OP10----ABC123`) and writes Part No., OP, and the source filename onto every merged row.
- **Real corruption detection, not guesswork** — every collected row is validated with an actual save + reload round-trip of the output workbook. If that fails, the tool binary-searches the row set to identify the exact offending file and source row, drops only that row, and continues.
- **Cell-level sanitization** — strips illegal XML control characters, truncates strings over Excel's 32,767-character cell limit, and converts NaN/Infinity floats to blank — the most common causes of "we found a problem with some content" repair prompts.
- **.xlsx / .xlsm / .xls support** — legacy `.xls` files are read via `xlrd`; `.xlsx`/`.xlsm` via `openpyxl`.
- **Safe to re-run** — automatically skips Excel lock files (`~$...`) and its own previous output(s), and if the output file already exists it saves as `..._1.xlsx`, `..._2.xlsx`, etc. instead of overwriting.
- **Flexible column placement** — write Part No./OP/Filename into fixed columns (shifting existing data right) or simply append them after each row's existing data.
- **Detailed console reporting** — per-file progress, rows collected, and a final summary listing any files skipped and any corrupt rows excluded (with file name + source row number).

---

## 🛠️ Installation

### Prerequisites

- Python 3.8+
- `pip`

### Required libraries

```bash
pip install openpyxl xlrd
```

---

## 🚀 Usage

### 1. Configure

Open `excel_merger.py` and edit the `CONFIG` block near the top of the file:

| Parameter             | Description                                                                 | Default                                                    |
| ---------------------- | ---------------------------------------------------------------------------- | ----------------------------------------------------------- |
| `SOURCE_DIR`           | Folder containing the source Excel files                                    | `Path(r"C:\temp\EXCEL_SOURCE")` |
| `OUTPUT_FILE`          | Full path (incl. filename) for the merged workbook                          | `Path(r"C:\temp\EXCEL_MERGED.xlsx")` |
| `SHEET_SELECTOR`       | Sheet to read from each source file — `int` (0-based index) or `str` (sheet name) | `0` |
| `APPENDIX_HEADER`      | Text marker where the data block starts (inclusive)                         | `"Tool No."` |
| `STOP_MARKER`          | Text marker where the data block ends (exclusive). If not found, reads to the last row. | `"CAM Programmer :"` |
| `START_COL_PART_OP`    | 1-based column where Part No./OP/Filename are written. `None` = append after each row's existing data (safest — avoids overwriting variable-width rows). An `int` (e.g. `18`) inserts them at a fixed column, shifting existing data right. | `18` (column R) |
| `FILENAME_SEPARATOR`   | Separator used in filenames, parsed as `<OP><separator><PART_NUMBER>`       | `"----"` |
| `MAX_COLUMNS_TO_SCAN`  | Hard cap on columns read per row — guards against bloated files where `max_column` reports Excel's absolute limit (16384) instead of the real ~10–20 columns of data | `200` |
| `SKIP_PREFIXES`        | Filename-stem prefixes to always skip. Defaults to Excel's own lock-file prefix (`~$`) plus the output file's own name, so re-running never re-ingests a previous merge. | `("~$", OUTPUT_FILE.stem)` |

> Text markers (`APPENDIX_HEADER`, `STOP_MARKER`) are matched case-insensitively and space-insensitively, so `"Tool No."`, `"tool no."`, and `"ToolNo."` all match.

### 2. Run

```bash
python excel_merger.py
```

### 3. Output

For each file, the console prints which sheet was used, the detected start/stop rows, and how many rows were collected. Once all files are read, the tool validates the full merged data set and prints a summary:

```
============================================================
DONE! Files merged: 42 | Files skipped: 1
Corrupt rows excluded from output (1):
 - OP20----XYZ789.xlsx (source row 14) [append() raised: ...]
Output: C:\...\AL10_PARTS_EXCEL_MERGED.xlsx
```

If `OUTPUT_FILE` already exists, the tool automatically writes to `..._1.xlsx` (or the next free number) instead of overwriting it.

---

## 📂 Example Directory Structure

```
project_root/
├── excel_merger.py
├── README.md
└── AL10_PARTS_EXCEL/
    ├── OP10----ABC123.xlsx
    ├── OP20----ABC123.xlsx
    └── OP10----XYZ789.xlsx
```

Filenames are expected to follow `<OP><FILENAME_SEPARATOR><PART_NUMBER>` — e.g. `OP10----ABC123.xlsx` → OP = `OP10`, Part No. = `ABC123`. Files that don't match the pattern are still merged, but their Part No./OP columns are left blank (a warning is printed).

---

## 🔧 Customization

**Change the source/output paths:**
```python
SOURCE_DIR = Path(r"path/to/your/source/directory")
OUTPUT_FILE = Path(r"path/to/your/output/file.xlsx")
```

**Skip specific files:** add stems/prefixes to `SKIP_PREFIXES`:
```python
SKIP_PREFIXES = ("~$", OUTPUT_FILE.stem, "temp_", "backup_")
```

**Append vs. fixed columns for metadata:**
```python
START_COL_PART_OP = None   # append after each row's own data (default-safe)
START_COL_PART_OP = 10     # force into column J, shifting existing data right
```

**Different filename delimiter:**
```python
FILENAME_SEPARATOR = "_"   # e.g. "OP10_ABC123.xlsx"
```

---

## 🧪 How corruption detection works

Rather than guessing which values are "risky," the tool builds a throwaway workbook from all collected rows and does a real `save()` → `load_workbook()` round-trip. If that fails:

1. Known-good rows are set aside.
2. The remaining candidate rows are binary-searched (capped at `MAX_BISECTION_ROUNDS = 50` rounds) to find the exact first row that breaks the round-trip.
3. That row is dropped and reported (source file + row number), and the search continues on the rest.

Before any of this, every cell value is also passed through a sanitizer that strips illegal XML control characters, truncates strings past Excel's 32,767-character cell limit, and converts NaN/Infinity floats to blank — which catches the majority of would-be corruption before it ever reaches the round-trip check.

---

## 🧪 Testing

1. **Sample run** — point `SOURCE_DIR` at a small folder of test files and confirm the merged output and console summary look right.
2. **Edge cases** — try files with mismatched filename patterns, very wide/malformed sheets, and at least one intentionally corrupt cell to confirm it's detected and reported rather than silently breaking the output.

---

## 📝 License

MIT License — see [LICENSE](LICENSE).

---

## 🤝 Contributing

1. Fork the repository.
2. Create a branch: `git checkout -b feature/your-feature`
3. Commit: `git commit -am 'Add some feature'`
4. Push: `git push origin feature/your-feature`
5. Open a Pull Request.

---

## 🙏 Acknowledgments

- [openpyxl](https://openpyxl.readthedocs.io/) — reading/writing `.xlsx`/`.xlsm`.
- [xlrd](https://xlrd.readthedocs.io/) — reading legacy `.xls` files.
