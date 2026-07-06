# Excel Merger Tool

A **Python-based utility** for merging multiple Excel files into a single, consolidated output file. This tool is designed to handle large datasets, detect and exclude corrupt rows, and preserve data integrity during the merge process.

---

## 📌 Features

- **Merge Multiple Excel Files**: Combines data from all Excel files in a specified directory into a single output file.
- **Customizable Data Extraction**: Supports reading from a specific sheet (by index or name) and extracting data between defined markers.
- **Corrupt Row Detection**: Automatically identifies and excludes corrupt rows to ensure the output file remains error-free.
- **Flexible Column Handling**: Allows specifying a fixed column for writing additional metadata (e.g., Part No., OP, Filename) or appending it dynamically.
- **Configurable Filename Parsing**: Supports custom separators for parsing filenames into meaningful components.
- **Progress Tracking**: Provides real-time feedback on the merge process, including skipped files, merged files, and corrupt rows.

---

## 🛠️ Installation

### Prerequisites

- Python 3.8 or higher
- `pip` (Python package manager)

### Required Libraries

Install the required dependencies using the following command:

```bash
pip install openpyxl xlrd
```

---

## 🚀 Usage

### 1. **Configure the Tool**

Edit the `CONFIG` section in `excel_merger.py` to customize the following parameters:


| Parameter             | Description                                                       | Default Value                                                                           |
| --------------------- | ----------------------------------------------------------------- | --------------------------------------------------------------------------------------- |
| `SOURCE_DIR`          | Directory containing the source Excel files                       | `Path(r"C:\temp\EXCEL_SOURCE")`                                                         |
| `OUTPUT_FILE`         | Path and filename for the merged output file                      | `Path(r"C:\temp\EXCEL_MERGED.xlsx")`                                                    |
| `SHEET_SELECTOR`      | Sheet to read from each source file (0-based index or sheet name) | `0`                                                                                     |
| `APPENDIX_HEADER`     | Text marker identifying the start of data rows                    | `"Tool No."`                                                                            |
| `STOP_MARKER`         | Text marker identifying the end of data rows                      | `"CAM Programmer :"`                                                                    |
| `START_COL_PART_OP`   | Column (1-based) for writing Part No., OP, and Filename           | `18` (Column R)                                                                         |
| `FILENAME_SEPARATOR`  | Separator used in filenames (e.g., `OP10----ABC123`)              | `"----"`                                                                                |
| `SKIP_PREFIXES`       | Filename prefixes to skip during merging                          | `[]`                                                                                    |
| `MAX_COLUMNS_TO_SCAN` | Maximum columns to scan per row                                   | `50`                                                                                    |


### 2. **Run the Tool**

Execute the script from the command line:

```bash
python excel_merger.py
```

### 3. **Output**

- The merged file will be saved to the specified `OUTPUT_FILE` path.
- The console will display:
  - Number of files merged and skipped.
  - Details of any corrupt rows excluded from the output.

---

## 📂 Example Directory Structure

```
project_root/
├── excel_merger.py
├── README.md
└── source_excel_files/
    ├── file1.xlsx
    ├── file2.xlsx
    └── file3.xlsx
```

---

## 🔧 Customization

### Changing the Source Directory

Update the `SOURCE_DIR` variable in the `CONFIG` section to point to your directory:

```python
SOURCE_DIR = Path(r"path/to/your/source/directory")
```

### Changing the Output File

Update the `OUTPUT_FILE` variable:

```python
OUTPUT_FILE = Path(r"path/to/your/output/file.xlsx")
```

### Skipping Specific Files

Add prefixes to the `SKIP_PREFIXES` list to exclude files:

```python
SKIP_PREFIXES = ["temp_", "backup_"]
```

### Adjusting Column Handling

Set `START_COL_PART_OP` to `None` to append metadata dynamically or specify a fixed column:

```python
START_COL_PART_OP = None  # Dynamic appending
# OR
START_COL_PART_OP = 10  # Fixed column (e.g., Column J)
```

---

## 🧪 Testing

1. **Test with Sample Data**:
  - Create a test directory with a few Excel files.
  - Run the script and verify the output file.
2. **Edge Cases**:
  - Test with files containing corrupt rows.
  - Test with varying column widths and data formats.

---

## 📝 License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

---

## 🤝 Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature/your-feature`).
3. Commit your changes (`git commit -am 'Add some feature'`).
4. Push to the branch (`git push origin feature/your-feature`).
5. Open a Pull Request.

---

## 📧 Contact

For questions or feedback, please contact:

- **Author**: [Your Name](https://github.com/your-username)
- **Email**: [your.email@example.com](mailto:your.email@example.com)

---

## 📜 Changelog

### v1.0.0 (2026-07-06)

- Initial release of the Excel Merger Tool.
- Added support for merging Excel files with corrupt row detection.
- Added configurable parameters for data extraction and output.

---

## 🙏 Acknowledgments

- [openpyxl](https://openpyxl.readthedocs.io/) for Excel file handling.
- [xlrd](https://xlrd.readthedocs.io/) for reading older Excel file formats.
