# Excel Merge Tool 📊⇨📈

An automated tool for merging multiple Excel files, specifically designed for tables with identical headers. Supports both .xlsx and .xls formats.

## 🚀 Core Features

- **Smart Merging**: Automatically identifies and merges multiple Excel files with the same headers.
- **Format Preservation**: Maintains original cell styles and data types.
- **Batch Processing**: Supports one-click merging of entire folders.

## 🛠️ Technical Implementation

```python
# Core merge logic example
def merge_excels(file_list, output_path):
    combined_df = pd.DataFrame()
    for file in file_list:
        df = pd.read_excel(file)
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    combined_df.to_excel(output_path, index=False)
```

**Tech Stack**:
- pandas (Data processing)
- openpyxl (Excel operations)
- tkinter (GUI interface)

## 📦 Usage

### GUI Operation
1. Double-click `excel合并工具.exe` to run.
2. Select the folder containing the Excel files.
3. Set the output file path.
4. Click the "Start Merge" button.

### Command Line Operation
```bash
python excel合并工具_code.py -i "Input Folder Path" -o "Output File.xlsx"
```

## 📂 File Structure
```
python_Excel_merge_tool/
├── excel合并工具.exe      # Executable program
├── excel合并工具_code.py  # Source code
```

## 💡 Typical Use Cases

1. **Monthly Report Merging**: Combine 30 individual sheets into a quarterly summary.
2. **Multi-Department Data Aggregation**: Consolidate reports from Sales, Finance, and Production departments.
3. **Scientific Data Processing**: Merge repeated experimental data sets.

## 🚨 Important Notes

- Ensure all files have identical headers.
- It is recommended to back up original files beforehand.
- Merging 100,000+ rows may take approximately 2 minutes.
- Chinese file paths require Python 3.7 or higher.

---

⭐ **If this tool saves you time, please give it a Star!**  
🐛 **Issue Reporting**: Please attach sample files and screenshots of error logs.
