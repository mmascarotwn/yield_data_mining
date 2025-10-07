# Excel File Merger

A Python module that merges two Excel files with automatic duplicate detection and removal.

## Features

- 🔍 **Automatic Duplicate Detection**: Identifies and removes duplicate rows when merging
- 📁 **GUI File Selection**: User-friendly popup dialogs for file selection
- 🔄 **Column Alignment**: Automatically aligns columns between different Excel files
- 💾 **Safe Saving**: Creates backups before overwriting files
- 📊 **Detailed Logging**: Comprehensive logging of the merge process
- ⚡ **Easy to Use**: Simple one-function call for complete merging

## Installation

1. **Install required dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **The main dependencies are**:
   - `pandas` - For data manipulation
   - `openpyxl` - For Excel file handling
   - `tkinter` - For GUI dialogs (usually comes with Python)

## Usage

### Simple Usage (Recommended)

```python
from src.utils.excel_merger import merge_excel_files

# Run the complete merge process with GUI
success = merge_excel_files()
```

### Advanced Usage

```python
from src.utils.excel_merger import ExcelMerger

# Create merger instance for step-by-step control
merger = ExcelMerger()

# Step 1: Select files via GUI
main_file, secondary_file = merger.select_files()

# Step 2: Perform merge
if merger.merge_files():
    # Step 3: Save the result
    merger.save_merged_file()
```

### Running the Demo

```bash
# Simple demo
python demo_excel_merger.py

# Advanced demo with step-by-step control
python demo_excel_merger.py --advanced
```

## How It Works

1. **File Selection**: 
   - Opens GUI dialogs to select main and secondary Excel files
   - Main file will be updated with new data
   - Secondary file contains data to be merged

2. **Data Loading**:
   - Loads both Excel files into pandas DataFrames
   - Handles various Excel formats (.xlsx, .xls)

3. **Column Alignment**:
   - Ensures both files have the same column structure
   - Adds missing columns with null values

4. **Duplicate Detection**:
   - Compares all rows between the two files
   - Uses row-wise hashing for efficient comparison
   - Identifies unique rows that don't exist in the main file

5. **Merging**:
   - Adds only non-duplicate rows to the main file
   - Preserves original data structure

6. **Saving**:
   - Creates a backup of the original main file
   - Saves the merged data to the main file location
   - Provides detailed summary of changes

## File Structure

```
src/
├── utils/
│   ├── __init__.py
│   └── excel_merger.py      # Main merger module
├── demo_excel_merger.py     # Demo script
├── requirements.txt         # Dependencies
└── README.md               # This file
```

## Example Workflow

1. **Run the merger**:
   ```bash
   python demo_excel_merger.py
   ```

2. **Select main file**: Choose the Excel file you want to update
3. **Select secondary file**: Choose the Excel file with new data to add
4. **Review results**: The tool will show:
   - Number of rows in original file
   - Number of new unique rows found
   - Number of duplicates detected
5. **Confirm save**: Choose whether to save the merged file

## Error Handling

- **File Format Validation**: Ensures files are valid Excel formats
- **Column Mismatch**: Automatically handles different column structures
- **Memory Management**: Efficient processing of large files
- **Backup Creation**: Automatic backup before overwriting files
- **User Feedback**: Clear error messages and progress updates

## Logging

The module provides detailed logging information:
- File loading progress
- Duplicate detection results
- Column alignment changes
- Save operation status

## Limitations

- Both files must be Excel format (.xlsx or .xls)
- Duplicate detection is based on exact row matching
- Large files may require significant memory
- GUI requires a display (not suitable for headless servers)

## Troubleshooting

**Import Error**: Make sure you're running from the project root directory and have installed dependencies.

**File Selection Cancelled**: The GUI dialogs require user interaction - make sure to select both files.

**Memory Issues**: For very large files, consider processing in chunks or using a more powerful machine.

**Permission Errors**: Ensure you have write permissions for the main file location.

## Contributing

To contribute to this module:
1. Fork the repository
2. Create a feature branch
3. Add tests for new functionality
4. Submit a pull request

## License

This project is licensed under the MIT License.
