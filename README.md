# Excel File Merger & Yield Calculator

A comprehensive Python toolkit for Excel file processing with merging capabilities and yield calculations.

## Modules

### 1. Excel File Merger
Merges two Excel files with automatic duplicate detection and removal.

### 2. Yield Calculator  
Adds e_yield and asm_yield columns to Excel files with flexible calculation methods.

## Features

### Excel Merger Features
- 🔍 **Automatic Duplicate Detection**: Identifies and removes duplicate rows when merging
- 📊 **Multi-Sheet Support**: Automatically processes all sheets with matching names
- 📁 **GUI File Selection**: User-friendly popup dialogs for file selection
- 🔄 **Column Alignment**: Automatically aligns columns between different Excel files
- 💾 **Safe Saving**: Creates backups before overwriting files
- 📈 **Detailed Statistics**: Per-sheet and total merge statistics
- ⚡ **Easy to Use**: Simple one-function call for complete merging

### Yield Calculator Features  
- 📊 **Multi-Sheet Processing**: Handles Excel files with multiple worksheets
- 🔢 **Flexible Calculations**: Multiple methods for calculating yield values
- 📁 **GUI Interface**: Easy file selection and parameter input
- 💾 **Safe Processing**: Creates backups before modification
- 📈 **Comprehensive Logging**: Detailed process tracking and statistics
- ⚡ **User-Friendly**: Simple interface for complex calculations

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

### Running the Demo

```bash
# Excel Merger Demo
python demo_excel_merger.py

# Yield Calculator Demo  
python demo_yield_calculator.py

# Advanced merger demo with step-by-step control
python demo_excel_merger.py --advanced
```

## Usage Examples

### Excel Merger Usage

#### Simple Usage (Recommended)

```python
from src.utils.excel_merger import merge_excel_files

# Run the complete merge process with GUI
success = merge_excel_files()
```

#### Advanced Usage

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

### Yield Calculator Usage

#### Simple Usage (Recommended)

```python
from src.utils.yield_calculator import add_yield_columns

# Run the complete yield calculation process with GUI
success = add_yield_columns()
```

#### Advanced Usage

```python
from src.utils.yield_calculator import YieldCalculator

# Create calculator instance for step-by-step control
calculator = YieldCalculator()

# Step 1: Select input file
input_file = calculator.select_input_file()

# Step 2: Load the file
calculator.load_excel_file()

# Step 3: Add yield columns with custom methods
calculator.add_yield_columns('default', 'column_name_here')

# Step 4: Save the result
calculator.save_processed_file()
```

## How It Works

### Excel Merger Process

1. **File Selection**: 
   - Opens GUI dialogs to select main and secondary Excel files
   - Main file will be updated with new data
   - Secondary file contains data to be merged

2. **Multi-Sheet Detection**:
   - Scans both files to identify available sheets
   - Finds sheets with matching names for processing
   - Preserves sheets that exist only in the main file

3. **Data Loading**:
   - Loads all sheets from both Excel files into pandas DataFrames
   - Handles various Excel formats (.xlsx, .xls)

4. **Column Alignment**:
   - Ensures both files have the same column structure
   - Adds missing columns with null values

5. **Duplicate Detection**:
   - Compares all rows between the two files (per sheet)
   - Uses row-wise hashing for efficient comparison
   - Identifies unique rows that don't exist in the main file

6. **Merging**:
   - Adds only non-duplicate rows to the main file
   - Processes each matching sheet independently
   - Preserves original data structure

7. **Saving**:
   - Creates a backup of the original main file
   - Saves the merged data with all sheets
   - Provides detailed per-sheet summary of changes

### Yield Calculator Process

1. **File Selection**:
   - Opens GUI dialog to select Excel file for processing
   - Can handle both single-sheet and multi-sheet files

2. **Data Loading**:
   - Loads all sheets from the Excel file
   - Analyzes existing column structure

3. **Calculation Method Selection**:
   - User specifies calculation methods for e_yield and asm_yield
   - Supports multiple calculation approaches:
     - **Default**: Sets all values to 0.0
     - **Column Copy**: Copies values from existing columns
     - **Formula**: Uses pandas expressions for calculations

4. **Yield Calculation**:
   - Applies specified calculations to each sheet
   - Handles errors gracefully with fallback to default values
   - Maintains data integrity throughout processing

5. **Column Addition**:
   - Adds e_yield and asm_yield columns to all sheets
   - Preserves original data and structure

6. **Saving**:
   - Creates backup of original file
   - Saves processed file with new yield columns
   - Provides detailed summary of processing results

## Calculation Methods

### Yield Calculator Methods

- **'default'**: Sets all yield values to 0.0
- **'column_name'**: Copies values from an existing column (e.g., 'existing_yield')
- **'formula'**: Uses pandas expressions for complex calculations

#### Formula Examples:
```python
# Percentage calculation
'pass_count / total_count * 100'

# Yield with efficiency factor
'good_units / tested_units * 0.95'

# Combined calculation
'(pass_units + rework_units) / total_units * 100'
```

## File Structure

```
yield_data_mining/
├── src/
│   └── utils/
│       ├── __init__.py
│       ├── excel_merger.py      # Excel file merger module
│       └── yield_calculator.py  # Yield calculation module
├── demo_excel_merger.py         # Excel merger demo script
├── demo_yield_calculator.py     # Yield calculator demo script
├── requirements.txt             # Dependencies
└── README.md                   # This file
```

## Example Workflow

### Complete Data Processing Pipeline

```python
# Step 1: Merge Excel files
from src.utils.excel_merger import merge_excel_files
success = merge_excel_files()

# Step 2: Add yield columns to the merged file
from src.utils.yield_calculator import add_yield_columns
success = add_yield_columns()
```

This creates a complete pipeline for:
1. Merging multiple Excel data sources
2. Adding standardized yield calculation columns
3. Preparing data for further analysis

## Dependencies

- `pandas>=1.5.0` - Data manipulation and analysis
- `openpyxl>=3.1.0` - Excel file reading/writing
- `xlrd>=2.0.0` - Excel file reading support
- `numpy>=1.24.0` - Numerical computing
- `tkinter` - GUI dialogs (usually included with Python)

## License

This project is provided as-is for educational and development purposes.

## Contributing

1. Fork the repository
2. Create your feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## Support

For issues or questions:
1. Check the console logs for detailed error messages
2. Ensure all dependencies are installed correctly
3. Verify that Excel files are not corrupted or password-protected
4. Check that file paths are accessible and writable
