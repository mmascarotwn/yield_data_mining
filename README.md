# Excel File Merger & Yield Calculator

A comprehensive Python toolkit for Excel file processing with merging capabilities and yield calculations.

## Modules

### 1. Excel File Merger
Merges two Excel files with automatic duplicate detection and removal.

### 2. Yield Calculator  
Adds e_yield and asm_yield columns to Excel files with flexible calculation methods.

### 3. Web Scraper
Scrapes data from websites and integrates it with Excel files as new columns.

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
- 📊 **Configurable Target Sheet**: Easy sheet name configuration via TARGET_SHEET_NAME variable
- 🔢 **Configurable Column Names**: Customizable column names for yield calculations
- 📁 **GUI Interface**: Easy file selection and parameter input
- 💾 **Safe Processing**: Creates backups before modification
- 📈 **Comprehensive Logging**: Detailed process tracking and statistics
- ⚡ **User-Friendly**: Simple interface for complex calculations

### Web Scraper Features
- 🌐 **Website Automation**: Selenium-based browser automation for complex interactions
- 📝 **Form Filling**: Automatic form field completion and button clicking
- 🔍 **Data Extraction**: Extract text, attributes, and other data from web pages
- ⚙️ **Configurable Rules**: JSON-based configuration for websites and scraping rules
- 📊 **Excel Integration**: Seamlessly add scraped data as new columns to Excel files
- 🔧 **Multiple Browsers**: Support for Chrome, Firefox, and headless operation
- 🛡️ **Error Handling**: Robust error handling with retry mechanisms
- 📈 **Detailed Logging**: Comprehensive logging of scraping operations

## Installation

1. **Install required dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **The main dependencies are**:
   - `pandas` - For data manipulation
   - `openpyxl` - For Excel file handling
   - `tkinter` - For GUI dialogs (usually comes with Python)

3. **Optional web scraping dependencies** (for web_scraper.py):
   ```bash
   pip install selenium beautifulsoup4 requests
   ```
   
   - `selenium` - For browser automation
   - `beautifulsoup4` - For HTML parsing
   - `requests` - For HTTP requests
   
   **Note**: You'll also need to install browser drivers:
   - [ChromeDriver](https://chromedriver.chromium.org/) for Chrome
   - [GeckoDriver](https://github.com/mozilla/geckodriver/releases) for Firefox

## Usage

### Running the Demo

```bash
# Excel Merger Demo
python demo_excel_merger.py

# Yield Calculator Demo  
python demo_yield_calculator.py

# Web Scraper Demo
python demo_web_scraper.py

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

### Web Scraper Usage

#### Simple Usage (Recommended)

```python
from src.utils.web_scraper import run_web_scraper

# Run the complete web scraping and Excel integration process
success = run_web_scraper()
```

#### Advanced Usage with Configuration

```python
from src.utils.web_scraper import WebScraper, WebScrapingConfig

# Create and configure scraper
scraper = WebScraper()
config = WebScrapingConfig()

# Configure websites
config.add_website(
    name="data_source",
    url="https://example.com/data",
    description="Source for data extraction"
)

# Configure scraping rules
config.add_scraping_rule(
    rule_name="extract_value",
    website_name="data_source",
    selector=".data-value",
    action_type="extract"
)

# Set configuration and run
scraper.config = config
scraper.setup_driver()
scraped_data = scraper.scrape_website_data("data_source", ["extract_value"])

# Integrate with Excel
scraper.select_excel_file()
scraper.load_excel_file()
scraper.add_scraped_data_to_excel(scraped_data)
scraper.save_excel_file()
```

#### Configuration File Usage

```python
from src.utils.web_scraper import WebScraper, WebScrapingConfig

# Load configuration from JSON file
config = WebScrapingConfig()
config.load_config("web_scraper_config.json")

# Use configuration
scraper = WebScraper()
scraper.config = config
success = scraper.run_complete_scraping_process()
```

#### Yield Calculator Configuration

Users can easily configure the yield calculator by modifying variables at the top of the file:

```python
# In src/utils/yield_calculator.py

# Configuration - Change this to target a different sheet
TARGET_SHEET_NAME = 'MyDataSheet'  # Instead of default 'Sheet1'

# Configuration - Column names for yield calculations
E_YIELD_NUMERATOR_COL = 'Pass_Count'      # Instead of 'Data 2'
E_YIELD_DENOMINATOR_COL = 'Total_Count'   # Instead of 'Data 3'
ASM_YIELD_NUMERATOR_COL = 'Good_Units'    # Instead of 'Data 1'
ASM_YIELD_DENOMINATOR_COL = 'Test_Units'  # Instead of 'Data 2'
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

### Web Scraper Process

1. **Configuration Setup**:
   - Define target websites with URLs and descriptions
   - Create scraping rules for data extraction
   - Configure Excel integration settings

2. **Browser Initialization**:
   - Sets up Selenium WebDriver (Chrome/Firefox)
   - Configures browser options (headless mode, timeouts)
   - Handles driver installation and setup

3. **Website Navigation**:
   - Navigates to configured URLs
   - Handles page loading and timeouts
   - Manages browser state and cookies

4. **Interactive Operations**:
   - Fill form fields with specified values
   - Click buttons and interact with page elements
   - Wait for dynamic content to load

5. **Data Extraction**:
   - Extract text content from specified elements
   - Get attribute values (href, src, etc.)
   - Handle multiple data points per page

6. **Excel Integration**:
   - Load target Excel file and sheets
   - Add scraped data as new columns with configurable prefix
   - Maintain data integrity and Excel structure

7. **Data Persistence**:
   - Save updated Excel file with scraped data
   - Create backups before modification
   - Provide detailed logging of scraping operations

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
│       ├── yield_calculator.py  # Yield calculation module
│       └── web_scraper.py       # Web scraping module
├── demo_excel_merger.py         # Excel merger demo script
├── demo_yield_calculator.py     # Yield calculator demo script
├── demo_web_scraper.py          # Web scraper demo script
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

# Step 3: Scrape web data and integrate with Excel
from src.utils.web_scraper import scrape_and_add_to_excel
success = scrape_and_add_to_excel()
```

This creates a complete pipeline for:
1. Merging multiple Excel data sources
2. Adding standardized yield calculation columns
3. Scraping web data and integrating it with Excel
4. Preparing data for further analysis

## Dependencies

- `pandas>=1.5.0` - Data manipulation and analysis
- `openpyxl>=3.1.0` - Excel file reading/writing
- `xlrd>=2.0.0` - Excel file reading support
- `numpy>=1.24.0` - Numerical computing
- `tkinter` - GUI dialogs (usually included with Python)
- `selenium` - Web browser automation
- `webdriver-manager` - WebDriver management for Selenium

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
