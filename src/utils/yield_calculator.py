#!/usr/bin/env python3
"""
Yield Calculator Module

This module provides functionality to add yield calculation columns to Excel files.
It specifically adds 'e_yield' and 'asm_yield' columns to the 'Sheet1' tab only,
leaving all other sheets unchanged.

Features:
- Targeted Sheet1 processing only
- Calculates e_yield as Data 2/Data 3 for each row
- Calculates asm_yield as Data 1/Data 2 for each row
- GUI-based file selection
- Preservesif __name__ == "__main__":
    # Run the yield calculator
    print("📊 Starting Yield Calculator...")
    print("📁 Please select your Excel file in the popup dialog")
    print("🔢 Will calculate e_yield as Data 2/Data 3 for each row in Sheet1")
    print("📋 Will calculate asm_yield as Data 1/Data 2 for each row in Sheet1")
    print("🔒 All other sheets will be preserved unchanged")
    
    success = add_yield_columns()
    
    if success:
        print("✅ Yield calculation process completed successfully!")
    else:
        print("❌ Yield calculation process failed or was cancelled")eets unchanged
- Backup creation before modification
- Detailed logging and progress reporting
- Safe division handling (handles division by zero and missing columns)
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
from pathlib import Path
from typing import Optional, List, Dict, Tuple
import logging
import numpy as np

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class YieldCalculator:
    """
    A class to handle adding yield calculation columns to Sheet1 of Excel files.
    """
    
    def __init__(self):
        self.input_file_path: Optional[str] = None
        self.output_file_path: Optional[str] = None
        self.sheets_data: Optional[Dict[str, pd.DataFrame]] = None
        self.processed_sheets: Optional[Dict[str, pd.DataFrame]] = None
        
    def select_input_file(self) -> Optional[str]:
        """
        Open file selection dialog to choose the Excel file to process.
        
        Returns:
            Path to selected file or None if cancelled
        """
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        try:
            messagebox.showinfo("File Selection", "Please select the Excel file to add yield columns to")
            input_file = filedialog.askopenfilename(
                title="Select Excel File for Yield Calculation",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            
            if not input_file:
                messagebox.showwarning("Warning", "No file selected")
                return None
            
            self.input_file_path = input_file
            logger.info(f"Selected input file: {input_file}")
            
            return input_file
            
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting file: {str(e)}")
            return None
        finally:
            root.destroy()
    
    def load_excel_file(self) -> bool:
        """
        Load the selected Excel file, specifically targeting 'Sheet1'.
        
        Returns:
            True if file loaded successfully, False otherwise
        """
        try:
            if not self.input_file_path:
                raise ValueError("Input file path not set. Please select a file first.")
            
            logger.info("Loading Excel file...")
            
            # Try to load Sheet1 specifically
            try:
                # First, check what sheets are available
                excel_file = pd.ExcelFile(self.input_file_path)
                available_sheets = excel_file.sheet_names
                logger.info(f"Available sheets in file: {available_sheets}")
                
                if 'Sheet1' in available_sheets:
                    # Load Sheet1 specifically
                    df = pd.read_excel(self.input_file_path, sheet_name='Sheet1')
                    self.sheets_data = {"Sheet1": df}
                    logger.info(f"Loaded 'Sheet1': {df.shape[0]} rows, {df.shape[1]} columns")
                else:
                    # Sheet1 not found, inform user and load first available sheet as Sheet1
                    first_sheet = available_sheets[0]
                    logger.warning(f"'Sheet1' not found. Loading first available sheet '{first_sheet}' as 'Sheet1'")
                    df = pd.read_excel(self.input_file_path, sheet_name=first_sheet)
                    self.sheets_data = {"Sheet1": df}
                    logger.info(f"Loaded '{first_sheet}' as 'Sheet1': {df.shape[0]} rows, {df.shape[1]} columns")
                    
                    messagebox.showinfo("Sheet Information", 
                        f"'Sheet1' not found in the file.\n"
                        f"Using '{first_sheet}' instead.\n\n"
                        f"Available sheets: {available_sheets}")
                
                excel_file.close()
                    
            except Exception as e:
                # Fallback to loading the default sheet
                logger.warning(f"Could not load specific sheet, loading default sheet: {str(e)}")
                df = pd.read_excel(self.input_file_path)
                self.sheets_data = {"Sheet1": df}
                logger.info(f"Loaded default sheet as 'Sheet1': {df.shape[0]} rows, {df.shape[1]} columns")
            
            return True
            
        except Exception as e:
            logger.error(f"Error loading Excel file: {str(e)}")
            messagebox.showerror("Error", f"Error loading Excel file: {str(e)}")
            return False
    
    def get_yield_calculation_method(self) -> Tuple[str, str]:
        """
        Get user input for yield calculation methods.
        
        Returns:
            Tuple of (e_yield_method, asm_yield_method)
        """
        root = tk.Tk()
        root.withdraw()
        
        try:
            # Show available columns for reference (from Sheet1)
            if self.sheets_data and 'Sheet1' in self.sheets_data:
                sheet1_df = self.sheets_data['Sheet1']
                columns_info = f"Available columns in 'Sheet1': {list(sheet1_df.columns)}"
                messagebox.showinfo("Column Information", columns_info)
            
            # E-Yield calculation method
            e_yield_method = simpledialog.askstring(
                "E-Yield Calculation",
                "Enter the calculation method for 'e_yield' column:\n\n"
                "Options:\n"
                "1. 'default' - Set all values to 0.0\n"
                "2. 'column_name' - Copy values from existing column\n"
                "3. 'formula' - Enter a pandas expression (e.g., 'col1 / col2 * 100')\n\n"
                "Enter your choice:",
                initialvalue="default"
            )
            
            if not e_yield_method:
                return None, None
                
            # ASM-Yield calculation method
            asm_yield_method = simpledialog.askstring(
                "ASM-Yield Calculation",
                "Enter the calculation method for 'asm_yield' column:\n\n"
                "Options:\n"
                "1. 'default' - Set all values to 0.0\n"
                "2. 'column_name' - Copy values from existing column\n"
                "3. 'formula' - Enter a pandas expression (e.g., 'col1 / col2 * 100')\n\n"
                "Enter your choice:",
                initialvalue="default"
            )
            
            if not asm_yield_method:
                return None, None
            
            return e_yield_method.strip(), asm_yield_method.strip()
            
        except Exception as e:
            logger.error(f"Error getting calculation methods: {str(e)}")
            return None, None
        finally:
            root.destroy()
    
    def calculate_yield_value(self, df: pd.DataFrame, method: str, column_name: str) -> pd.Series:
        """
        Calculate yield values based on the specified method.
        
        Args:
            df: DataFrame to process
            method: Calculation method
            column_name: Name of the column being calculated (for logging)
            
        Returns:
            Pandas Series with calculated values
        """
        try:
            if method.lower() == 'default':
                # Set all values to 0.0
                result = pd.Series([0.0] * len(df), index=df.index)
                logger.info(f"Set {column_name} to default value (0.0) for {len(df)} rows")
                
            elif method in df.columns:
                # Copy values from existing column
                result = df[method].copy()
                logger.info(f"Copied {column_name} values from column '{method}'")
                
            else:
                # Try to evaluate as a formula
                try:
                    # Replace column names in the formula with actual column references
                    formula = method
                    for col in df.columns:
                        formula = formula.replace(col, f"df['{col}']")
                    
                    result = eval(formula)
                    if not isinstance(result, pd.Series):
                        result = pd.Series([result] * len(df), index=df.index)
                    
                    logger.info(f"Calculated {column_name} using formula: {method}")
                    
                except Exception as formula_error:
                    logger.warning(f"Formula evaluation failed for {column_name}: {str(formula_error)}")
                    logger.warning(f"Setting {column_name} to default value (0.0)")
                    result = pd.Series([0.0] * len(df), index=df.index)
            
            return result
            
        except Exception as e:
            logger.error(f"Error calculating {column_name}: {str(e)}")
            # Return default values on error
            return pd.Series([0.0] * len(df), index=df.index)
    
    def add_yield_columns(self, e_yield_method: str, asm_yield_method: str) -> bool:
        """
        Add e_yield and asm_yield columns to Sheet1 only.
        e_yield will be calculated as Data 2/Data 3 for each row.
        asm_yield will be calculated as Data 1/Data 2 for each row.
        
        Args:
            e_yield_method: Method for calculating e_yield (not used - hardcoded to Data 2/Data 3)
            asm_yield_method: Method for calculating asm_yield (not used - hardcoded to Data 1/Data 2)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.sheets_data:
                raise ValueError("No data loaded. Please load a file first.")
            
            if 'Sheet1' not in self.sheets_data:
                raise ValueError("Sheet1 not found in the loaded data.")
            
            self.processed_sheets = {}
            
            # Process only Sheet1
            sheet_name = 'Sheet1'
            df = self.sheets_data[sheet_name]
            logger.info(f"Processing sheet: {sheet_name}")
            
            # Create a copy of the DataFrame
            processed_df = df.copy()
            
            # Check if required columns exist for e_yield calculation
            if 'Data 2' in processed_df.columns and 'Data 3' in processed_df.columns:
                # Calculate e_yield as data2/data3
                try:
                    # Handle division by zero and invalid values
                    processed_df['e_yield'] = processed_df['Data 2'] / processed_df['Data 3']
                    # Replace inf and -inf with 0
                    processed_df['e_yield'] = processed_df['e_yield'].replace([float('inf'), float('-inf')], 0)
                    # Fill NaN values with 0
                    processed_df['e_yield'] = processed_df['e_yield'].fillna(0)
                    logger.info(f"Calculated 'e_yield' as Data 2/Data 3 for Sheet1")
                except Exception as calc_error:
                    logger.warning(f"Error in e_yield calculation: {str(calc_error)}")
                    logger.warning("Setting e_yield to 0.0 for all rows")
                    processed_df['e_yield'] = 0.0
            else:
                # If Data 2 or Data 3 columns don't exist, set e_yield to 0.0
                missing_cols = []
                if 'Data 2' not in processed_df.columns:
                    missing_cols.append('Data 2')
                if 'Data 3' not in processed_df.columns:
                    missing_cols.append('Data 3')
                
                logger.warning(f"Required columns missing for e_yield calculation: {missing_cols}")
                logger.warning("Setting e_yield to 0.0 for all rows")
                processed_df['e_yield'] = 0.0
                
                # Show warning to user
                messagebox.showwarning("Missing Columns", 
                    f"Cannot calculate e_yield as Data 2/Data 3.\n"
                    f"Missing columns: {missing_cols}\n\n"
                    f"Available columns: {list(processed_df.columns)}\n\n"
                    f"Setting e_yield to 0.0 for all rows.")
            
            # Add asm_yield column - calculate as Data 1/Data 2
            if 'Data 1' in processed_df.columns and 'Data 2' in processed_df.columns:
                try:
                    # Handle division by zero and invalid values
                    processed_df['asm_yield'] = processed_df['Data 1'] / processed_df['Data 2']
                    # Replace inf and -inf with 0
                    processed_df['asm_yield'] = processed_df['asm_yield'].replace([float('inf'), float('-inf')], 0)
                    # Fill NaN values with 0
                    processed_df['asm_yield'] = processed_df['asm_yield'].fillna(0)
                    logger.info(f"Calculated 'asm_yield' as Data 1/Data 2 for Sheet1")
                except Exception as calc_error:
                    logger.warning(f"Error in asm_yield calculation: {str(calc_error)}")
                    logger.warning("Setting asm_yield to 0.0 for all rows")
                    processed_df['asm_yield'] = 0.0
            else:
                # If Data 1 or Data 2 columns don't exist, set asm_yield to 0.0
                missing_cols_asm = []
                if 'Data 1' not in processed_df.columns:
                    missing_cols_asm.append('Data 1')
                if 'Data 2' not in processed_df.columns:
                    missing_cols_asm.append('Data 2')
                
                logger.warning(f"Required columns missing for asm_yield calculation: {missing_cols_asm}")
                logger.warning("Setting asm_yield to 0.0 for all rows")
                processed_df['asm_yield'] = 0.0
                
                # Show warning to user if e_yield warning wasn't already shown
                e_yield_missing = 'Data 2' not in processed_df.columns or 'Data 3' not in processed_df.columns
                if not e_yield_missing:
                    messagebox.showwarning("Missing Columns for ASM Yield", 
                        f"Cannot calculate asm_yield as Data 1/Data 2.\n"
                        f"Missing columns: {missing_cols_asm}\n\n"
                        f"Available columns: {list(processed_df.columns)}\n\n"
                        f"Setting asm_yield to 0.0 for all rows.")
                elif 'Data 1' in missing_cols_asm:
                    # Show additional warning for Data 1 if it's missing and not mentioned in e_yield warning
                    messagebox.showwarning("Additional Missing Column", 
                        f"Column 'Data 1' is also missing for asm_yield calculation.\n"
                        f"Setting asm_yield to 0.0 for all rows.")
            
            self.processed_sheets[sheet_name] = processed_df
            
            logger.info(f"Added yield columns to '{sheet_name}'. "
                      f"Final shape: {processed_df.shape[0]} rows, {processed_df.shape[1]} columns")
            
            logger.info("Yield columns added to Sheet1 successfully")
            return True
            
        except Exception as e:
            logger.error(f"Error adding yield columns: {str(e)}")
            messagebox.showerror("Error", f"Error adding yield columns: {str(e)}")
            return False
    
    def save_processed_file(self, output_path: Optional[str] = None) -> bool:
        """
        Save the processed DataFrame with yield columns to an Excel file.
        Preserves all original sheets and only modifies Sheet1.
        
        Args:
            output_path: Optional custom output path. If None, creates a new file.
            
        Returns:
            True if save successful, False otherwise
        """
        try:
            if not self.processed_sheets:
                raise ValueError("No processed data to save. Please add yield columns first.")
            
            # Determine output path
            if output_path is None:
                input_path = Path(self.input_file_path)
                output_path = str(input_path.with_name(f"{input_path.stem}_with_yields{input_path.suffix}"))
            
            self.output_file_path = output_path
            
            # Create backup of original file
            backup_path = str(Path(self.input_file_path).with_suffix('.yield_backup.xlsx'))
            if os.path.exists(self.input_file_path):
                # Create backup with all original sheets
                original_sheets = pd.read_excel(self.input_file_path, sheet_name=None)
                with pd.ExcelWriter(backup_path, engine='openpyxl') as writer:
                    for sheet_name, df in original_sheets.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                logger.info(f"Backup created: {backup_path}")
            
            # Load all original sheets to preserve them
            try:
                all_original_sheets = pd.read_excel(self.input_file_path, sheet_name=None)
            except:
                all_original_sheets = {}
            
            # Save the processed file(s)
            logger.info(f"Saving processed file to: {output_path}")
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Save the processed Sheet1
                processed_sheet1 = self.processed_sheets['Sheet1']
                processed_sheet1.to_excel(writer, sheet_name='Sheet1', index=False)
                logger.info(f"Saved modified 'Sheet1' with {len(processed_sheet1)} rows and {len(processed_sheet1.columns)} columns")
                
                # Save all other original sheets unchanged
                for sheet_name, df in all_original_sheets.items():
                    if sheet_name != 'Sheet1':  # Don't overwrite the modified Sheet1
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        logger.info(f"Preserved original sheet '{sheet_name}' with {len(df)} rows and {len(df.columns)} columns")
            
            messagebox.showinfo("Success", 
                f"File with yield columns successfully saved to:\n{output_path}\n\n"
                f"Modified: Sheet1\n"
                f"  • Added e_yield column (calculated as Data 2/Data 3)\n"
                f"  • Added asm_yield column (calculated as Data 1/Data 2)\n\n"
                f"Preserved: All other original sheets unchanged")
            
            logger.info("File saved successfully")
            return True
            
        except Exception as e:
            logger.error(f"Error saving file: {str(e)}")
            messagebox.showerror("Error", f"Error saving file: {str(e)}")
            return False
    
    def run_complete_process(self) -> bool:
        """
        Run the complete yield calculation process: select file, add columns, and save.
        
        Returns:
            True if entire process successful, False otherwise
        """
        try:
            # Step 1: Select input file
            input_file = self.select_input_file()
            if not input_file:
                return False
            
            # Step 2: Load file
            if not self.load_excel_file():
                return False
            
            # Step 3: Add empty yield columns to Sheet1 only
            if not self.add_yield_columns("default", "default"):  # Both columns will be empty (0.0)
                return False
            
            # Step 4: Show summary and ask about saving
            sheet1_df = self.processed_sheets['Sheet1']
            total_rows = len(sheet1_df)
            
            summary_message = (
                f"Yield columns added successfully to Sheet1!\n\n"
                f"Sheet processed: Sheet1\n"
                f"Rows in Sheet1: {total_rows}\n"
                f"Columns added:\n"
                f"  • e_yield: calculated as Data 2/Data 3\n"
                f"  • asm_yield: calculated as Data 1/Data 2\n\n"
                f"All other sheets will be preserved unchanged.\n\n"
                f"Do you want to save the processed file?"
            )
            
            save_choice = messagebox.askyesno("Save File", summary_message)
            
            if save_choice:
                return self.save_processed_file()
            else:
                logger.info("User chose not to save the processed file")
                return True
                
        except Exception as e:
            logger.error(f"Error in complete yield calculation process: {str(e)}")
            messagebox.showerror("Error", f"Error in complete yield calculation process: {str(e)}")
            return False


def add_yield_columns():
    """
    Convenience function to run the complete yield calculation process.
    """
    calculator = YieldCalculator()
    return calculator.run_complete_process()


if __name__ == "__main__":
    # Run the yield calculator
    print("📊 Starting Yield Calculator...")
    print("📁 Please select your Excel file in the popup dialog")
    print("� Will add empty e_yield and asm_yield columns to Sheet1 only")
    print("🔒 All other sheets will be preserved unchanged")
    
    success = add_yield_columns()
    
    if success:
        print("✅ Yield calculation process completed successfully!")
    else:
        print("❌ Yield calculation process failed or was cancelled")
