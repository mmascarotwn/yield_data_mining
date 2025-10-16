#!/usr/bin/env python3
"""
Yield Calculator Module

This module provides functionality to add yield calculation columns to Excel files.
It specifically adds 'e_yield' and 'asm_yield' columns to the configured target sheet,
leaving all other sheets unchanged.

Features:
- Configurable target sheet (via TARGET_SHEET_NAME variable)
- Configurable column names for yield calculations
- Calculates e_yield as E_YIELD_NUMERATOR_COL/E_YIELD_DENOMINATOR_COL for each row
- Calculates asm_yield as ASM_YIELD_NUMERATOR_COL/ASM_YIELD_DENOMINATOR_COL for each row
- GUI-based file selection
- Preserves all other sheets unchanged
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

# Configuration - Change this to target a different sheet
TARGET_SHEET_NAME = 'hbm_test_yield'

# Configuration - Column names for yield calculations
# Users can modify these variables to match their Excel file's column names
E_YIELD_NUMERATOR_COL = 'qty_e_out'    # Column used as numerator for e_yield calculation
E_YIELD_DENOMINATOR_COL = 'qty_e_in'  # Column used as denominator for e_yield calculation
ASM_YIELD_NUMERATOR_COL = 'qty_asm_out'  # Column used as numerator for asm_yield calculation
ASM_YIELD_DENOMINATOR_COL = 'qty_asm_in' # Column used as denominator for asm_yield calculation

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class YieldCalculator:
    """
    A class to handle adding yield calculation columns to Excel files.
    
    The target sheet name is configurable via the TARGET_SHEET_NAME variable.
    By default, it processes 'Sheet1', but users can modify this at the top of the file.
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
        Load the selected Excel file, specifically targeting the configured target sheet.
        
        Returns:
            True if file loaded successfully, False otherwise
        """
        try:
            if not self.input_file_path:
                raise ValueError("Input file path not set. Please select a file first.")
            
            logger.info("Loading Excel file...")
            
            # Try to load the target sheet specifically
            try:
                # First, check what sheets are available
                excel_file = pd.ExcelFile(self.input_file_path)
                available_sheets = excel_file.sheet_names
                logger.info(f"Available sheets in file: {available_sheets}")
                
                if TARGET_SHEET_NAME in available_sheets:
                    # Load target sheet specifically
                    df = pd.read_excel(self.input_file_path, sheet_name=TARGET_SHEET_NAME)
                    self.sheets_data = {TARGET_SHEET_NAME: df}
                    logger.info(f"Loaded '{TARGET_SHEET_NAME}': {df.shape[0]} rows, {df.shape[1]} columns")
                else:
                    # Target sheet not found, inform user and load first available sheet as target
                    first_sheet = available_sheets[0]
                    logger.warning(f"'{TARGET_SHEET_NAME}' not found. Loading first available sheet '{first_sheet}' as '{TARGET_SHEET_NAME}'")
                    df = pd.read_excel(self.input_file_path, sheet_name=first_sheet)
                    self.sheets_data = {TARGET_SHEET_NAME: df}
                    logger.info(f"Loaded '{first_sheet}' as '{TARGET_SHEET_NAME}': {df.shape[0]} rows, {df.shape[1]} columns")
                    
                    messagebox.showinfo("Sheet Information", 
                        f"'{TARGET_SHEET_NAME}' not found in the file.\n"
                        f"Using '{first_sheet}' instead.\n\n"
                        f"Available sheets: {available_sheets}")
                
                excel_file.close()
                    
            except Exception as e:
                # Fallback to loading the default sheet
                logger.warning(f"Could not load specific sheet, loading default sheet: {str(e)}")
                df = pd.read_excel(self.input_file_path)
                self.sheets_data = {TARGET_SHEET_NAME: df}
                logger.info(f"Loaded default sheet as '{TARGET_SHEET_NAME}': {df.shape[0]} rows, {df.shape[1]} columns")
            
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
            # Show available columns for reference (from target sheet)
            if self.sheets_data and TARGET_SHEET_NAME in self.sheets_data:
                target_df = self.sheets_data[TARGET_SHEET_NAME]
                columns_info = f"Available columns in '{TARGET_SHEET_NAME}': {list(target_df.columns)}"
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
        Add e_yield and asm_yield columns to the target sheet only.
        e_yield will be calculated as E_YIELD_NUMERATOR_COL/E_YIELD_DENOMINATOR_COL for each row.
        asm_yield will be calculated as ASM_YIELD_NUMERATOR_COL/ASM_YIELD_DENOMINATOR_COL for each row.
        
        Args:
            e_yield_method: Method for calculating e_yield (not used - uses configured column names)
            asm_yield_method: Method for calculating asm_yield (not used - uses configured column names)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.sheets_data:
                raise ValueError("No data loaded. Please load a file first.")
            
            if TARGET_SHEET_NAME not in self.sheets_data:
                raise ValueError(f"{TARGET_SHEET_NAME} not found in the loaded data.")
            
            self.processed_sheets = {}
            
            # Process only the target sheet
            sheet_name = TARGET_SHEET_NAME
            df = self.sheets_data[sheet_name]
            logger.info(f"Processing sheet: {sheet_name}")
            
            # Create a copy of the DataFrame
            processed_df = df.copy()
            
            # Check if required columns exist for e_yield calculation
            if E_YIELD_NUMERATOR_COL in processed_df.columns and E_YIELD_DENOMINATOR_COL in processed_df.columns:
                # Calculate e_yield as numerator/denominator
                try:
                    # Handle division by zero and invalid values
                    processed_df['e_yield'] = processed_df[E_YIELD_NUMERATOR_COL] / processed_df[E_YIELD_DENOMINATOR_COL]
                    # Replace inf and -inf with 0
                    processed_df['e_yield'] = processed_df['e_yield'].replace([float('inf'), float('-inf')], 0)
                    # Fill NaN values with 0
                    processed_df['e_yield'] = processed_df['e_yield'].fillna(0)
                    logger.info(f"Calculated 'e_yield' as {E_YIELD_NUMERATOR_COL}/{E_YIELD_DENOMINATOR_COL} for {TARGET_SHEET_NAME}")
                except Exception as calc_error:
                    logger.warning(f"Error in e_yield calculation: {str(calc_error)}")
                    logger.warning("Setting e_yield to 0.0 for all rows")
                    processed_df['e_yield'] = 0.0
            else:
                # If required columns don't exist, set e_yield to 0.0
                missing_cols = []
                if E_YIELD_NUMERATOR_COL not in processed_df.columns:
                    missing_cols.append(E_YIELD_NUMERATOR_COL)
                if E_YIELD_DENOMINATOR_COL not in processed_df.columns:
                    missing_cols.append(E_YIELD_DENOMINATOR_COL)
                
                logger.warning(f"Required columns missing for e_yield calculation: {missing_cols}")
                logger.warning("Setting e_yield to 0.0 for all rows")
                processed_df['e_yield'] = 0.0
                
                # Show warning to user
                messagebox.showwarning("Missing Columns", 
                    f"Cannot calculate e_yield as {E_YIELD_NUMERATOR_COL}/{E_YIELD_DENOMINATOR_COL}.\n"
                    f"Missing columns: {missing_cols}\n\n"
                    f"Available columns: {list(processed_df.columns)}\n\n"
                    f"Setting e_yield to 0.0 for all rows.")
            
            # Add asm_yield column - calculate as numerator/denominator
            if ASM_YIELD_NUMERATOR_COL in processed_df.columns and ASM_YIELD_DENOMINATOR_COL in processed_df.columns:
                try:
                    # Handle division by zero and invalid values
                    processed_df['asm_yield'] = processed_df[ASM_YIELD_NUMERATOR_COL] / processed_df[ASM_YIELD_DENOMINATOR_COL]
                    # Replace inf and -inf with 0
                    processed_df['asm_yield'] = processed_df['asm_yield'].replace([float('inf'), float('-inf')], 0)
                    # Fill NaN values with 0
                    processed_df['asm_yield'] = processed_df['asm_yield'].fillna(0)
                    logger.info(f"Calculated 'asm_yield' as {ASM_YIELD_NUMERATOR_COL}/{ASM_YIELD_DENOMINATOR_COL} for {TARGET_SHEET_NAME}")
                except Exception as calc_error:
                    logger.warning(f"Error in asm_yield calculation: {str(calc_error)}")
                    logger.warning("Setting asm_yield to 0.0 for all rows")
                    processed_df['asm_yield'] = 0.0
            else:
                # If required columns don't exist, set asm_yield to 0.0
                missing_cols_asm = []
                if ASM_YIELD_NUMERATOR_COL not in processed_df.columns:
                    missing_cols_asm.append(ASM_YIELD_NUMERATOR_COL)
                if ASM_YIELD_DENOMINATOR_COL not in processed_df.columns:
                    missing_cols_asm.append(ASM_YIELD_DENOMINATOR_COL)
                
                logger.warning(f"Required columns missing for asm_yield calculation: {missing_cols_asm}")
                logger.warning("Setting asm_yield to 0.0 for all rows")
                processed_df['asm_yield'] = 0.0
                
                # Show warning to user if e_yield warning wasn't already shown
                e_yield_missing = E_YIELD_NUMERATOR_COL not in processed_df.columns or E_YIELD_DENOMINATOR_COL not in processed_df.columns
                if not e_yield_missing:
                    messagebox.showwarning("Missing Columns for ASM Yield", 
                        f"Cannot calculate asm_yield as {ASM_YIELD_NUMERATOR_COL}/{ASM_YIELD_DENOMINATOR_COL}.\n"
                        f"Missing columns: {missing_cols_asm}\n\n"
                        f"Available columns: {list(processed_df.columns)}\n\n"
                        f"Setting asm_yield to 0.0 for all rows.")
                elif ASM_YIELD_NUMERATOR_COL in missing_cols_asm:
                    # Show additional warning for missing numerator if it's different from e_yield columns
                    messagebox.showwarning("Additional Missing Column", 
                        f"Column '{ASM_YIELD_NUMERATOR_COL}' is also missing for asm_yield calculation.\n"
                        f"Setting asm_yield to 0.0 for all rows.")
            
            self.processed_sheets[sheet_name] = processed_df
            
            logger.info(f"Added yield columns to '{sheet_name}'. "
                      f"Final shape: {processed_df.shape[0]} rows, {processed_df.shape[1]} columns")
            
            logger.info("Yield columns added to target sheet successfully")
            return True
            
        except Exception as e:
            logger.error(f"Error adding yield columns: {str(e)}")
            messagebox.showerror("Error", f"Error adding yield columns: {str(e)}")
            return False
    
    def save_processed_file(self, output_path: Optional[str] = None) -> bool:
        """
        Save the processed DataFrame with yield columns to an Excel file.
        Preserves all original sheets and only modifies the target sheet.
        
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
                # Save the processed target sheet
                processed_target_sheet = self.processed_sheets[TARGET_SHEET_NAME]
                processed_target_sheet.to_excel(writer, sheet_name=TARGET_SHEET_NAME, index=False)
                logger.info(f"Saved modified '{TARGET_SHEET_NAME}' with {len(processed_target_sheet)} rows and {len(processed_target_sheet.columns)} columns")
                
                # Save all other original sheets unchanged
                for sheet_name, df in all_original_sheets.items():
                    if sheet_name != TARGET_SHEET_NAME:  # Don't overwrite the modified target sheet
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        logger.info(f"Preserved original sheet '{sheet_name}' with {len(df)} rows and {len(df.columns)} columns")
            
            messagebox.showinfo("Success", 
                f"File with yield columns successfully saved to:\n{output_path}\n\n"
                f"Modified: {TARGET_SHEET_NAME}\n"
                f"  • Added e_yield column (calculated as {E_YIELD_NUMERATOR_COL}/{E_YIELD_DENOMINATOR_COL})\n"
                f"  • Added asm_yield column (calculated as {ASM_YIELD_NUMERATOR_COL}/{ASM_YIELD_DENOMINATOR_COL})\n\n"
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
            
            # Step 3: Add yield columns to target sheet only
            if not self.add_yield_columns("default", "default"):  # Both columns calculated as specified
                return False
            
            # Step 4: Show summary and ask about saving
            target_sheet_df = self.processed_sheets[TARGET_SHEET_NAME]
            total_rows = len(target_sheet_df)
            
            summary_message = (
                f"Yield columns added successfully to {TARGET_SHEET_NAME}!\n\n"
                f"Sheet processed: {TARGET_SHEET_NAME}\n"
                f"Rows in {TARGET_SHEET_NAME}: {total_rows}\n"
                f"Columns added:\n"
                f"  • e_yield: calculated as {E_YIELD_NUMERATOR_COL}/{E_YIELD_DENOMINATOR_COL}\n"
                f"  • asm_yield: calculated as {ASM_YIELD_NUMERATOR_COL}/{ASM_YIELD_DENOMINATOR_COL}\n\n"
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
    print(f"🔢 Will calculate e_yield as {E_YIELD_NUMERATOR_COL}/{E_YIELD_DENOMINATOR_COL} for each row in {TARGET_SHEET_NAME}")
    print(f"📋 Will calculate asm_yield as {ASM_YIELD_NUMERATOR_COL}/{ASM_YIELD_DENOMINATOR_COL} for each row in {TARGET_SHEET_NAME}")
    print("🔒 All other sheets will be preserved unchanged")
    
    success = add_yield_columns()
    
    if success:
        print("✅ Yield calculation process completed successfully!")
    else:
        print("❌ Yield calculation process failed or was cancelled")
