#!/usr/bin/env python3
"""
Excel File Merger Module

This module provides functionality to merge two Excel files with duplicate detection and removal.
It includes a GUI for file selection and handles the merging process automatically.

Features:
- Single sheet and multi-sheet Excel file merging
- Automatic duplicate detection and removal
- Column alignment between files
- GUI-based file selection
- Backup creation before overwriting
- Detailed logging and progress reporting

Multi-sheet Support:
- Automatically detects and processes all sheets with matching names
- Preserves sheets that exist only in the main file
- Provides detailed per-sheet merge statistics
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from pathlib import Path
from typing import Tuple, Optional, List
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ExcelMerger:
    """
    A class to handle merging of two Excel files with duplicate checking.
    """
    
    def __init__(self):
        self.main_file_path: Optional[str] = None
        self.secondary_file_path: Optional[str] = None
        self.main_df: Optional[pd.DataFrame] = None
        self.secondary_df: Optional[pd.DataFrame] = None
        self.main_sheets: Optional[dict] = None  # Dictionary of sheet_name: DataFrame
        self.secondary_sheets: Optional[dict] = None  # Dictionary of sheet_name: DataFrame
        self.merged_sheets: Optional[dict] = None  # Dictionary of merged sheets
        
    def select_files(self) -> Tuple[Optional[str], Optional[str]]:
        """
        Open file selection dialogs to choose main and secondary Excel files.
        
        Returns:
            Tuple of (main_file_path, secondary_file_path)
        """
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        try:
            # Select main file
            messagebox.showinfo("File Selection", "Please select the MAIN Excel file (this will be updated)")
            main_file = filedialog.askopenfilename(
                title="Select Main Excel File",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            
            if not main_file:
                messagebox.showwarning("Warning", "No main file selected")
                return None, None
            
            # Select secondary file
            messagebox.showinfo("File Selection", "Please select the SECONDARY Excel file (data to merge)")
            secondary_file = filedialog.askopenfilename(
                title="Select Secondary Excel File",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            
            if not secondary_file:
                messagebox.showwarning("Warning", "No secondary file selected")
                return None, None
            
            self.main_file_path = main_file
            self.secondary_file_path = secondary_file
            
            logger.info(f"Selected main file: {main_file}")
            logger.info(f"Selected secondary file: {secondary_file}")
            
            return main_file, secondary_file
            
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting files: {str(e)}")
            return None, None
        finally:
            root.destroy()
    
    def load_excel_files(self) -> bool:
        """
        Load the selected Excel files into DataFrames (all sheets).
        
        Returns:
            True if both files loaded successfully, False otherwise
        """
        try:
            if not self.main_file_path or not self.secondary_file_path:
                raise ValueError("File paths not set. Please select files first.")
            
            # Get sheet information
            main_sheets, secondary_sheets, common_sheets = self.get_sheet_info()
            
            if not common_sheets:
                messagebox.showwarning("Warning", "No common sheet names found between the files!")
                # Load only the first sheet as fallback
                logger.info("Loading first sheet from each file as fallback...")
                self.main_df = pd.read_excel(self.main_file_path)
                self.secondary_df = pd.read_excel(self.secondary_file_path)
                logger.info(f"Main file (first sheet) loaded: {self.main_df.shape[0]} rows, {self.main_df.shape[1]} columns")
                logger.info(f"Secondary file (first sheet) loaded: {self.secondary_df.shape[0]} rows, {self.secondary_df.shape[1]} columns")
                return True
            
            # Load all sheets from both files
            logger.info("Loading all sheets from main Excel file...")
            self.main_sheets = pd.read_excel(self.main_file_path, sheet_name=None)
            
            logger.info("Loading all sheets from secondary Excel file...")
            self.secondary_sheets = pd.read_excel(self.secondary_file_path, sheet_name=None)
            
            # Log sheet information
            for sheet_name in self.main_sheets:
                df = self.main_sheets[sheet_name]
                logger.info(f"Main file sheet '{sheet_name}': {df.shape[0]} rows, {df.shape[1]} columns")
            
            for sheet_name in self.secondary_sheets:
                df = self.secondary_sheets[sheet_name]
                logger.info(f"Secondary file sheet '{sheet_name}': {df.shape[0]} rows, {df.shape[1]} columns")
            
            # Also set the main_df and secondary_df for backward compatibility (use first common sheet)
            if common_sheets:
                first_common_sheet = common_sheets[0]
                self.main_df = self.main_sheets[first_common_sheet]
                self.secondary_df = self.secondary_sheets[first_common_sheet]
            
            return True
            
        except Exception as e:
            logger.error(f"Error loading Excel files: {str(e)}")
            messagebox.showerror("Error", f"Error loading Excel files: {str(e)}")
            return False
    
    def align_columns(self, main_df: pd.DataFrame, secondary_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Align columns between two DataFrames to ensure they have the same structure.
        
        Args:
            main_df: Main DataFrame
            secondary_df: Secondary DataFrame
            
        Returns:
            Tuple of (aligned_main_df, aligned_secondary_df)
        """
        try:
            # Get all unique columns from both DataFrames
            main_cols = set(main_df.columns)
            secondary_cols = set(secondary_df.columns)
            all_cols = main_cols.union(secondary_cols)
            
            # Create copies to avoid modifying originals
            aligned_main = main_df.copy()
            aligned_secondary = secondary_df.copy()
            
            # Add missing columns to both DataFrames
            for col in all_cols:
                if col not in aligned_main.columns:
                    aligned_main[col] = None
                    logger.info(f"Added missing column '{col}' to main DataFrame")
                
                if col not in aligned_secondary.columns:
                    aligned_secondary[col] = None
                    logger.info(f"Added missing column '{col}' to secondary DataFrame")
            
            # Reorder columns to match
            column_order = sorted(all_cols)
            aligned_main = aligned_main[column_order]
            aligned_secondary = aligned_secondary[column_order]
            
            return aligned_main, aligned_secondary
            
        except Exception as e:
            logger.error(f"Error aligning columns: {str(e)}")
            return main_df, secondary_df
    
    def find_duplicates(self, main_df: pd.DataFrame, secondary_df: pd.DataFrame) -> pd.DataFrame:
        """
        Find duplicate rows between secondary and main DataFrames.
        
        Args:
            main_df: Main DataFrame
            secondary_df: Secondary DataFrame
            
        Returns:
            DataFrame containing non-duplicate rows from secondary file
        """
        try:
            logger.info("Checking for duplicates...")
            
            # Create a copy of secondary data for processing
            secondary_copy = secondary_df.copy()
            
            # Convert both DataFrames to string for comparison to handle mixed types
            main_str = main_df.astype(str)
            secondary_str = secondary_copy.astype(str)
            
            # Find duplicates by comparing all columns
            # Create a hash of each row for efficient comparison
            main_hashes = main_str.apply(lambda x: hash(tuple(x)), axis=1)
            secondary_hashes = secondary_str.apply(lambda x: hash(tuple(x)), axis=1)
            
            # Find rows in secondary that are NOT in main
            duplicate_mask = secondary_hashes.isin(main_hashes)
            non_duplicate_rows = secondary_copy[~duplicate_mask]
            
            duplicates_found = duplicate_mask.sum()
            unique_rows = len(non_duplicate_rows)
            
            logger.info(f"Found {duplicates_found} duplicate rows")
            logger.info(f"Found {unique_rows} unique rows to add")
            
            return non_duplicate_rows
            
        except Exception as e:
            logger.error(f"Error finding duplicates: {str(e)}")
            return pd.DataFrame()
    
    def merge_files(self) -> bool:
        """
        Merge the secondary file into the main file, avoiding duplicates.
        Handles multiple sheets if they exist.
        
        Returns:
            True if merge successful, False otherwise
        """
        try:
            # Load files
            if not self.load_excel_files():
                return False
            
            # Check if we have multiple sheets to process
            if self.main_sheets is not None and self.secondary_sheets is not None:
                # Multi-sheet processing
                main_sheets, secondary_sheets, common_sheets = self.get_sheet_info()
                
                if not common_sheets:
                    messagebox.showwarning("Warning", "No common sheet names found. Processing first sheet only.")
                    # Fall back to single sheet processing
                    return self._merge_single_file()
                
                # Initialize merged sheets dictionary
                self.merged_sheets = {}
                total_new_rows = 0
                
                # Process each common sheet
                for sheet_name in common_sheets:
                    if sheet_name in self.main_sheets and sheet_name in self.secondary_sheets:
                        main_df = self.main_sheets[sheet_name]
                        secondary_df = self.secondary_sheets[sheet_name]
                        
                        original_count = len(main_df)
                        merged_df = self.merge_single_sheet(sheet_name, main_df, secondary_df)
                        new_rows = len(merged_df) - original_count
                        total_new_rows += new_rows
                        
                        self.merged_sheets[sheet_name] = merged_df
                
                # Copy any sheets that are only in main file
                for sheet_name in self.main_sheets:
                    if sheet_name not in common_sheets:
                        self.merged_sheets[sheet_name] = self.main_sheets[sheet_name]
                        logger.info(f"Copied sheet '{sheet_name}' from main file (no matching sheet in secondary)")
                
                logger.info(f"Multi-sheet merge completed. Total new rows added: {total_new_rows}")
                
                # Update main_df for backward compatibility (use first common sheet)
                if common_sheets:
                    self.main_df = self.merged_sheets[common_sheets[0]]
                
                return True
            else:
                # Single sheet processing (fallback)
                return self._merge_single_file()
            
        except Exception as e:
            logger.error(f"Error during merge: {str(e)}")
            messagebox.showerror("Error", f"Error during merge: {str(e)}")
            return False
    
    def _merge_single_file(self) -> bool:
        """
        Helper method for single sheet merging (backward compatibility).
        
        Returns:
            True if merge successful, False otherwise
        """
        try:
            # Align columns
            self.main_df, self.secondary_df = self.align_columns(self.main_df, self.secondary_df)
            
            # Find non-duplicate rows
            unique_rows = self.find_duplicates(self.main_df, self.secondary_df)
            
            if len(unique_rows) == 0:
                messagebox.showinfo("Info", "No new unique rows found to add!")
                return True
            
            # Merge the data
            logger.info("Merging data...")
            merged_df = pd.concat([self.main_df, unique_rows], ignore_index=True)
            
            # Update the main DataFrame
            self.main_df = merged_df
            
            logger.info(f"Merge completed. Final dataset has {len(merged_df)} rows")
            return True
            
        except Exception as e:
            logger.error(f"Error during single file merge: {str(e)}")
            return False
    
    def save_merged_file(self, output_path: Optional[str] = None) -> bool:
        """
        Save the merged DataFrame(s) to an Excel file.
        Handles both single sheet and multi-sheet scenarios.
        
        Args:
            output_path: Optional custom output path. If None, overwrites the main file.
            
        Returns:
            True if save successful, False otherwise
        """
        try:
            # Determine output path
            if output_path is None:
                output_path = self.main_file_path
            
            # Create backup of original file
            if output_path == self.main_file_path:
                backup_path = str(Path(self.main_file_path).with_suffix('.backup.xlsx'))
                if os.path.exists(self.main_file_path):
                    # Create backup with all original sheets
                    original_sheets = pd.read_excel(self.main_file_path, sheet_name=None)
                    with pd.ExcelWriter(backup_path, engine='openpyxl') as writer:
                        for sheet_name, df in original_sheets.items():
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                    logger.info(f"Backup created: {backup_path}")
            
            # Save the merged file(s)
            logger.info(f"Saving merged file to: {output_path}")
            
            if self.merged_sheets is not None:
                # Multi-sheet save
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    for sheet_name, df in self.merged_sheets.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        logger.info(f"Saved sheet '{sheet_name}' with {len(df)} rows")
                
                messagebox.showinfo("Success", 
                    f"Multi-sheet file successfully saved to:\n{output_path}\n\n"
                    f"Sheets processed: {list(self.merged_sheets.keys())}")
            else:
                # Single sheet save (backward compatibility)
                if self.main_df is None:
                    raise ValueError("No data to save. Please merge files first.")
                
                self.main_df.to_excel(output_path, index=False)
                messagebox.showinfo("Success", f"File successfully saved to:\n{output_path}")
            
            logger.info("File saved successfully")
            return True
            
        except Exception as e:
            logger.error(f"Error saving file: {str(e)}")
            messagebox.showerror("Error", f"Error saving file: {str(e)}")
            return False
    
    def run_complete_merge(self) -> bool:
        """
        Run the complete merge process: select files, merge, and save.
        Handles both single sheet and multi-sheet scenarios.
        
        Returns:
            True if entire process successful, False otherwise
        """
        try:
            # Step 1: Select files
            main_file, secondary_file = self.select_files()
            if not main_file or not secondary_file:
                return False
            
            # Step 2: Merge files
            if not self.merge_files():
                return False
            
            # Step 3: Calculate merge statistics
            if self.merged_sheets is not None:
                # Multi-sheet statistics
                total_original_rows = 0
                total_final_rows = 0
                sheet_info = []
                
                for sheet_name in self.merged_sheets:
                    if sheet_name in self.main_sheets:
                        original_rows = len(self.main_sheets[sheet_name])
                        final_rows = len(self.merged_sheets[sheet_name])
                        new_rows = final_rows - original_rows
                        
                        total_original_rows += original_rows
                        total_final_rows += final_rows
                        
                        sheet_info.append(f"  • {sheet_name}: {original_rows} → {final_rows} (+{new_rows})")
                
                info_message = (
                    f"Multi-sheet merge completed successfully!\n\n"
                    f"Sheets processed: {len(self.merged_sheets)}\n"
                    f"Total original rows: {total_original_rows}\n"
                    f"Total rows after merge: {total_final_rows}\n"
                    f"Total new rows added: {total_final_rows - total_original_rows}\n\n"
                    f"Per-sheet breakdown:\n" + "\n".join(sheet_info) + "\n\n"
                    f"Do you want to save the merged file?\n"
                    f"(This will overwrite the main file)"
                )
            else:
                # Single sheet statistics
                original_rows = len(pd.read_excel(self.main_file_path))
                final_rows = len(self.main_df)
                new_rows = final_rows - original_rows
                
                info_message = (
                    f"Merge completed successfully!\n\n"
                    f"Original rows in main file: {original_rows}\n"
                    f"Rows after merge: {final_rows}\n"
                    f"New rows added: {new_rows}\n\n"
                    f"Do you want to save the merged file?\n"
                    f"(This will overwrite the main file)"
                )
            
            # Step 4: Ask user about saving
            save_choice = messagebox.askyesno("Save File", info_message)
            
            if save_choice:
                return self.save_merged_file()
            else:
                logger.info("User chose not to save the merged file")
                return True
                
        except Exception as e:
            logger.error(f"Error in complete merge process: {str(e)}")
            messagebox.showerror("Error", f"Error in complete merge process: {str(e)}")
            return False
        
    def get_sheet_info(self) -> Tuple[List[str], List[str], List[str]]:
        """
        Get sheet information from both Excel files.
        
        Returns:
            Tuple of (main_sheets, secondary_sheets, common_sheets)
        """
        try:
            if not self.main_file_path or not self.secondary_file_path:
                raise ValueError("File paths not set. Please select files first.")
            
            # Get sheet names from both files
            main_excel = pd.ExcelFile(self.main_file_path)
            secondary_excel = pd.ExcelFile(self.secondary_file_path)
            
            main_sheet_names = main_excel.sheet_names
            secondary_sheet_names = secondary_excel.sheet_names
            
            # Find common sheet names
            common_sheets = list(set(main_sheet_names).intersection(set(secondary_sheet_names)))
            
            logger.info(f"Main file sheets: {main_sheet_names}")
            logger.info(f"Secondary file sheets: {secondary_sheet_names}")
            logger.info(f"Common sheets found: {common_sheets}")
            
            main_excel.close()
            secondary_excel.close()
            
            return main_sheet_names, secondary_sheet_names, common_sheets
            
        except Exception as e:
            logger.error(f"Error getting sheet info: {str(e)}")
            messagebox.showerror("Error", f"Error getting sheet info: {str(e)}")
            return [], [], []
    
    def merge_single_sheet(self, sheet_name: str, main_df: pd.DataFrame, secondary_df: pd.DataFrame) -> pd.DataFrame:
        """
        Merge a single sheet from secondary file into main file, avoiding duplicates.
        
        Args:
            sheet_name: Name of the sheet being merged
            main_df: Main DataFrame for this sheet
            secondary_df: Secondary DataFrame for this sheet
            
        Returns:
            Merged DataFrame
        """
        try:
            logger.info(f"Processing sheet: {sheet_name}")
            
            # Align columns
            aligned_main, aligned_secondary = self.align_columns(main_df, secondary_df)
            
            # Find non-duplicate rows
            unique_rows = self.find_duplicates(aligned_main, aligned_secondary)
            
            if len(unique_rows) == 0:
                logger.info(f"No new unique rows found in sheet '{sheet_name}'")
                return aligned_main
            
            # Merge the data
            logger.info(f"Merging data for sheet '{sheet_name}'...")
            merged_df = pd.concat([aligned_main, unique_rows], ignore_index=True)
            
            logger.info(f"Sheet '{sheet_name}' merge completed. Rows: {len(aligned_main)} -> {len(merged_df)} (+{len(unique_rows)})")
            return merged_df
            
        except Exception as e:
            logger.error(f"Error merging sheet '{sheet_name}': {str(e)}")
            return main_df


def merge_excel_files():
    """
    Convenience function to run the complete Excel merge process.
    """
    merger = ExcelMerger()
    return merger.run_complete_merge()


if __name__ == "__main__":
    # Run the Excel merger
    print("🔄 Starting Excel File Merger...")
    print("📁 Please select your Excel files in the popup dialogs")
    print("📊 Multi-sheet support: Will automatically process all sheets with matching names")
    
    success = merge_excel_files()
    
    if success:
        print("✅ Excel merge process completed successfully!")
    else:
        print("❌ Excel merge process failed or was cancelled")
