#!/usr/bin/env python3
"""
Excel File Merger Module

This module provides functionality to merge two Excel files with duplicate detection and removal.
It includes a GUI for file selection and handles the merging process automatically.
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
        Load the selected Excel files into DataFrames.
        
        Returns:
            True if both files loaded successfully, False otherwise
        """
        try:
            if not self.main_file_path or not self.secondary_file_path:
                raise ValueError("File paths not set. Please select files first.")
            
            # Load main file
            logger.info("Loading main Excel file...")
            self.main_df = pd.read_excel(self.main_file_path)
            logger.info(f"Main file loaded: {self.main_df.shape[0]} rows, {self.main_df.shape[1]} columns")
            
            # Load secondary file
            logger.info("Loading secondary Excel file...")
            self.secondary_df = pd.read_excel(self.secondary_file_path)
            logger.info(f"Secondary file loaded: {self.secondary_df.shape[0]} rows, {self.secondary_df.shape[1]} columns")
            
            return True
            
        except Exception as e:
            logger.error(f"Error loading Excel files: {str(e)}")
            messagebox.showerror("Error", f"Error loading Excel files: {str(e)}")
            return False
    
    def align_columns(self) -> bool:
        """
        Align columns between the two DataFrames to ensure they have the same structure.
        
        Returns:
            True if alignment successful, False otherwise
        """
        try:
            # Get all unique columns from both DataFrames
            main_cols = set(self.main_df.columns)
            secondary_cols = set(self.secondary_df.columns)
            all_cols = main_cols.union(secondary_cols)
            
            # Add missing columns to both DataFrames
            for col in all_cols:
                if col not in self.main_df.columns:
                    self.main_df[col] = None
                    logger.info(f"Added missing column '{col}' to main file")
                
                if col not in self.secondary_df.columns:
                    self.secondary_df[col] = None
                    logger.info(f"Added missing column '{col}' to secondary file")
            
            # Reorder columns to match
            column_order = sorted(all_cols)
            self.main_df = self.main_df[column_order]
            self.secondary_df = self.secondary_df[column_order]
            
            logger.info("Column alignment completed")
            return True
            
        except Exception as e:
            logger.error(f"Error aligning columns: {str(e)}")
            return False
    
    def find_duplicates(self) -> pd.DataFrame:
        """
        Find duplicate rows between secondary and main DataFrames.
        
        Returns:
            DataFrame containing non-duplicate rows from secondary file
        """
        try:
            logger.info("Checking for duplicates...")
            
            # Create a copy of secondary data for processing
            secondary_copy = self.secondary_df.copy()
            
            # Convert both DataFrames to string for comparison to handle mixed types
            main_str = self.main_df.astype(str)
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
            messagebox.showerror("Error", f"Error finding duplicates: {str(e)}")
            return pd.DataFrame()
    
    def merge_files(self) -> bool:
        """
        Merge the secondary file into the main file, avoiding duplicates.
        
        Returns:
            True if merge successful, False otherwise
        """
        try:
            # Load files
            if not self.load_excel_files():
                return False
            
            # Align columns
            if not self.align_columns():
                return False
            
            # Find non-duplicate rows
            unique_rows = self.find_duplicates()
            
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
            logger.error(f"Error during merge: {str(e)}")
            messagebox.showerror("Error", f"Error during merge: {str(e)}")
            return False
    
    def save_merged_file(self, output_path: Optional[str] = None) -> bool:
        """
        Save the merged DataFrame to an Excel file.
        
        Args:
            output_path: Optional custom output path. If None, overwrites the main file.
            
        Returns:
            True if save successful, False otherwise
        """
        try:
            if self.main_df is None:
                raise ValueError("No data to save. Please merge files first.")
            
            # Determine output path
            if output_path is None:
                output_path = self.main_file_path
            
            # Create backup of original file
            if output_path == self.main_file_path:
                backup_path = str(Path(self.main_file_path).with_suffix('.backup.xlsx'))
                if os.path.exists(self.main_file_path):
                    pd.read_excel(self.main_file_path).to_excel(backup_path, index=False)
                    logger.info(f"Backup created: {backup_path}")
            
            # Save the merged file
            logger.info(f"Saving merged file to: {output_path}")
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
            
            # Step 3: Ask user about saving
            save_choice = messagebox.askyesno(
                "Save File", 
                f"Merge completed successfully!\n\n"
                f"Original rows in main file: {len(pd.read_excel(self.main_file_path))}\n"
                f"Rows after merge: {len(self.main_df)}\n"
                f"New rows added: {len(self.main_df) - len(pd.read_excel(self.main_file_path))}\n\n"
                f"Do you want to save the merged file?\n"
                f"(This will overwrite the main file)"
            )
            
            if save_choice:
                return self.save_merged_file()
            else:
                logger.info("User chose not to save the merged file")
                return True
                
        except Exception as e:
            logger.error(f"Error in complete merge process: {str(e)}")
            messagebox.showerror("Error", f"Error in complete merge process: {str(e)}")
            return False


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
    
    success = merge_excel_files()
    
    if success:
        print("✅ Excel merge process completed successfully!")
    else:
        print("❌ Excel merge process failed or was cancelled")
