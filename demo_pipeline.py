#!/usr/bin/env python3
"""
Integration Demo - Complete Data Processing Pipeline

This script demonstrates how to use both the excel_merger and yield_calculator
modules together to create a complete data processing pipeline.
"""

import sys
import os
from pathlib import Path

# Add the src directory to the Python path
current_dir = Path(__file__).parent
src_dir = current_dir / "src"
sys.path.insert(0, str(src_dir))

from utils.yt_merge_tables import merge_excel_files
from utils.dint_yield_calculator import add_yield_columns

def main():
    """
    Run the complete data processing pipeline.
    """
    print("=" * 70)
    print("🔄 COMPLETE DATA PROCESSING PIPELINE")
    print("=" * 70)
    print()
    print("This pipeline will:")
    print("  1. Merge two Excel files (with duplicate detection)")
    print("  2. Add e_yield and asm_yield columns to the merged file")
    print()
    print("Benefits:")
    print("  ✅ Multi-sheet support for both operations")
    print("  ✅ Automatic duplicate detection and removal")
    print("  ✅ Flexible yield calculation methods")
    print("  ✅ Comprehensive backup and logging")
    print("  ✅ User-friendly GUI interface")
    print()
    
    # Step 1: Excel File Merging
    print("🔄 STEP 1: Excel File Merging")
    print("-" * 35)
    print("Select your main Excel file and secondary Excel file to merge...")
    
    merge_success = merge_excel_files()
    
    if not merge_success:
        print("❌ Excel merging failed or was cancelled.")
        print("Pipeline stopped.")
        return False
    
    print("✅ Excel merging completed successfully!")
    print()
    
    # Ask user if they want to continue with yield calculation
    import tkinter as tk
    from tkinter import messagebox
    
    root = tk.Tk()
    root.withdraw()
    
    continue_choice = messagebox.askyesno(
        "Continue Pipeline",
        "Excel merging completed successfully!\n\n"
        "Do you want to continue with adding yield columns?\n\n"
        "You'll be able to select the merged file (or any other Excel file) "
        "to add e_yield and asm_yield columns."
    )
    
    root.destroy()
    
    if not continue_choice:
        print("Pipeline stopped at user request.")
        return True
    
    # Step 2: Yield Column Addition
    print("🔄 STEP 2: Yield Column Addition")
    print("-" * 35)
    print("Select the Excel file to add yield columns to...")
    print("(This can be the merged file from Step 1 or any other Excel file)")
    
    yield_success = add_yield_columns()
    
    if not yield_success:
        print("❌ Yield calculation failed or was cancelled.")
        return False
    
    print("✅ Yield calculation completed successfully!")
    print()
    
    # Pipeline completion
    print("🎉 PIPELINE COMPLETED SUCCESSFULLY!")
    print("-" * 35)
    print("Your data has been:")
    print("  ✅ Merged (with duplicate removal)")
    print("  ✅ Enhanced with yield calculations")
    print("  ✅ Backed up for safety")
    print()
    print("Ready for further analysis! 📊")
    
    return True

def run_individual_tools():
    """
    Allow user to run individual tools instead of the complete pipeline.
    """
    import tkinter as tk
    from tkinter import messagebox
    
    root = tk.Tk()
    root.withdraw()
    
    choice = messagebox.askyesnocancel(
        "Tool Selection",
        "Choose which tool to run:\n\n"
        "• YES: Excel File Merger only\n"
        "• NO: Yield Calculator only\n" 
        "• CANCEL: Complete Pipeline"
    )
    
    root.destroy()
    
    if choice is True:
        print("🔄 Running Excel File Merger...")
        return merge_excel_files()
    elif choice is False:
        print("📊 Running Yield Calculator...")
        return add_yield_columns()
    else:
        print("🔄 Running Complete Pipeline...")
        return main()

if __name__ == "__main__":
    print("=" * 70)
    print("🚀 DATA PROCESSING TOOLKIT")
    print("=" * 70)
    print()
    
    # Ask user what they want to do
    import tkinter as tk
    from tkinter import messagebox
    
    root = tk.Tk()
    root.withdraw()
    
    pipeline_choice = messagebox.askyesno(
        "Processing Options",
        "Welcome to the Data Processing Toolkit!\n\n"
        "Choose your processing option:\n\n"
        "• YES: Run Complete Pipeline (Merge + Yield Calculation)\n"
        "• NO: Run Individual Tools\n\n"
        "Recommendation: Use Complete Pipeline for new data processing."
    )
    
    root.destroy()
    
    if pipeline_choice:
        success = main()
    else:
        success = run_individual_tools()
    
    print("=" * 70)
    if success:
        print("🎉 Processing completed successfully!")
    else:
        print("❌ Processing was cancelled or encountered an error.")
    print("=" * 70)
