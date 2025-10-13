"""
Demo script for Excel File Merger

This script demonstrates how to use the Excel merger module.
Run this script to merge two Excel files with duplicate checking.
"""

import sys
from pathlib import Path

# Add the src directory to the Python path
project_root = Path(__file__).parent
src_path = project_root / 'src'
sys.path.insert(0, str(src_path))

from utils.merge_excel_files import merge_excel_files, ExcelMerger


def main():
    """
    Main function to demonstrate Excel merging functionality.
    """
    print("=" * 60)
    print("🔄 EXCEL FILE MERGER DEMO")
    print("=" * 60)
    print()
    print("This tool will help you merge two Excel files:")
    print("1. 📁 Select a MAIN file (will be updated)")
    print("2. 📁 Select a SECONDARY file (data to add)")
    print("3. 🔍 Check for duplicates")
    print("4. ➕ Add only unique rows")
    print("5. 💾 Save the updated main file")
    print()
    print("Features:")
    print("✅ Automatic duplicate detection")
    print("✅ Column alignment")
    print("✅ Backup creation")
    print("✅ User-friendly GUI")
    print()
    
    # Option 1: Simple one-line merge
    print("🚀 Starting merge process...")
    print("📋 Please follow the popup dialogs to select your files")
    print()
    
    try:
        # Run the complete merge process
        success = merge_excel_files()
        
        if success:
            print("🎉 Merge completed successfully!")
        else:
            print("⚠️  Merge was cancelled or failed")
            
    except Exception as e:
        print(f"❌ Error during merge: {str(e)}")
    
    print()
    print("=" * 60)
    print("Demo completed. Thank you for using Excel File Merger!")
    print("=" * 60)


def advanced_demo():
    """
    Advanced demo showing step-by-step control of the merge process.
    """
    print("🔧 ADVANCED MERGE DEMO")
    print("=" * 40)
    
    # Create merger instance
    merger = ExcelMerger()
    
    try:
        # Step 1: Select files
        print("Step 1: Selecting files...")
        main_file, secondary_file = merger.select_files()
        
        if not main_file or not secondary_file:
            print("❌ File selection cancelled")
            return
        
        print(f"✅ Main file: {Path(main_file).name}")
        print(f"✅ Secondary file: {Path(secondary_file).name}")
        
        # Step 2: Load and process
        print("\nStep 2: Loading and processing files...")
        if merger.merge_files():
            print("✅ Files merged successfully")
            
            # Step 3: Show results
            print(f"\n📊 Results:")
            print(f"   Final rows: {len(merger.main_df)}")
            
            # Step 4: Save
            print("\nStep 3: Saving...")
            if merger.save_merged_file():
                print("✅ File saved successfully")
            else:
                print("❌ Save failed")
        else:
            print("❌ Merge failed")
            
    except Exception as e:
        print(f"❌ Error in advanced demo: {str(e)}")


if __name__ == "__main__":
    # Check if user wants advanced demo
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "--advanced":
        advanced_demo()
    else:
        main()
