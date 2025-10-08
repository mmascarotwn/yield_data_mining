#!/usr/bin/env python3
"""
Demo script for the Yield Calculator Module

This script demonstrates how to use the yield_calculator module to add
e_yield and asm_yield columns to Excel files.
"""

import sys
import os
from pathlib import Path

# Add the src directory to the Python path
current_dir = Path(__file__).parent
src_dir = current_dir / "src"
sys.path.insert(0, str(src_dir))

from utils.yield_calculator import add_yield_columns

def main():
    """
    Run the yield calculator demo.
    """
    print("=" * 60)
    print("📊 YIELD CALCULATOR DEMO")
    print("=" * 60)
    print()
    print("This demo will help you add 'e_yield' and 'asm_yield' columns")
    print("to your Excel files.")
    print()
    print("Calculation Methods Available:")
    print("  • 'default' - Sets all values to 0.0")
    print("  • 'column_name' - Copies values from an existing column")
    print("  • 'formula' - Uses a pandas expression for calculation")
    print()
    print("Example formulas:")
    print("  • 'pass_count / total_count * 100' (for percentage)")
    print("  • 'good_units / tested_units * 100'")
    print("  • 'yield_value * 0.95' (for adjusted yield)")
    print()
    print("Starting the yield calculator...")
    print("-" * 60)
    
    # Run the yield calculator
    success = add_yield_columns()
    
    print("-" * 60)
    if success:
        print("✅ Demo completed successfully!")
        print("🎉 Your Excel file now has e_yield and asm_yield columns!")
    else:
        print("❌ Demo was cancelled or encountered an error.")
    print("=" * 60)

if __name__ == "__main__":
    main()
