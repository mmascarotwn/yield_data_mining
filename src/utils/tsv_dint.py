#!/usr/bin/env python3

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select a CSV file",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    root.destroy()
    return file_path

def fill_missing_source_columns(df):
    for col in [
        'TSV_RES_DIG_DRAM0_1_2_3_TOTAL_FAILS',
        'TSV_RES_DIG_DRAM4_5_6_7_TOTAL_FAILS',
        'TSV_RES_DIG_DRAM8_9_10_11_TOTAL_FAILS'
    ]:
        if col in df.columns:
            df[col] = df[col].fillna(0)
    return df

def add_D0_3_column(df):
    src, tgt = 'TSV_RES_DIG_DRAM0_1_2_3_TOTAL_FAILS', 'D0_3_TSV_RES_DIG'
    if tgt not in df.columns:
        df[tgt] = df[src]
    else:
        df.loc[df[tgt].isna(), tgt] = df.loc[df[tgt].isna(), src]
    return df

def add_D4_7_column(df):
    src1, src2, tgt = 'TSV_RES_DIG_DRAM4_5_6_7_TOTAL_FAILS', 'TSV_RES_DIG_DRAM0_1_2_3_TOTAL_FAILS', 'D4_7_TSV_RES_DIG'
    if tgt not in df.columns:
        df[tgt] = df[src1] - df[src2]
    else:
        mask = df[tgt].isna()
        df.loc[mask, tgt] = df.loc[mask, src1] - df.loc[mask, src2]
    return df

def add_D8_11_column(df):
    src1, src2, tgt = 'TSV_RES_DIG_DRAM8_9_10_11_TOTAL_FAILS', 'TSV_RES_DIG_DRAM4_5_6_7_TOTAL_FAILS', 'D8_11_TSV_RES_DIG'
    if tgt not in df.columns:
        df[tgt] = df[src1] - df[src2]
    else:
        mask = df[tgt].isna()
        df.loc[mask, tgt] = df.loc[mask, src1] - df.loc[mask, src2]
    return df

def add_fid(df):
    def format_row(row_val):
        try:
            val = int(row_val)
            prefix = 'P' if val >= 0 else 'N'
            return f"{prefix}{abs(val):02d}"
        except:
            return 'P00'

    def format_col(col_val):
        try:
            val = int(col_val)
            return f"{val:02d}"
        except:
            return '00'

    if all(col in df.columns for col in ['LOT', 'WAFER', 'ROW', 'COL']):
        df['FID'] = df.apply(
            lambda row: f"{row['LOT']}:{row['WAFER']}:{format_row(row['ROW'])}:{format_col(row['COL'])}", axis=1
        )
    return df

def add_retest_column(df):
    def extract_retest(log_str):
        try:
            parts = log_str.split('.')
            if len(parts) > 5:
                return parts[5]
            return ''
        except:
            return ''
    
    if 'LOG' in df.columns:
        df['RETEST'] = df['LOG'].apply(extract_retest)
    return df

def flag_latest_retest(df):
    if all(col in df.columns for col in ['FID', 'Process_id', 'RETEST']):
        df['RETEST'] = pd.to_numeric(df['RETEST'], errors='coerce')
        idx = df.groupby(['FID', 'Process_id'])['RETEST'].idxmax()
        df['latest_retest'] = ''
        df.loc[idx, 'latest_retest'] = 'last'
    return df

def process_file(file_path):
    df = pd.read_csv(file_path)
    df = fill_missing_source_columns(df)
    df = add_D0_3_column(df)
    df = add_D4_7_column(df)
    df = add_D8_11_column(df)
    df = add_fid (df)
    df = add_retest_column(df)
    df = flag_latest_retest(df)
    df.to_csv(file_path, index=False)
    return file_path

def main():
    file_path = select_file()
    if not file_path:
        print("No file selected.")
        return
    try:
        processed_path = process_file(file_path)
        messagebox.showinfo("Success", f"File processed and saved:\n{processed_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    main()