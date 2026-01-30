
import pandas as pd
import os

file_path = r"c:\Users\jalpa\Desktop\atlas\ep-ppt\wetransfer_all-media-data_2025-11-20_0937\report data for all media\Print\1. Aug-25.xlsx"

try:
    # Read only the first few rows to get headers
    df = pd.read_excel(file_path, nrows=5)
    print("Columns found:")
    print(df.columns.tolist())
    
    # Check for specific columns
    required_cols = ["SUPER CATEGORY", "PRODUCT GROUP"]
    missing = [col for col in required_cols if col not in df.columns]
    
    if missing:
        print(f"MISSING COLUMNS: {missing}")
    else:
        print("Required columns found.")
        print("Sample data:")
        print(df[required_cols].head())

except Exception as e:
    print(f"Error reading file: {e}")
