import zipfile
import os

FILE_PATH = r"c:\Users\jalpa\Desktop\atlas\ep-ppt\wetransfer_all-media-data_2025-11-20_0937\report data for all media\Digital\Aug -25.xlsx"

print(f"Inspecting File: {FILE_PATH}")

try:
    with zipfile.ZipFile(FILE_PATH, 'r') as zf:
        print("\n--- Files in Zip ---")
        for name in zf.namelist():
            print(name)
            
        print("\n--- SharedStrings Content (First 500 chars) ---")
        if "xl/sharedStrings.xml" in zf.namelist():
            with zf.open("xl/sharedStrings.xml") as f:
                print(f.read(500).decode('utf-8', errors='ignore'))
        else:
            print("No sharedStrings.xml found!")

        print("\n--- Sheet 1 Content (First 1000 chars) ---")
        # Try to find the first sheet
        sheets = [n for n in zf.namelist() if n.startswith("xl/worksheets/sheet")]
        if sheets:
            with zf.open(sheets[0]) as f:
                print(f"Reading {sheets[0]}:")
                print(f.read(1000).decode('utf-8', errors='ignore'))
        else:
            print("No sheets found!")

except Exception as e:
    print(f"Error: {e}")
