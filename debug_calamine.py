
from python_calamine import CalamineWorkbook
import sys

file_path = r"c:\Users\jalpa\Desktop\atlas\ep-ppt\wetransfer_all-media-data_2025-11-20_0937\report data for all media\Print\1. Aug-25.xlsx"

print(f"Testing direct calamine on: {file_path}")

try:
    wb = CalamineWorkbook.from_path(file_path)
    print(f"Workbook loaded. Sheets: {wb.sheet_names}")
    
    # Try reading the first sheet
    sheet_name = wb.sheet_names[0]
    rows = wb.get_sheet_by_name(sheet_name).to_python()
    print(f"Read {len(rows)} rows from {sheet_name}")
    print(f"Row 1: {rows[0]}")
    
except Exception as e:
    print(f"Error: {e}")
