import os
import pandas as pd
import json
import sys

# Configuration
SOURCE_DIR = r"c:\Users\jalpa\Desktop\atlas\ep-ppt\wetransfer_all-media-data_2025-11-20_0937\report data for all media"
OUTPUT_FILE = r"c:\Users\jalpa\Desktop\atlas\ep-ppt\output_data.json"
REQUIRED_COLS = ["SUPER CATEGORY", "PRODUCT GROUP"]

def normalize_cols(df):
    """Normalize column names to uppercase and stripped."""
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

def extract_data():
    aggregated_data = {} # Key: Super Category, Value: Set of Product Groups
    
    print(f"Scanning directory: {SOURCE_DIR}")
    
    files_processed = 0
    files_skipped = 0
    
    for root, dirs, files in os.walk(SOURCE_DIR):
        for file in files:
            if file.lower().endswith(".xlsx") and not file.startswith("~$"):
                file_path = os.path.join(root, file)
                print(f"Processing: {file}")
                try:
                    # Using engine='calamine' for extreme speed
                    # Using standard pandas.read_excel
                    df = pd.read_excel(file_path, engine='calamine')
                    
                    # Normalize headers
                    df = normalize_cols(df)
                    
                    # Check if required columns exist
                    if all(col in df.columns for col in REQUIRED_COLS):
                        # Drop rows where either column is NaN
                        # Make copy to avoid SettingWithCopyWarning
                        df_clean = df[REQUIRED_COLS].dropna().drop_duplicates()
                        
                        for _, row in df_clean.iterrows():
                            super_cat = str(row["SUPER CATEGORY"]).strip()
                            prod_group = str(row["PRODUCT GROUP"]).strip()
                            
                            if super_cat and prod_group and super_cat.upper() != "SUPER CATEGORY":
                                if super_cat not in aggregated_data:
                                    aggregated_data[super_cat] = set()
                                aggregated_data[super_cat].add(prod_group)
                        
                        files_processed += 1
                    else:
                        print(f"  [WARN] Missing required columns in {file}")
                        files_skipped += 1
                        
                except Exception as e:
                    print(f"  [ERROR] Processing {file}: {e}")
                    files_skipped += 1

    # Convert to list of dictionaries for JSON output
    output_list = []
    for super_cat, prod_groups in aggregated_data.items():
        output_list.append({
            "super_category": super_cat,
            "product_groups_array": sorted(list(prod_groups))
        })
    
    # Sort output for consistent results
    output_list.sort(key=lambda x: x["super_category"])

    # Write to JSON
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(output_list, f, indent=4, ensure_ascii=False)
    
    print("-" * 30)
    print(f"Processing Complete.")
    print(f"Files Processed: {files_processed}")
    print(f"Files Skipped: {files_skipped}")
    print(f"Unique Super Categories Found: {len(output_list)}")
    print(f"Output saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    try:
        extract_data()
    except KeyboardInterrupt:
        print("\nInterrupted.")
