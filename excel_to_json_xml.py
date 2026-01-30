import os
import json
import zipfile
import xml.etree.ElementTree as ET
import re

# Configuration
SOURCE_DIR = r"c:\Users\jalpa\Desktop\atlas\ep-ppt\wetransfer_all-media-data_2025-11-20_0937\report data for all media"
OUTPUT_FILE = r"c:\Users\jalpa\Desktop\atlas\ep-ppt\output_data.json"
REQUIRED_COLS = ["SUPER CATEGORY", "PRODUCT GROUP"]

def normalize(text):
    if text is None:
        return ""
    return str(text).strip().upper()

def strip_ns(tag):
    if '}' in tag:
        return tag.split('}', 1)[1]
    return tag

def get_shared_strings(zf):
    strings = []
    try:
        if "xl/sharedStrings.xml" in zf.namelist():
            with zf.open("xl/sharedStrings.xml") as f:
                tree = ET.parse(f)
                root = tree.getroot()
                # Iterate all 'si' elements regardless of namespace
                for si in root.iter():
                    if strip_ns(si.tag) == 'si':
                        # Find 't'
                        text_val = ""
                        # Check direct 't' child
                        for child in si:
                            if strip_ns(child.tag) == 't':
                                if child.text:
                                    text_val += child.text
                            elif strip_ns(child.tag) == 'r':
                                # Rich text
                                for r_child in child:
                                    if strip_ns(r_child.tag) == 't':
                                        if r_child.text:
                                            text_val += r_child.text
                        strings.append(text_val)
    except Exception as e:
        print(f"    Error reading sharedStrings: {e}")
    return strings

def col_letter_to_index(col_str):
    num = 0
    for c in col_str:
        num = num * 26 + (ord(c.upper()) - ord('A') + 1)
    return num - 1

def parse_sheet_rows(zf, sheet_path, shared_strings):
    with zf.open(sheet_path) as f:
        for event, elem in ET.iterparse(f, events=('end',)):
            if strip_ns(elem.tag) == 'row':
                row_data = []
                current_col_idx = 0
                
                # Iterate children directly to find cells 'c'
                for c in elem:
                    if strip_ns(c.tag) == 'c':
                        r_attr = c.get('r') # e.g. A1, might be missing
                        t_attr = c.get('t') # Type
                        
                        # Find value 'v'
                        val = None
                        for v_node in c:
                            if strip_ns(v_node.tag) == 'v':
                                val = v_node.text
                                break
                        
                        # Determine column index
                        col_idx = -1
                        if r_attr:
                            match = re.match(r"([A-Za-z]+)(\d+)", r_attr)
                            if match:
                                col_letter = match.group(1)
                                col_idx = col_letter_to_index(col_letter)
                        
                        if col_idx == -1:
                            col_idx = current_col_idx
                        else:
                            # If we jumped, fill gaps? 
                            # For now just use the index, but we yield a dense list later?
                            # Actually, simpler to just append if we don't care about empty explicit gaps at start
                            # But let's respect the col_idx if present
                            pass

                        current_col_idx = col_idx + 1

                        cell_value = ""
                        if t_attr == 'inlineStr':
                            # <is><t>...</t></is>
                            for is_node in c:
                                if strip_ns(is_node.tag) == 'is':
                                    for t_node in is_node:
                                        if strip_ns(t_node.tag) == 't':
                                            cell_value += (t_node.text or "")
                        elif val is not None:
                            if t_attr == 's': # Shared string
                                try:
                                    s_idx = int(val)
                                    if s_idx < len(shared_strings):
                                        cell_value = shared_strings[s_idx]
                                except:
                                    pass
                            elif t_attr == 'str': # Inline string
                                cell_value = val
                            else:
                                cell_value = val # Number/Other
                        
                        # Ensure row_data is long enough
                        while len(row_data) <= col_idx:
                            row_data.append("")
                        row_data[col_idx] = cell_value
                
                yield row_data
                elem.clear()

def process_file(file_path):
    local_data = {}
    if not zipfile.is_zipfile(file_path):
        return None

    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            shared_strings = get_shared_strings(zf)
            
            # Find sheet
            sheet_to_use = None
            # Prioritize standard names
            if "xl/worksheets/sheet1.xml" in zf.namelist():
                sheet_to_use = "xl/worksheets/sheet1.xml"
            else:
                for n in zf.namelist():
                    if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                        sheet_to_use = n
                        break
            
            if not sheet_to_use:
                return None

            rows_iter = parse_sheet_rows(zf, sheet_to_use, shared_strings)
            
            header_map = {}
            found_headers = False
            
            row_count = 0
            for row in rows_iter:
                row_count += 1
                # Removed header warning print
                    
                # Removed progress print for speed
                
                if not found_headers:
                    row_upper = [normalize(x) for x in row]
                    
                    if all(req in row_upper for req in REQUIRED_COLS):
                         # Found headers
                         for idx, val in enumerate(row_upper):
                             if val in REQUIRED_COLS:
                                 header_map[val] = idx
                         found_headers = True
                else:
                    if "SUPER CATEGORY" in header_map and "PRODUCT GROUP" in header_map:
                        sc_idx = header_map["SUPER CATEGORY"]
                        pg_idx = header_map["PRODUCT GROUP"]
                        
                        if len(row) > max(sc_idx, pg_idx):
                            sc = str(row[sc_idx]).strip()
                            pg = str(row[pg_idx]).strip()
                            
                            if sc and pg and normalize(sc) != "SUPER CATEGORY" and sc.lower() != "nan":
                                if sc not in local_data:
                                    local_data[sc] = set()
                                local_data[sc].add(pg)

    except Exception as e:
        print(f"Error processing {file_path}: {e}")

    return local_data

def extract_data():
    aggregated_data = {} 
    print(f"Scanning directory: {SOURCE_DIR}")
    
    files_processed = 0
    
    try:
        for root, dirs, files in os.walk(SOURCE_DIR):
            for file in files:
                if file.lower().endswith(".xlsx") and not file.startswith("~$"):
                    file_path = os.path.join(root, file)
                    print(f"Processing: {file}")
                    
                    data = process_file(file_path)
                    if data:
                        files_processed += 1
                        for sc, pgs in data.items():
                            if sc not in aggregated_data:
                                aggregated_data[sc] = set()
                            aggregated_data[sc].update(pgs)
                    else:
                        # It might be valid to have no data if file is empty or headers missing
                        pass

    except KeyboardInterrupt:
        print("\n[WARN] Interrupted by user. Saving data extracted so far...")
    except Exception as e:
        print(f"\n[ERROR] Unexpected error: {e}")
        
    # Output
    
    # Output
    output_list = []
    for super_cat, prod_groups in aggregated_data.items():
        output_list.append({
            "super_category": super_cat,
            "product_groups_array": sorted(list(prod_groups))
        })
    output_list.sort(key=lambda x: x["super_category"])

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(output_list, f, indent=4, ensure_ascii=False)
    
    print("-" * 30)
    print(f"Done. Processed {files_processed} files with data.")
    print(f"Total Super Categories: {len(output_list)}")
    print(f"Saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    extract_data()
