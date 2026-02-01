
import os
import json
import zipfile
import xml.etree.ElementTree as ET
import re
import multiprocessing

# Configuration
SOURCE_DIR = r"c:\Users\jalpa\Desktop\atlas\ep-ppt\wetransfer_all-media-data_2025-11-20_0937\report data for all media"
OUTPUT_FILE = r"c:\Users\jalpa\Desktop\atlas\ep-ppt\location_data.json"
# Target columns
REQUIRED_COLS = ["STATE", "CITY"]

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
                for si in root.iter():
                    if strip_ns(si.tag) == 'si':
                        text_val = ""
                        for child in si:
                            if strip_ns(child.tag) == 't':
                                if child.text:
                                    text_val += child.text
                            elif strip_ns(child.tag) == 'r':
                                for r_child in child:
                                    if strip_ns(r_child.tag) == 't':
                                        if r_child.text:
                                            text_val += r_child.text
                        strings.append(text_val)
    except: pass
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
                for c in elem:
                    if strip_ns(c.tag) == 'c':
                        r_attr = c.get('r')
                        t_attr = c.get('t')
                        val = None
                        for v_node in c:
                            if strip_ns(v_node.tag) == 'v':
                                val = v_node.text
                                break
                        col_idx = -1
                        if r_attr:
                            match = re.match(r"([A-Za-z]+)(\d+)", r_attr)
                            if match:
                                col_letter = match.group(1)
                                col_idx = col_letter_to_index(col_letter)
                        if col_idx == -1: col_idx = current_col_idx
                        current_col_idx = col_idx + 1

                        cell_value = ""
                        if t_attr == 'inlineStr':
                            for is_node in c:
                                if strip_ns(is_node.tag) == 'is':
                                    for t_node in is_node:
                                        if strip_ns(t_node.tag) == 't':
                                            cell_value += (t_node.text or "")
                        elif val is not None:
                            if t_attr == 's':
                                try:
                                    s_idx = int(val)
                                    if s_idx < len(shared_strings):
                                        cell_value = shared_strings[s_idx]
                                except: pass
                            else:
                                cell_value = val
                        while len(row_data) <= col_idx:
                            row_data.append("")
                        row_data[col_idx] = cell_value
                yield row_data
                elem.clear()

def process_file_worker(file_path):
    results = set()
    try:
        if not zipfile.is_zipfile(file_path): return []
        with zipfile.ZipFile(file_path, 'r') as zf:
            shared_strings = get_shared_strings(zf)
            sheet_to_use = None
            if "xl/worksheets/sheet1.xml" in zf.namelist():
                sheet_to_use = "xl/worksheets/sheet1.xml"
            else:
                for n in zf.namelist():
                    if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                        sheet_to_use = n
                        break
            if not sheet_to_use: return []

            rows_iter = parse_sheet_rows(zf, sheet_to_use, shared_strings)
            header_map = {}
            found_headers = False
            
            for row in rows_iter:
                if not found_headers:
                    row_upper = [normalize(x) for x in row]
                    if all(req in row_upper for req in REQUIRED_COLS):
                         for idx, val in enumerate(row_upper):
                             if val in REQUIRED_COLS:
                                 header_map[val] = idx
                         found_headers = True
                else:
                    if "STATE" in header_map and "CITY" in header_map:
                        st_idx = header_map["STATE"]
                        ct_idx = header_map["CITY"]
                        if len(row) > max(st_idx, ct_idx):
                            state = str(row[st_idx]).strip()
                            city = str(row[ct_idx]).strip()
                            if state and city and normalize(state) != "STATE" and state.lower() != "nan":
                                results.add((state, city))
    except: pass
    return list(results)

def main():
    print(f"Scanning directory: {SOURCE_DIR}")
    file_paths = []
    for root, dirs, files in os.walk(SOURCE_DIR):
        for file in files:
            if file.lower().endswith(".xlsx") and not file.startswith("~$"):
                file_paths.append(os.path.join(root, file))
    
    print(f"Found {len(file_paths)} files. Starting multiprocessing for Location Data...")
    
    aggregated_data = {}
    cpu_count = max(1, multiprocessing.cpu_count() - 1)
    
    with multiprocessing.Pool(cpu_count) as pool:
        all_results = pool.map(process_file_worker, file_paths)
        for file_res in all_results:
            if not file_res: continue
            for state, city in file_res:
                if state and city:
                    if state not in aggregated_data:
                        aggregated_data[state] = set()
                    aggregated_data[state].add(city)

    output_list = []
    for state, cities in aggregated_data.items():
        output_list.append({
            "state": state,
            "cities_array": sorted(list(cities))
        })
    output_list.sort(key=lambda x: x["state"])

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(output_list, f, indent=4, ensure_ascii=False)
    
    print("-" * 30)
    print(f"Processing Complete.")
    print(f"Total States Found: {len(output_list)}")
    print(f"Saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
