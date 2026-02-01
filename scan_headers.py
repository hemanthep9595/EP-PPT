
import os
import zipfile
import xml.etree.ElementTree as ET
import re

SOURCE_DIR = r"c:\Users\jalpa\Desktop\atlas\ep-ppt\wetransfer_all-media-data_2025-11-20_0937\report data for all media"

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
    except:
        pass
    return strings

def parse_first_row(zf, sheet_path, shared_strings):
    with zf.open(sheet_path) as f:
        for event, elem in ET.iterparse(f, events=('end',)):
            if strip_ns(elem.tag) == 'row':
                row_data = []
                current_col_idx = 0
                for c in elem:
                    if strip_ns(c.tag) == 'c':
                        t_attr = c.get('t')
                        val = None
                        for v_node in c:
                            if strip_ns(v_node.tag) == 'v':
                                val = v_node.text
                                break
                        
                        r_attr = c.get('r')
                        col_idx = -1
                        if r_attr:
                            match = re.match(r"([A-Za-z]+)(\d+)", r_attr)
                            if match:
                                from string import ascii_uppercase
                                # Simple converter for now, assuming single/double letters
                                col_letter = match.group(1).upper()
                                # Quick hack for AA/AB etc if needed, but headers usually few
                                num = 0
                                for char in col_letter:
                                    num = num * 26 + (ord(char) - ord('A') + 1)
                                col_idx = num - 1
                        
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
                
                return [normalize(x) for x in row_data]
    return []

def scan_headers():
    seen_headers = set()
    print(f"Scanning headers in: {SOURCE_DIR}")
    
    results = {}
    
    for root, dirs, files in os.walk(SOURCE_DIR):
        files_in_dir = [f for f in files if f.lower().endswith(".xlsx") and not f.startswith("~$")]
        for file in files_in_dir[:3]:
            file_path = os.path.join(root, file)
            try:
                if not zipfile.is_zipfile(file_path): continue
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
                    
                    if sheet_to_use:
                        headers = parse_first_row(zf, sheet_to_use, shared_strings)
                        results[file] = headers
                        for h in headers: seen_headers.add(h)
            except: pass
            
    import json
    with open("c:\\Users\\jalpa\\Desktop\\atlas\\ep-ppt\\headers_dump.json", "w", encoding='utf-8') as f:
        json.dump({"files": results, "unique_headers": sorted(list(seen_headers))}, f, indent=4)

if __name__ == "__main__":
    scan_headers()
