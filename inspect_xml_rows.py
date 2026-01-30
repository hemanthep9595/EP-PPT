import zipfile
import xml.etree.ElementTree as ET
import re

FILE_PATH = r"c:\Users\jalpa\Desktop\atlas\ep-ppt\wetransfer_all-media-data_2025-11-20_0937\report data for all media\Digital\Aug -25.xlsx"

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

def inspect_rows():
    print(f"Inspecting File: {FILE_PATH}")
    with zipfile.ZipFile(FILE_PATH, 'r') as zf:
        shared_strings = get_shared_strings(zf)
        print(f"Found {len(shared_strings)} shared strings.")
        
        sheet_path = "xl/worksheets/sheet1.xml"
        if sheet_path not in zf.namelist():
             # find first
             for n in zf.namelist():
                 if n.startswith("xl/worksheets/sheet"):
                     sheet_path = n
                     break
        
        print(f"Reading Sheet: {sheet_path}")
        
        with zf.open(sheet_path) as f:
            for event, elem in ET.iterparse(f, events=('end',)):
                if strip_ns(elem.tag) == 'row':
                    row_data = []
                    for c in elem:
                        if strip_ns(c.tag) == 'c':
                            t_attr = c.get('t')
                            val = None
                            for v in c:
                                if strip_ns(v.tag) == 'v':
                                    val = v.text
                                    break
                            
                            c_val = ""
                            if val:
                                if t_attr == 's':
                                    try:
                                        c_val = shared_strings[int(val)]
                                    except:
                                        c_val = f"<err:{val}>"
                                elif t_attr == 'str':
                                    c_val = val
                                else:
                                    c_val = val
                            
                            # Check inlineStr
                            if t_attr == 'inlineStr':
                                for is_node in c:
                                     if strip_ns(is_node.tag) == 'is':
                                         for t_node in is_node:
                                             if strip_ns(t_node.tag) == 't':
                                                 c_val += (t_node.text or "")

                            row_data.append(c_val)
                            
                    if len(row_data) > 0:
                        # Print only if potential header
                        row_str = str(row_data).upper()
                        if 'CATEGORY' in row_str or 'GROUP' in row_str:
                            print(f"ROW {elem.get('r')}: {row_data}")
                        
                    # Stop after 100 rows
                    if int(elem.get('r') or 0) > 100: 
                         break

if __name__ == "__main__":
    inspect_rows()
