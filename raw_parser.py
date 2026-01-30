import zipfile
import xml.etree.ElementTree as ET
import traceback

def read_xlsx_raw(filename: str) -> list:
    """
    Fallback parser for detecting data when openpyxl fails due to style errors.
    Parses the internal XML directly associated with sheet1.
    """
    print("Attempting raw XML parsing...")
    try:
        with zipfile.ZipFile(filename, 'r') as z:
            # 1. Parse Shared Strings
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    # namespace usually: {http://schemas.openxmlformats.org/spreadsheetml/2006/main}sst
                    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    for si in root.findall('ns:si', ns):
                        t = si.find('ns:t', ns)
                        if t is not None:
                            shared_strings.append(t.text)
                        else:
                            # Handle rich text or other complications simply
                            shared_strings.append("")

            # 2. Parse Sheet1
            rows_list = []
            sheet_path = 'xl/worksheets/sheet1.xml'
            if sheet_path not in z.namelist():
                # Try finding it if name is different? usually sheet1.xml is standard for first sheet
                # but let's stick to standard for now.
                print("Could not find xl/worksheets/sheet1.xml")
                return []

            with z.open(sheet_path) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                
                sheet_data = root.find('ns:sheetData', ns)
                if sheet_data is None:
                    return []
                    
                for row in sheet_data.findall('ns:row', ns):
                    # Value extraction logic...
                    # Store by column index
                    cells = {}
                    max_col = -1
                    current_col = 0 # 0-based tracking
                    
                    for c in row.findall('ns:c', ns):
                        r_ref = c.get('r') # e.g. "A1"
                        
                        col_idx = -1
                        if r_ref:
                            # Parse column index from A1, B1, AA1 etc.
                            col_str = "".join(filter(str.isalpha, r_ref))
                            idx = 0
                            for char in col_str:
                                idx = idx * 26 + (ord(char.upper()) - ord('A')) + 1
                            col_idx = idx - 1 # 0-based
                            current_col = col_idx
                        else:
                            col_idx = current_col
                        
                        t = c.get('t') # type
                        
                        # Value extraction
                        val = ""
                        if t == 's': # shared string
                            v = c.find('ns:v', ns)
                            if v is not None and v.text:
                                try:
                                    val = shared_strings[int(v.text)]
                                except: pass
                        elif t == 'inlineStr': # inline string
                            is_elem = c.find('ns:is', ns)
                            if is_elem is not None:
                                t_elem = is_elem.find('ns:t', ns)
                                if t_elem is not None:
                                    val = t_elem.text
                        else: # number or string formula result
                            v = c.find('ns:v', ns)
                            if v is not None:
                                val = v.text
                        
                        if val is None: val = ""
                        
                        cells[col_idx] = val
                        if col_idx > max_col:
                            max_col = col_idx
                            
                        # Move to next column for next iteration if r is missing
                        current_col += 1
                        
                    # Construct list
                    row_list = [''] * (max_col + 1)
                    for i, v in cells.items():
                        row_list[i] = v
                        
                    rows_list.append(row_list)
            
            return rows_list
            
    except Exception as e:
        print(f"Raw XML parsing failed: {e}")
        traceback.print_exc()
        return []
