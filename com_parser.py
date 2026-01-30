import os
import traceback
import logging

try:
    import win32com.client
    HAS_COM = True
except ImportError:
    HAS_COM = False

logger = logging.getLogger("LSMParser.COM")

def read_xlsx_via_com(file_path):
    """
    Uses Microsoft Excel via COM to read a spreadsheet.
    This is used as a fallback for files with DRM (NASCA) or where openpyxl fails.
    Returns: list of lists (rows of cells)
    """
    if not HAS_COM:
        logger.warning("pywin32 not installed, cannot use COM fallback.")
        return None
        
    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        return None

    excel = None
    wb = None
    data = []

    print("Attempting to read via Excel COM...")
    logger.info("Attempting to read via Excel COM...")

    try:
        # 1. Initialize Excel
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
        except Exception as e:
            logger.error(f"Failed to initialize Excel COM: {e}")
            print("Error: Could not start Excel. Is it installed?")
            return None

        # 2. Open Workbook
        try:
            wb = excel.Workbooks.Open(abs_path, ReadOnly=True)
        except Exception as e:
             logger.error(f"Excel COM could not open file: {e}")
             print(f"Excel could not open the file. Is it locked or corrupted?")
             return None
        
        # 3. Read Data
        try:
            ws = wb.Worksheets(1) # 1-based index
            
            # Efficiently read UsedRange
            used_range = ws.UsedRange
            # .Value returns a tuple of tuples
            raw_data = used_range.Value
            
            # Convert to list of lists and handle potential None for empty cells if necessary
            # (Though tuple of tuples is fine for iteration usually)
            
            if raw_data:
                # If it's a single cell, it returns a value, not a tuple
                if not isinstance(raw_data, tuple):
                     data = [[raw_data]]
                else:
                     # Ensure it's a list of lists for consistency with other parsers
                     # data = [list(row) for row in raw_data] 
                     # Actually, raw_data is a tuple of tuples.
                     # But unexpected things happen if UsedRange is empty or 1x1.
                     pass 
                     
                data = raw_data
                
            logger.info(f"Excel COM read {len(data) if data else 0} rows.")
            
        except Exception as e:
            logger.error(f"Error reading data from worksheet: {e}")
            return None

    except Exception as e:
        logger.error(f"Unexpected COM error: {e}")
        traceback.print_exc()
        return None
        
    finally:
        # Cleanup
        try:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except Exception as e:
            logger.warning(f"Error closing Excel COM: {e}")
            
    return data
