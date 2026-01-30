from config import SBB_TYPE, DEV_TYPE, SBB_VW_MODE, ON_OFF

def hex_string_to_list(hex_str: str) -> list:
    """Converts a space-separated hex string to a list of hex strings."""
    if not hex_str:
        return []
    return [x for x in hex_str.strip().split(' ') if x]

def hex_list_to_ascii(hex_list: list) -> str:
    """Converts a list of hex strings to ASCII string."""
    res = ""
    for h in hex_list:
        if h == "00":
            continue
        try:
            res += chr(int(h, 16))
        except ValueError:
            pass
    return res.strip()

def _is_sublist(sub: list, main: list) -> bool:
    n = len(sub)
    if n == 0: return True
    if n > len(main): return False
    for i in range(len(main) - n + 1):
        if main[i : i+n] == sub:
            return True
    return False

from typing import List, Optional, Union

def get_mdc_data(log_entries: list, search_pattern_str: str) -> Optional[list]:
    """
    Finds the LAST entry matching the search pattern in the log entries.
    Returns the DATA portion of the MDC response (as a list of hex strings).
    """
    pattern_parts = hex_string_to_list(search_pattern_str)
    
    for entry in reversed(log_entries):
        # Parse entry first to handle varying whitespace
        parts = hex_string_to_list(entry['resp'])
        
        if not _is_sublist(pattern_parts, parts):
            continue

        if len(parts) < 5:
            continue
        
        try:
            # Array: 0=AA, 1=FF, 2=ID, 3=LEN
            data_len = int(parts[3], 16)
            
            # Slicing for Data portion: starts at index 6
            # Ends at (Length + 3) because parts includes header
            start_idx = 6
            end_idx = data_len + 3
            
            if len(parts) > end_idx:
                return parts[start_idx : end_idx + 1]
            else:
                return parts[start_idx:]
        except (ValueError, IndexError):
            continue
                
    return None

def get_mdc_ascii(log_entries: list, search_pattern: str) -> str:
    """Wrapper to get ASCII string from MDC data."""
    data = get_mdc_data(log_entries, search_pattern)
    if not data:
        return "NULL"
    return hex_list_to_ascii(data)
