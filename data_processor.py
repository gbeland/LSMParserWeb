from typing import List, Dict, Tuple, Optional, Union
from config import (
    SBB_TYPE, SBB_VW_MODE, ON_OFF, INPUT_SOURCES, POWER_STATUS, STATUS_CODES, NETWORK_MODES,
    CMD_MODEL_NAME, CMD_SERIAL_NUM, CMD_FW_MAIN, CMD_DEVICE_NAME, CMD_VW_MODE, 
    CMD_MAC_ADDR, CMD_GROUP_IP, CMD_IP_MODE, CMD_IMEI, CMD_INPUT_SOURCE, CMD_POWER_STATUS, 
    CMD_STATUS_CODE, CMD_LAYOUT, CMD_TEMP, CMD_BACKLIGHT, CMD_CC_CAB, CMD_CC_MOD, 
    CMD_CC_PIX, CMD_SEAM_COR
)
from mdc_parser import get_mdc_ascii, get_mdc_data, hex_list_to_ascii

def process_logs(sbb_list: dict, cab_list: dict) -> tuple[list, list, dict]:
    """
    Processes the raw logs and returns structured data for reports.
    Returns: (sbb_data, cab_data, layout_data)
    """
    
    sbb_data = [] # List of dicts
    cab_data = [] # List of dicts
    layout_data = {} # SBB_IP -> list of cab layout dicts
    
    # helper for group lookup
    group_lookup = [] # List of {'sbb_ip', 'group_num', 'group_ip'}

    # 1. Process SBBs
    # Sort for consistent ordering, maybe?
    # sbb_list keys are (ip, id). Sort by IP.
    sorted_sbbs = sorted(sbb_list.items(), key=lambda x: x[0][0])
    
    for idx, ((sbb_ip, sbb_id), logs) in enumerate(sorted_sbbs):
        sbb_entry = {}
        sbb_entry['meta_col_idx'] = idx + 2 # For excel col mapping if needed
        sbb_entry['name_header'] = f"SBB {idx + 1}"
        sbb_entry['ip'] = sbb_ip
        
        # Model
        model_name = get_mdc_ascii(logs, CMD_MODEL_NAME)
        if model_name == "NULL" or not model_name:
            fw_ver = get_mdc_ascii(logs, CMD_FW_MAIN)
            if fw_ver and fw_ver != "NULL":
                short_fw = fw_ver[:6].replace(" ", "")
                found_type = "Unknown"
                for k, v in SBB_TYPE.items():
                    if k in short_fw:
                        found_type = v
                        break
                model_name = found_type
        sbb_entry['model'] = model_name
        
        sbb_entry['sn'] = get_mdc_ascii(logs, CMD_SERIAL_NUM)
        sbb_entry['fw_main'] = get_mdc_ascii(logs, CMD_FW_MAIN)
        sbb_entry['fw_add'] = _parse_add_versions(logs, model_name)
        sbb_entry['sbb_name'] = get_mdc_ascii(logs, CMD_DEVICE_NAME)
        
        # Groups
        gips = _parse_group_ips(logs)
        sbb_entry['groups'] = gips # dict {1: ip, 2: ip...}
        for g_num, ip_str in gips.items():
            if ip_str != "0.0.0.0":
                group_lookup.append({
                    'sbb_ip': sbb_ip,
                    'group_num': g_num,
                    'group_ip': ip_str
                })
                
        # VW Mode
        sbb_entry['vw_mode'] = _parse_vw_mode(logs)

        # New Features: Input, Power, Status
        sbb_entry['input'] = _parse_simple_lookup(logs, CMD_INPUT_SOURCE, INPUT_SOURCES)
        sbb_entry['power'] = _parse_simple_lookup(logs, CMD_POWER_STATUS, POWER_STATUS)
        sbb_entry['status'] = _parse_simple_lookup(logs, CMD_STATUS_CODE, STATUS_CODES)
        
        # New Features: Network & Identity
        sbb_entry['mac'] = _parse_mac_address(logs)
        sbb_entry['ip_mode'] = _parse_simple_lookup(logs, CMD_IP_MODE, NETWORK_MODES)
        sbb_entry['imei'] = get_mdc_ascii(logs, CMD_IMEI)
        # sbb_name is already getting 0x67 (Device Name) at line 45

        # Placeholders for Resolution (calculated later)
        sbb_entry['res_sbb'] = "Unknown"
        sbb_entry['res_cab'] = "Unknown"
        
        # Video Offset
        layout_info = _extract_layout(logs)
        if layout_info:
            x, y, w, h = layout_info
            sbb_entry['video_offset'] = f"{x}x{y}"
        else:
            sbb_entry['video_offset'] = "Unknown"
        
        sbb_data.append(sbb_entry)

    # 2. Process Cabinets (via Groups to maintain logical hierarchy)
    # But cab_list is flat. We iterate groups to find cabs.
    
    for grp in group_lookup:
        s_ip = grp['sbb_ip']
        g_ip = grp['group_ip']
        
        # Find matching cabs
        curr_group_cabs = []
        for (cip, cid), logs in cab_list.items():
            if cip == g_ip:
                curr_group_cabs.append(((cip, cid), logs))
        
        for (cip, cid), logs in curr_group_cabs:
            c_entry = {}
            c_entry['sbb_ip'] = s_ip
            c_entry['group_ip'] = g_ip
            c_entry['cid'] = cid
            c_entry['sn'] = get_mdc_ascii(logs, CMD_SERIAL_NUM)
            c_entry['model'] = get_mdc_ascii(logs, CMD_MODEL_NAME)
            
            m_fw, f_fw = _parse_cab_fw(logs)
            c_entry['fw_main'] = m_fw
            c_entry['fw_fpga'] = f_fw
            
            # Temp
            c_entry['temp'] = None
            t_bytes = get_mdc_data(logs, CMD_TEMP)
            if t_bytes and len(t_bytes) > 3:
                c_entry['temp'] = int(t_bytes[3], 16)
                
            # Backlight
            c_entry['backlight'] = None
            bl_bytes = get_mdc_data(logs, CMD_BACKLIGHT)
            if bl_bytes and len(bl_bytes) > 1:
                c_entry['backlight'] = int(bl_bytes[1], 16)
                
            # Binaries
            c_entry['cc_cab'] = _parse_binary(logs, CMD_CC_CAB)
            c_entry['cc_mod'] = _parse_binary(logs, CMD_CC_MOD)
            c_entry['cc_pix'] = _parse_binary(logs, CMD_CC_PIX)
            c_entry['seam'] = _parse_binary(logs, CMD_SEAM_COR)
            
            cab_data.append(c_entry)
            
            # Layout
            c_entry['video_location'] = ""
            layout_info = _extract_layout(logs)
            if layout_info:
                x, y, w, h = layout_info
                x_sbb, y_sbb = _apply_group_offset(x, y, grp['group_num'])
                c_entry['video_location'] = f"{x_sbb}x{y_sbb}"
                
                if s_ip not in layout_data: layout_data[s_ip] = []
                layout_data[s_ip].append({
                    'group': grp['group_num'],
                    'cid': cid,
                    'x_grp': x, 'y_grp': y,
                    'w': w, 'h': h,
                    'x_sbb': x_sbb, 'y_sbb': y_sbb
                })

    # 3. Post-Process SBB Resolution from Layout Data
    for sbb in sbb_data:
        sip = sbb['ip']
        if sip in layout_data:
            cabs = layout_data[sip]
            if cabs:
                # Assuming all cabs have same size for now, or just taking first
                w = cabs[0]['w']
                h = cabs[0]['h']
                sbb['res_cab'] = f"{w}x{h}"
                
                # Calculate total SBB resolution
                # Max X + Width, Max Y + Height
                # Note: This is a simplification. Legacy script calculates cols x rows * cab_res.
                # Let's try to match legacy logic: Cols = Unique X Counts, Rows = Unique Y Counts
                
                unique_x = set(c['x_sbb'] for c in cabs)
                unique_y = set(c['y_sbb'] for c in cabs)
                
                # SBox 2 has x_sbb like 0, 1920.
                # But cabinets inside might be 1x6 per group?
                # User says 12x3 total.
                # If SBox has 2 Groups (G1, G3).
                # If G1 is 1x6 (6 cabs). G3 is 1x6 (6 cabs)? No 36 total. 18 per group.
                # If 18 per group. Layout 6x3?
                # 6x3 = 18.
                # If 2 groups side by side (0 and 1920 offset).
                # Total 12x3.
                # So we need to count unique X coordinates across ALL cabs in the SBox.
                # unique_x set length should be 12.
                # unique_y set length should be 3.
                
                cols = len(unique_x)
                rows = len(unique_y)
                
                tot_w = cols * w
                tot_h = rows * h
                sbb['res_sbb'] = f"{tot_w}x{tot_h}"
                sbb['layout_str'] = f"{cols}x{rows}" # Adding for reference if needed
            else:
                sbb['layout_str'] = "0x0"


    # 4. Calculate Cabinet Statistics
    cab_stats = calculate_cab_stats(cab_data)

    return sbb_data, cab_data, layout_data, cab_stats

def calculate_cab_stats(cab_data):
    """Calculates summary statistics for cabinets."""
    stats = {
        'count': len(cab_data),
        'mode_model': "N/A", 'mode_fw_main': "N/A", 'mode_fw_fpga': "N/A",
        'temp_min': None, 'temp_max': None, 'temp_avg': None
    }
    
    if not cab_data:
        return stats
        
    # Helper for mode
    def get_mode(key):
        vals = [c[key] for c in cab_data if c.get(key)]
        if not vals: return "N/A"
        # Manual mode calculation to avoid importing statistics if not needed, 
        # but max(set) is easy.
        return max(set(vals), key=vals.count)

    stats['mode_model'] = get_mode('model')
    stats['mode_fw_main'] = get_mode('fw_main')
    stats['mode_fw_fpga'] = get_mode('fw_fpga')
    
    # Temp Stats
    temps = [c['temp'] for c in cab_data if c.get('temp') is not None]
    if temps:
        stats['temp_min'] = min(temps)
        stats['temp_max'] = max(temps)
        stats['temp_avg'] = round(sum(temps) / len(temps), 1)
        
    return stats

# --- Parsing Helpers (Copied/Adapted) ---

def _parse_add_versions(logs, model_name):
    ver_pattern = None
    count_loc = 0
    data_start = 0
    
    if "AU" in str(model_name): 
         ver_pattern = "41 D2 32"; count_loc = 8; data_start = 11
    elif "3U" in str(model_name):
         ver_pattern = "41 1B A4"; count_loc = 7; data_start = 10
    
    if not ver_pattern: return ""
    
    for entry in reversed(logs):
        if ver_pattern in entry['resp']:
            parts = entry['parts']
            try:
                field_count = int(parts[count_loc], 16)
                current_idx = data_start
                vers = []
                for _ in range(field_count):
                    if current_idx >= len(parts): break
                    f_len = int(parts[current_idx-1], 16)
                    f_end = current_idx + f_len - 1
                    f_hex = parts[current_idx : f_end+1]
                    vers.append(hex_list_to_ascii(f_hex).replace(" ", ""))
                    current_idx = f_end + 3
                return "\n".join(vers)
            except:
                pass
    return ""

def _parse_group_ips(logs: list) -> dict:
    res = {}
    raw_gips = get_mdc_data(logs, CMD_GROUP_IP)
    if raw_gips:
        dec_gips = [int(h, 16) for h in raw_gips]
        curr_ptr = 1 
        for g_num in range(1, 5):
            if curr_ptr + 3 < len(dec_gips):
                ip_bytes = dec_gips[curr_ptr : curr_ptr+4]
                res[g_num] = ".".join(str(b) for b in ip_bytes)
                curr_ptr += 4
    return res

def _parse_vw_mode(logs: list) -> str:
    vw_hex = get_mdc_data(logs, CMD_VW_MODE)
    if vw_hex and len(vw_hex) > 0:
        key_h = vw_hex[0]
        if key_h in SBB_VW_MODE:
            return SBB_VW_MODE[key_h]
    return "Unknown"

def _parse_mac_address(logs: list) -> str:
    # Command 0x1B Sub 0x81
    # Response: AA FF LEN 'A' 1B 81 [MAC 12 bytes or 6 bytes?]
    # PDF says: 0x81 MAC Addr (Hex)
    mac_bytes = get_mdc_data(logs, CMD_MAC_ADDR)
    if mac_bytes and len(mac_bytes) >= 6:
        # Assuming last 6 bytes are the MAC if longer, or just take first 6?
        # PDF Example: 1st Byte ... Val9. It might be variable or specific.
        # Often MAC is 6 bytes.
        
        # Let's take the first 6 bytes if available
        relevant = mac_bytes[:6]
        return ":".join(relevant).upper()
    return "Unknown"

def _parse_cab_fw(logs):
    for entry in reversed(logs):
        if "41 1B A4" in entry['resp']:
            parts = entry['parts']
            try:
                curr = 10
                l = int(parts[curr-1], 16)
                m = hex_list_to_ascii(parts[curr : curr+l]).replace(" ", "")
                curr = curr + l + 3 
                l = int(parts[curr-1], 16)
                f = hex_list_to_ascii(parts[curr : curr+l]).replace(" ", "")
                if f: f = f[:-1] # User requested trim
                return m, f
            except:
                pass
    return "", ""

def _parse_binary(logs: list, pattern: str) -> str:
    b = get_mdc_data(logs, pattern)
    if b and len(b) > 1:
        return ON_OFF.get(b[1], "Unknown")
    return "Unknown"

def _extract_layout(logs: list) -> Optional[tuple]:
    layout_bytes = get_mdc_data(logs, CMD_LAYOUT)
    if layout_bytes and len(layout_bytes) >= 11:
        try:
            x = int("".join(layout_bytes[3:5]), 16)
            y = int("".join(layout_bytes[5:7]), 16)
            w = int("".join(layout_bytes[7:9]), 16)
            h = int("".join(layout_bytes[9:11]), 16)
            return x, y, w, h
        except:
            pass
    return None

def _apply_group_offset(x, y, gid):
    if gid == 2: y += 1080
    elif gid == 3: x += 1920
    elif gid == 4: x += 1920; y += 1080
    return x, y

def _parse_simple_lookup(logs, pattern, lookup_table):
    """Generic helper for single-byte return values mapped to a dict."""
    val_hex = get_mdc_data(logs, pattern)
    if val_hex and len(val_hex) > 0:
        if len(val_hex) >= 1:
            key = val_hex[0] # Hex string like "14"
            return lookup_table.get(key, f"Unknown ({key})")
    return "Unknown"
