# --- Configuration & Constants ---
VERSION = "1.0"

# Color Codes
COLOR_GREEN = "90EE90"
COLOR_GRAY = "D3D3D3"
COLOR_YELLOW = "FFFF00"
COLOR_TAN = "FFE4B5"
COLOR_VIOLET = "E6E6FA"
COLOR_BLUE = "ADD8E6"
COLOR_PINK = "FFB6C1" # Light Pink

# MDC Data Tables
SBB_TYPE = {
    "MSM": "SBB-MBOX",
    "HMS": "SBB-SNOW",
    "KTM": "SBB-SNOW",
    "GSS": "SBB-ISO8E"
}

DEV_TYPE = {
    "00": "Reserved",
    "01": "SBox",
    "02": "Cabinet-IS/IFH/IFH-D",
    "03": "Cabinet-IFJ/2in1"
}

SBB_VW_MODE = {
    "00": "Off",
    "01": "On"
}

ON_OFF = {
    "00": "Off",
    "01": "On"
}

# Input Source (0x14)
INPUT_SOURCES = {
    "14": "PC",
    "18": "DVI",
    "0C": "AV",
    "04": "S-Video",
    "08": "Component",
    "20": "MagicInfo",
    "1F": "DVI_VIDEO",
    "21": "HDMI1",
    "22": "HDMI1_PC",
    "23": "HDMI2",
    "24": "HDMI2_PC",
    "25": "DisplayPort",
    "60": "HDBaseT"
}

# Power Status (0x11)
POWER_STATUS = {
    "00": "Off",
    "01": "On",
    "02": "Reboot"
}

# Network IP Mode (0x1B 0x85)
NETWORK_MODES = {
    "00": "Dynamic",
    "01": "Static"
}

# Status Error Codes (0x0D)
STATUS_CODES = {
    "00": "Normal",
    "01": "Fan Error",
    "02": "Fan Error",
    "03": "Lamp Error",
    "04": "Brightness Sensor Error",
    "06": "Source Error",
    "07": "Temp Error",
    "08": "Sent/Panel Error"
}

# --- MDC Command Constants ---
CMD_MODEL_NAME    = "41 8A"
CMD_SERIAL_NUM    = "41 0B"
CMD_FW_MAIN       = "41 0E"
CMD_DEVICE_NAME   = "41 67"
CMD_VW_MODE       = "41 84"
CMD_MAC_ADDR      = "41 1B 81"
CMD_GROUP_IP      = "41 1B 84"
CMD_IP_MODE       = "41 1B 85"
CMD_IMEI          = "41 1B 83"

CMD_INPUT_SOURCE  = "41 14"
CMD_POWER_STATUS  = "41 11"
CMD_STATUS_CODE   = "41 0D"

CMD_LAYOUT        = "41 8C A0"
CMD_TEMP          = "41 D0 84"
CMD_BACKLIGHT     = "41 D0 94"
CMD_CC_CAB        = "41 D0 9E"
CMD_CC_MOD        = "41 D0 99"
CMD_CC_PIX        = "41 D0 95"
CMD_SEAM_COR      = "41 D0 98"

