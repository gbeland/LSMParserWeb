# LSM Log Analyzer (Python Port)

> Based on the amazing work by Chuck Tinker.

A Python tool for analyzing Samsung LED Signage Manager (LSM) logs exported from Excel. This tool mimics the functionality of the original PowerShell script but provides faster processing and removes the dependency on Excel COM automation.

## Features

- **SBox Analysis**: Extracts Model, Firmware Version, Serial Number, Group IPs, and Video Wall Mode.
- **Cabinet Analysis**: Detailed report on Cabinet setup including FW versions, Temperature, Layout, and Color Correction status.
- **Layout Visualization**: Generates a visual grid of the cabinet layout based on coordinate data.
- **Excel Report**: Outputs a formatted Excel file with color-coded status indicators.
- **HTML Report**: Generates a rich-text HTML file for viewing results without Excel.

## Requirements

- Python 3.x
- `openpyxl`
- `tkinter` (usually included with Python)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/gbeland/LSMParserPy.git
   ```
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Run the script without arguments to open a file selection dialog:
```bash
### macOS / Linux
```bash
python3 main.py
```
Or with a file:
```bash
python3 main.py "path/to/log.xlsx"
```

### Windows
```bash
python main.py
```

> **Note for Windows Users:**
> To enable Excel COM fallback (required for opening DRM-protected files like NASCA), you must install `pywin32` separately:
> ```bash
> pip install pywin32
> ```

```

The script will generate a new file named `[OriginalName]-PyLLA.xlsx` in the same directory.
