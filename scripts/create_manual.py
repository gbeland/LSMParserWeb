from fpdf import FPDF
import os
from config import VERSION

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, f'LSM Parser v{VERSION} - User Manual', 0, 1, 'C')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.set_fill_color(200, 220, 255)
        self.cell(0, 6, title, 0, 1, 'L', 1)
        self.ln(4)

    def chapter_body(self, body):
        self.set_font('Arial', '', 11)
        self.multi_cell(0, 5, body)
        self.ln()

def create_manual():
    pdf = PDF()
    pdf.add_page()
    
    # Overview
    pdf.chapter_title('1. Overview')
    pdf.chapter_body(
        "LSM Parser is a standalone tool designed to analyze log files from Samsung Large Screen Manager (LSM). "
        "It parses raw Excel (.xlsx) logs, extracts technical data (IPs, Firmware, Temperatures), "
        "and generates human-readable reports in both Excel and HTML formats.\n\n"
        "Key Features:\n"
        "- Standalone: No installation or Excel license required.\n"
        "- Robust: Automatically handles and repairs corrupted log files.\n"
        "- Batch Processing: Supports wildcards to process multiple files at once.\n"
        "- Speed: significantly faster than legacy PowerShell tools."
    )

    # Installation
    pdf.chapter_title('2. Installation')
    pdf.chapter_body(
        "No installation is required. The application is provided as a single executable file: 'LSMParser.exe'.\n"
        "Simply copy this file to any location on your computer (e.g., Desktop or Documents) and it is ready to run."
    )

    # Usage
    pdf.chapter_title('3. Usage')
    pdf.chapter_body(
        "There are three ways to use the tool:\n\n"
        "A. GUI / File Selector (Easiest)\n"
        "   1. Double-click 'LSMParser.exe'.\n"
        "   2. A file selection window will appear.\n"
        "   3. Choose your target .xlsx log file.\n"
        "   4. The tool will run and generate reports in the same folder as the log file.\n\n"
        "B. Drag and Drop\n"
        "   1. Drag an .xlsx file directly onto 'LSMParser.exe'.\n"
        "   2. A console window will open, show progress, and wait for you to press Enter.\n\n"
        "C. Command Line / Batch\n"
        "   Run from a terminal (PowerShell or CMD) for advanced options:\n"
        "   > LSMParser.exe <filename>      (Process a single file)\n"
        "   > LSMParser.exe *.xlsx          (Process ALL excel files)\n"
        "   > LSMParser.exe \"Logs/*.xlsx\"   (Process files in a subfolder)\n"
        "   > LSMParser.exe --help          (Show help)\n"
        "   > LSMParser.exe --version       (Show version)\n\n"
        "   Note: If you paste a path with spaces without quotes, the tool will attempt\n"
        "   to auto-correct it ('Smart Stitching')."
    )

    # Output
    pdf.chapter_title('4. Output Reports')
    pdf.chapter_body(
        "For every input file (e.g., 'Log.xlsx'), two report files are created in the same directory:\n\n"
        "1. Excel Report ('Log-Parsed.xlsx'):\n"
        "   - SBBInfo: Firmware versions, Groups, Video Wall resolution.\n"
        "   - CabInfo: Detailed cabinet status (Temp, Backlight, Color Correction).\n"
        "   - CabLayouts: Visual grid showing physical arrangement.\n\n"
        "2. HTML Report ('Log-Parsed.html'):\n"
        "   - A quick-view summary that can be opened in any web browser.\n\n"
        "Note: Files ending in '-Parsed.xlsx' are ignored to prevent re-processing."
    )
    
    # File Loading Strategy
    pdf.chapter_title('5. File Loading Strategy (DRM Support)')
    pdf.chapter_body(
        "The parser uses a smart multi-stage strategy to open Excel files, ensuring maximum compatibility even with encrypted files:\n\n"
        "1. NASCA DRM Detection: If a file is encrypted with NASCA DRM, the tool immediately attempts to use Microsoft Excel (via COM automation) to decrypt and read it transparently. *Requires Microsoft Excel to be installed.*\n"
        "2. Standard Parser: For normal files, it uses a high-speed library (OpenPyXL) that does not require Excel.\n"
        "3. Fallback Mode: If the Standard Parser fails (due to file structure issues) or if DRM is present but Excel fails, it falls back to a 'Raw XML' parser.\n\n"
        "Note: Users with NASCA-protected files must have Microsoft Excel installed and be logged in to their DRM client."
    )

    # Troubleshooting
    pdf.chapter_title('6. Troubleshooting')
    pdf.chapter_body(
        "- 'PyWin32 not installed': The executable is missing required libraries for DRM support. Please use the latest official release.\n"
        "- 'Excel Automation failed': You have a DRM file, but Microsoft Excel is not installed or cannot open the file. Ensure you can open the file manually in Excel first.\n"
        "- 'IllegalCharacterError': The log contained hidden control characters. The tool automatically sanitizes these fields.\n"
        "- Window closes instantly: The tool is designed to exit when finished. Run from a command prompt to see persistent output."
    )

    output_path = "LSMParser_Manual.pdf"
    pdf.output(output_path, 'F')
    print(f"Manual generated: {os.path.abspath(output_path)}")

if __name__ == '__main__':
    create_manual()
