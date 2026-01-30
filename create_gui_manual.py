from fpdf import FPDF
import os
from config import VERSION

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, f'LSM Parser GUI v{VERSION} - User Manual', 0, 1, 'C')
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
        "LSM Parser GUI is a modern, user-friendly application for analyzing Samsung Large Screen Manager (LSM) logs. "
        "It provides all the powerful parsing capabilities of the command-line tool wrapped in a convenient graphical interface.\n\n"
        "Key Features:\n"
        "- Interactive View: View generated reports directly within the application.\n"
        "- PDF Export: Easily generate and save PDF versions of reports.\n"
        "- Email Ready: 'Copy PDF' feature makes sharing reports via email effortless.\n"
        "- Native Experience: runs as a standalone Windows application."
    )

    # Installation
    pdf.chapter_title('2. Installation')
    pdf.chapter_body(
        "No complicated installation is required.\n"
        "1. Locate 'LSMParserGUI.exe'.\n"
        "2. Copy it to your preferred location (e.g., Desktop or Documents).\n"
        "3. (Optional) Create a shortcut on your Desktop for easy access."
    )

    # Usage
    pdf.chapter_title('3. Using the Application')
    pdf.chapter_body(
        "Step 1: Launch\n"
        "Double-click 'LSMParserGUI.exe' to open the application.\n\n"
        "Step 2: Select Log File\n"
        "Click the 'Select Log File' button in the top-left corner. Navigate to and select your .xlsx log file.\n\n"
        "Step 3: Analyze\n"
        "The application will automatically process the file. A progress bar will show activity. Once complete, the report will appear in the main window.\n\n"
        "Step 4: Share Features\n"
        "- Generate PDF: Click 'Generate PDF' to save a permanent copy of the report.\n"
        "- Copy PDF: Click 'Copy PDF' to place the PDF file on your clipboard. You can then paste it directly into an Outlook email or Slack message.\n"
        "- Copy Image: Click 'Copy Report Image' to copy the visible report area as an image."
    )

    # Troubleshooting
    pdf.chapter_title('4. Troubleshooting')
    pdf.chapter_body(
        "- 'Not Responding': Large log files may take a moment to process. Please be patient; the interface will unlock once analysis is complete.\n"
        "- 'Failed to Generate PDF': Ensure you have write permissions to the folder where you are trying to save the PDF.\n"
        "- Display Issues: If the report looks incorrect, try resizing the window or generating a full PDF for a paginated view."
    )

    output_path = "LSMParserGUI_Manual.pdf"
    pdf.output(output_path, 'F')
    print(f"GUI Manual generated: {os.path.abspath(output_path)}")

if __name__ == '__main__':
    create_manual()
