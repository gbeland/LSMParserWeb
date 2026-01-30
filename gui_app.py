
import sys
import os
import shutil
import logging
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                            QMessageBox, QProgressBar, QTextEdit)
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtCore import QUrl, QThread, pyqtSignal, Qt, QStandardPaths, QMimeData
from PyQt6.QtGui import QAction, QIcon, QDesktopServices, QPixmap

def resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Add repo path to sys.path to access modules
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

from main import analyze_file_logic
from config import VERSION

# Configure logging for GUI
logging.basicConfig(filename="LSMParserGUI.log", level=logging.DEBUG, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger("LSMParserGUI")

class AnalyzerThread(QThread):
    finished = pyqtSignal(tuple) # (excel_path, html_path) or None
    error = pyqtSignal(str)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        try:
            result = analyze_file_logic(self.file_path)
            if result:
                self.finished.emit(result)
            else:
                self.error.emit("Analysis failed to produce output files.")
        except Exception as e:
            logger.error(f"Analysis Thread Error: {e}", exc_info=True)
            self.error.emit(str(e))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle(f"LSM Parser v{VERSION}")
        self.resize(1200, 800)
        
        # Central Widget & Layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Top Bar (Controls)
        top_layout = QHBoxLayout()
        
        self.btn_select = QPushButton("Select Log File")
        self.btn_select.clicked.connect(self.select_file)
        self.btn_select.setMinimumHeight(40)
        top_layout.addWidget(self.btn_select)

        self.btn_pdf = QPushButton("Generate PDF")
        self.btn_pdf.clicked.connect(self.generate_pdf)
        self.btn_pdf.setMinimumHeight(40)
        self.btn_pdf.setEnabled(False)
        top_layout.addWidget(self.btn_pdf)

        self.btn_copy_pdf = QPushButton("Copy PDF")
        self.btn_copy_pdf.clicked.connect(self.copy_pdf_to_clipboard)
        self.btn_copy_pdf.setMinimumHeight(40)
        self.btn_copy_pdf.setEnabled(False)
        top_layout.addWidget(self.btn_copy_pdf)

        self.btn_copy_img = QPushButton("Copy Report Image")
        self.btn_copy_img.clicked.connect(self.copy_image_to_clipboard)
        self.btn_copy_img.setMinimumHeight(40)
        self.btn_copy_img.setEnabled(False)
        top_layout.addWidget(self.btn_copy_img)
        
        main_layout.addLayout(top_layout)

        # Progress Bar
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        main_layout.addWidget(self.progress)

        # Web View (Report Display)
        self.web_view = QWebEngineView()
        main_layout.addWidget(self.web_view)

        # Status Bar
        self.status_label = QLabel("Ready")
        self.statusBar().addWidget(self.status_label)

        self.current_html_path = None
        self.current_excel_path = None
        self.pdf_output_path = None

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select LSM Log Excel File", "", "Excel Files (*.xlsx)")
        if file_path:
            self.start_analysis(file_path)

    def start_analysis(self, file_path):
        self.status_label.setText(f"Analyzing: {os.path.basename(file_path)}...")
        self.progress.setRange(0, 0) # Indeterminate
        self.progress.setVisible(True)
        self.btn_select.setEnabled(False)
        
        self.thread = AnalyzerThread(file_path)
        self.thread.finished.connect(self.on_analysis_finished)
        self.thread.error.connect(self.on_analysis_error)
        self.thread.start()

    def on_analysis_finished(self, paths):
        self.progress.setVisible(False)
        self.btn_select.setEnabled(True)
        
        self.current_excel_path, self.current_html_path = paths
        self.status_label.setText(f"Analysis Complete. Loaded: {os.path.basename(self.current_html_path)}")
        
        # Load HTML
        local_url = QUrl.fromLocalFile(os.path.abspath(self.current_html_path))
        self.web_view.setUrl(local_url)
        
        # Enable buttons
        self.btn_pdf.setEnabled(True)
        self.btn_copy_img.setEnabled(True)
        self.btn_copy_pdf.setEnabled(True)

    def on_analysis_error(self, message):
        self.progress.setVisible(False)
        self.btn_select.setEnabled(True)
        self.status_label.setText("Analysis Failed.")
        QMessageBox.critical(self, "Error", f"An error occurred:\n{message}")

    def generate_pdf(self):
        if not self.current_html_path:
            return

        default_name = os.path.splitext(os.path.basename(self.current_html_path))[0] + ".pdf"
        output_path, _ = QFileDialog.getSaveFileName(self, "Save PDF Report", default_name, "PDF Files (*.pdf)")
        
        if output_path:
            self.pdf_output_path = output_path
            # Use QWebEngineView printToPdf
            # Note: This is async. We need to handle the callback.
            self.web_view.page().printToPdf(output_path)
            self.web_view.page().pdfPrintingFinished.connect(self.on_pdf_finished)
            self.status_label.setText("Generating PDF...")

    def copy_pdf_to_clipboard(self):
        """Copies the PDF file to the clipboard (as a file/url) to paste into email."""
        # 1. Check if we have a PDF generated already
        if self.pdf_output_path and os.path.exists(self.pdf_output_path):
            self._do_copy_pdf(self.pdf_output_path)
        else:
            # 2. Must generate it first
            # We can't use self.generate_pdf() directly because it opens a dialog.
            # We want a smoother flow? Or just trigger that?
            # User request: "allows the user to copy a PDF version to send via email"
            # It implies we need the file.
            # Let's ask user to save it first if they haven't.
            
            # Simple approach: Call generate_pdf, and set a flag to copy after? 
            # But generate_pdf is async.
            self.pending_copy_action = True
            self.generate_pdf()
    
    def _do_copy_pdf(self, file_path):
        data = QMimeData()
        url = QUrl.fromLocalFile(file_path)
        data.setUrls([url])
        QApplication.clipboard().setMimeData(data)
        self.status_label.setText(f"PDF Copied to Clipboard: {os.path.basename(file_path)}")
        QMessageBox.information(self, "Copied", "PDF file has been copied to clipboard.\nYou can now paste it into an email.")


    def on_pdf_finished(self, file_path, success):
         # Disconnect signal to avoid multiple calls if multiple PDFs generated in session
        try:
            self.web_view.page().pdfPrintingFinished.disconnect(self.on_pdf_finished)
        except:
            pass

        if success:
            self.status_label.setText(f"PDF Saved: {file_path}")
            
            # Check for pending copy
            if getattr(self, 'pending_copy_action', False):
                self.pending_copy_action = False
                self._do_copy_pdf(file_path)
                return
            
            # Ask to open or copy path
            msg = QMessageBox()
            msg.setWindowTitle("PDF Generated")
            msg.setText(f"PDF saved successfully to:\n{file_path}")
            btn_open = msg.addButton("Open PDF", QMessageBox.ButtonRole.ActionRole)
            btn_copy = msg.addButton("Copy Path", QMessageBox.ButtonRole.ActionRole)
            msg.addButton(QMessageBox.StandardButton.Close)
            
            msg.exec()
            
            if msg.clickedButton() == btn_open:
                QDesktopServices.openUrl(QUrl.fromLocalFile(file_path))
            elif msg.clickedButton() == btn_copy:
                QApplication.clipboard().setText(file_path)
                self.status_label.setText("PDF Path copied to clipboard.")
                
        else:
            QMessageBox.critical(self, "Error", "Failed to generate PDF.")
            self.status_label.setText("PDF Generation Failed.")

    def copy_image_to_clipboard(self):
        # Capture the entire web view page logic is complex for scrolling content.
        # Simple approach: Check if we can just grab the viewport.
        # But QWebEngineView doesn't support grab() like normal widgets easily due to separate process.
        # We can select all and copy text, but user wants "paste into email" which implies image/rich text.
        
        # Alternative: Since we have the PDF, maybe we can't easily grab image of full scroll.
        # But user asked for "copy to clipboard to allow user to paste it into an email".
        # Pasting HTML into Outlook is tricky. Pasting an image is easier.
        
        # Let's try to notify user that best way is to snapshot visible area or use snipping tool?
        # Or, we can use grab() on the widget container. It might only grab visible part.
        
        pixmap = self.web_view.grab()
        QApplication.clipboard().setPixmap(pixmap)
        self.status_label.setText("Visible report area copied to clipboard.")
        QMessageBox.information(self, "Copied", "Visible report area copied to clipboard.\n\nFor a full document copy, please generate a PDF and attach it, or use the PDF viewer's snapshot tool.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName(f"LSM Parser v{VERSION}")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())
