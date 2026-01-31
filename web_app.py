import os
import sys
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, send_from_directory
from werkzeug.utils import secure_filename
from main import analyze_file_logic
from utils.email_sender import send_email_report
from utils.pdf_generator import generate_pdf_report

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Change this for production
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB limit

import time

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def cleanup_old_files(folder, retention_seconds=3600):
    """
    Deletes files in the specified folder that are older than retention_seconds.
    Also attempts to clean up generated report files in the current directory.
    """
    now = time.time()
    
    # 1. Clean 'uploads' folder
    if os.path.exists(folder):
        for filename in os.listdir(folder):
            filepath = os.path.join(folder, filename)
            try:
                if os.path.isfile(filepath):
                    file_age = now - os.path.getmtime(filepath)
                    if file_age > retention_seconds:
                        os.remove(filepath)
                        print(f"Cleaned up old upload: {filename}")
            except Exception as e:
                print(f"Error cleaning {filename}: {e}")

    # 2. Clean generated reports in CWD (Be careful only to delete report-like files)
    # Generated files usually end with .html, .xlsx (Parsed), .pdf, .json
    # We will look for files matching the output patterns AND older than retention
    cwd = os.getcwd()
    for filename in os.listdir(cwd):
        # Safety check: Only delete specific extensions and ensuring they aren't source code
        if filename.endswith((".html", ".pdf", ".json")) or (filename.endswith(".xlsx") and "Parsed" in filename):
            filepath = os.path.join(cwd, filename)
            try:
                if os.path.isfile(filepath):
                    file_age = now - os.path.getmtime(filepath)
                    if file_age > retention_seconds:
                        os.remove(filepath)
                        print(f"Cleaned up old report: {filename}")
            except Exception as e:
                print(f"Error cleaning {filename}: {e}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # Trigger cleanup on upload to keep storage usage in check without background tasks
    cleanup_old_files(app.config['UPLOAD_FOLDER'])

    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)
    
    if file and file.filename.endswith('.xlsx'):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            # Parse the file
            # analyze_file_logic returns (xlsx_path, html_path)
            # Paths returned are relative to CWD or absolute depending on how analyze_file_logic works.
            # Currently it saves in CWD. We might want to move them to uploads or a 'reports' folder.
            # For now, let's assume valid return.
            xlsx_path, html_path = analyze_file_logic(filepath)
            
            if not html_path:
                 flash('Parsing failed.')
                 return redirect(url_for('index'))

            # Pass the result filename (basename) to the report view
            report_filename = os.path.basename(html_path)
            return redirect(url_for('view_report', filename=report_filename))
            
        except Exception as e:
            flash(f'Error processing file: {str(e)}')
            return redirect(url_for('index'))
            
    else:
        flash('Invalid file type. Please upload an .xlsx file.')
        return redirect(url_for('index'))

@app.route('/report/<filename>')
def view_report(filename):
    # The analyze_file_logic saves files in the current working directory.
    # We serve them from there.
    return render_template('report.html', filename=filename)

@app.route('/reports/<path:filename>')
def serve_report_file(filename):
    # Serve the generated HTML file directly for the iframe
    return send_from_directory(os.getcwd(), filename)

@app.route('/download/<filename>/<fmt>')
def download_file(filename, fmt):
    # filename is the HTML report filename "Log-Parsed.html"
    # base name is "Log-Parsed"
    base_name = os.path.splitext(filename)[0]
    
    if fmt == 'pdf':
        try:
            pdf_path = generate_pdf_report(filename)
            return send_file(pdf_path, as_attachment=True)
        except Exception as e:
            flash(f"Error generating PDF: {e}")
            return redirect(url_for('view_report', filename=filename))
            
    elif fmt == 'xlsx':
        # Result xlsx is usually "Log-Parsed.xlsx"
        xlsx_filename = base_name + ".xlsx"
        if os.path.exists(xlsx_filename):
             return send_file(xlsx_filename, as_attachment=True)
        else:
             flash("Excel file not found.")
             return redirect(url_for('view_report', filename=filename))
             
    return redirect(url_for('view_report', filename=filename))

@app.route('/email', methods=['POST'])
def email_report():
    filename = request.form.get('filename')
    recipient = request.form.get('recipient')
    cc = request.form.get('cc')
    subject_from = request.form.get('from_field', 'LSM Parser Report')
    
    if not filename or not recipient:
        flash("Missing filename or recipient")
        return redirect(url_for('view_report', filename=filename))
        
    try:
        # We need the PDF to attach
        pdf_path = generate_pdf_report(filename)
        
        # Send
        send_email_report(recipient, cc, subject_from, pdf_path)
        flash(f"Email sent successfully to {recipient}")
    except Exception as e:
        flash(f"Failed to send email: {e}")
        
    return redirect(url_for('view_report', filename=filename))

if __name__ == '__main__':
    # SSL Context Logic
    ssl_context = None
    cert_file = 'cert.pem'
    key_file = 'key.pem'
    
    if os.path.exists(cert_file) and os.path.exists(key_file):
        print(f"Loading SSL context from {cert_file} and {key_file}")
        ssl_context = (cert_file, key_file)
    else:
        print("No SSL certificates found. Using Ad-hoc SSL (Self-signed, ephemeral) for HTTPS.")
        try:
             import OpenSSL
             ssl_context = 'adhoc'
        except ImportError:
             print("WARNING: pyopenssl not installed. HTTPS might fail with 'adhoc'. Install it with 'pip install pyopenssl'.")
             ssl_context = None

    if ssl_context:
        print("Running with HTTPS enabled.")
        app.run(debug=True, port=5000, ssl_context=ssl_context)
    else:
        print("Running in HTTP mode (No SSL).")
        app.run(debug=True, port=5000)
