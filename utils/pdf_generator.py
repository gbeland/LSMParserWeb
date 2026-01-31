import os

# Try to import PDF libraries
pdf_lib = None

try:
    from xhtml2pdf import pisa
    pdf_lib = 'xhtml2pdf'
except ImportError:
    try:
        from weasyprint import HTML
        pdf_lib = 'weasyprint'
    except ImportError:
        pdf_lib = None

def generate_pdf_report(html_filename):
    """
    Generates a PDF from the HTML report file.
    """
    html_path = os.path.join(os.getcwd(), html_filename)
    base_name = os.path.splitext(html_filename)[0]
    pdf_path = os.path.join(os.getcwd(), f"{base_name}.pdf")
    
    if not os.path.exists(html_path):
        raise FileNotFoundError(f"HTML Report not found: {html_path}")
        
    print(f"Generating PDF for {html_path}...")
    
    if pdf_lib == 'xhtml2pdf':
        with open(html_path, "r", encoding='utf-8') as source_html:
            with open(pdf_path, "wb") as output_pdf:
                pisa_status = pisa.CreatePDF(source_html, dest=output_pdf)
        if pisa_status.err:
            raise Exception(f"PDF Generation Failed: {pisa_status.err}")
            
    elif pdf_lib == 'weasyprint':
        HTML(filename=html_path).write_pdf(pdf_path)
        
    else:
        raise Exception("PDF generation libraries (xhtml2pdf or WeasyPrint) are not installed. Please install one of them to enable PDF support.")
    
    return pdf_path
