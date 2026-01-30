from pypdf import PdfReader
import re

def parse_pdf(filename):
    print(f"Reading {filename}...")
    try:
        reader = PdfReader(filename)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        
        with open("MDC_Extracted.txt", "w", encoding="utf-8") as f:
            f.write(text)
            
        print(f"Successfully extracted {len(text)} characters.")
        
        # Quick validation of typical MDC patterns
        hex_pattern = re.findall(r'0x[0-9A-Fa-f]{2}', text)
        print(f"Found {len(hex_pattern)} hex codes.")
        
    except Exception as e:
        print(f"Error reading PDF: {e}")

if __name__ == "__main__":
    parse_pdf("MDC.pdf")
