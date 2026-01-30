import re

def extract_strings(filename, min_len=4):
    with open(filename, 'rb') as f:
        content = f.read()
    
    # Regex for printable ASCII characters
    pattern = re.compile(b'[ -~]{' + str(min_len).encode() + b',}')
    strings = pattern.findall(content)
    
    with open('MDC_strings.txt', 'w', encoding='utf-8') as out:
        for s in strings:
            try:
                decoded = s.decode('utf-8')
                out.write(decoded + '\n')
            except:
                pass

if __name__ == '__main__':
    print("Extracting strings from MDC.pdf...")
    extract_strings('MDC.pdf', min_len=5)
    print("Done. Saved to MDC_strings.txt")
