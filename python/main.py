import docx
import re
import sys

def main(arg):
    file_path = arg[1]
    print(file_path)
    text = extract_text_from_word(file_path)
    text = '\n'.join(text)
    search = extract_number(text)
    print(search)

def extract_text_from_word(file_path):
    document = docx.Document(file_path)
    text = []
    for para in document.paragraphs:
        text.append(para.text)
    return text

def extract_number(text):
    match = re.search("Escritura Pública Número (.*?)-", text)
    if match:
        return match.group(1)
    else:
        return None

if __name__ == "__main__":
    main(sys.argv)