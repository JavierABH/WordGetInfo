import docx
import re

def main():
    file_path = "sample\input 5.doc"
    text = extract_text_from_word(file_path)
    text = '\n'.join(text)
    # search = extract_number(text)
    # print(search)
    print(text)

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
    main()