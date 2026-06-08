import sys
import markdown
from docx import Document
from docx.shared import Pt
from bs4 import BeautifulSoup

def md_to_docx(md_path, docx_path):
    with open(md_path, "r", encoding="utf-8") as f:
        md_content = f.read()
    
    # Convert markdown to HTML for easier parsing
    html_content = markdown.markdown(md_content)
    soup = BeautifulSoup(html_content, "html.parser")
    
    doc = Document()
    
    for element in soup.children:
        if element.name == 'h1':
            doc.add_heading(element.get_text(), level=1)
        elif element.name == 'h2':
            doc.add_heading(element.get_text(), level=2)
        elif element.name == 'h3':
            doc.add_heading(element.get_text(), level=3)
        elif element.name == 'p':
            p = doc.add_paragraph()
            # Handle basic bolding and italics if needed (simplified here)
            p.add_run(element.get_text())
        elif element.name == 'ul':
            for li in element.find_all('li'):
                doc.add_paragraph(li.get_text(), style='List Bullet')
        elif element.name == 'ol':
            for li in element.find_all('li'):
                doc.add_paragraph(li.get_text(), style='List Number')
    
    doc.save(docx_path)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python md_to_docx_pro.py <input_md> <output_docx>")
        sys.exit(1)
    
    md_to_docx(sys.argv[1], sys.argv[2])
