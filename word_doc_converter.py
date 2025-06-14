import os
import zipfile
import shutil

# Convert minimal HTML to Word XML (fixing the tag structure)
def html_to_word_xml(html):
    # Replace HTML tags with Word's XML equivalent
    html = html.replace('<b>', '<w:r><w:rPr><w:b/></w:rPr><w:t>').replace('</b>', '</w:t></w:r>')
    html = html.replace('<i>', '<w:r><w:rPr><w:i/></w:rPr><w:t>').replace('</i>', '</w:t></w:r>')
    html = html.replace('<u>', '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>').replace('</u>', '</w:t></w:r>')

    # Split the content by paragraphs
    paragraphs = html.split('</p>')

    body = ''
    for para in paragraphs:
        para = para.replace('<p>', '').strip()
        if para:
            body += f'''
            <w:p>
                <w:r>
                    <w:t>{para}</w:t>
                </w:r>
            </w:p>'''

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body}
  </w:body>
</w:document>'''

# Create a valid .docx file from the HTML content
def create_docx_from_html(html, output_path='output.docx'):
    if os.path.exists('temp_docx'):
        shutil.rmtree('temp_docx')

    os.makedirs('temp_docx/word/_rels', exist_ok=True)
    os.makedirs('temp_docx/_rels', exist_ok=True)

    # Write document.xml with the content
    with open('temp_docx/word/document.xml', 'w', encoding='utf-8') as f:
        f.write(html_to_word_xml(html))

    # Write [Content_Types].xml for the Word document package
    with open('temp_docx/[Content_Types].xml', 'w', encoding='utf-8') as f:
        f.write('''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')

    # Write _rels/.rels (relationships for the document)
    with open('temp_docx/_rels/.rels', 'w', encoding='utf-8') as f:
        f.write('''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')

    # Write word/_rels/document.xml.rels (optional, empty)
    with open('temp_docx/word/_rels/document.xml.rels', 'w', encoding='utf-8') as f:
        f.write('''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')

    # Create the final .docx file by zipping all the components
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
        for root, dirs, files in os.walk('temp_docx'):
            for file in files:
                full_path = os.path.join(root, file)
                archive_name = os.path.relpath(full_path, 'temp_docx')
                docx.write(full_path, archive_name)

    # Clean up the temporary directory
    shutil.rmtree('temp_docx')
    print("added")
    print(f"✅ Created: {output_path}")

# === Run this script using the actual HTML file ===
if __name__ == '__main__':
    input_html_path = '/Users/bhargav/Desktop/test/1.htm'
    output_docx_path = '/Users/bhargav/Desktop/test/output.docx'

    if not os.path.exists(input_html_path):
        print(f"❌ HTML file not found: {input_html_path}")
    else:
        with open(input_html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        create_docx_from_html(html_content, output_docx_path)
