from flask import Flask, request, render_template, send_file 
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Inches
import os
import docx
from waitress import serve

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'docx'}

# Ensure the upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def merge_documents(files):
    merged_content = {}
    current_heading = ' '  # Default heading for content without a proper heading
    merged_content[current_heading] = {}

    for file in files:
        doc = Document(file)
        current_subheading = None
        for para in doc.paragraphs:
            if is_heading(para):
                current_heading = para.text.strip()
                current_subheading = None
                if current_heading not in merged_content:
                    merged_content[current_heading] = {}

            elif is_subheading(para):
                current_subheading = para.text.strip()
                if current_subheading not in merged_content[current_heading]:
                    merged_content[current_heading][current_subheading] = {'paragraphs': [], 'tables': []}

            elif current_heading and current_subheading:
                merged_content[current_heading][current_subheading]['paragraphs'].append(para)

            elif current_heading:
                if None not in merged_content[current_heading]:
                    merged_content[current_heading][None] = {'paragraphs': [], 'tables': []}
                merged_content[current_heading][None]['paragraphs'].append(para)

        # Append tables in each document to the appropriate section in merged content
        for table in doc.tables:
            if current_heading:
                if current_subheading:
                    merged_content[current_heading][current_subheading]['tables'].append(table)
                else:
                    merged_content[current_heading][None]['tables'].append(table)

    # Create a new document for the merged output
    merged_doc = Document()

    # Add merged headings, subheadings, and combined tables to the new document
    for heading, subcontent in merged_content.items():
        merged_doc.add_heading(heading, level=1)
        for subheading, elements in subcontent.items():
            if subheading:
                merged_doc.add_heading(subheading, level=2)

            # Add paragraphs under the heading/subheading
            for para in elements['paragraphs']:
                copy_paragraph_and_images(para, merged_doc)

            # Merge all tables under the same heading/subheading with identical headers
            if elements['tables']:
                combined_tables = merge_similar_tables(elements['tables'])
                for table in combined_tables:
                    copy_table(table, merged_doc)

    output_path = 'merged_report.docx'
    merged_doc.save(output_path)
    return output_path

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        try:
            files = request.files.getlist('files')
            if files and all(allowed_file(f.filename) for f in files):
                file_paths = []
                for f in files:
                    filename = secure_filename(f.filename)
                    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    f.save(file_path)
                    file_paths.append(file_path)

                merged_file_path = merge_documents(file_paths)
                return send_file(merged_file_path, as_attachment=True)
            else:
                return "Invalid file type. Only .docx files are allowed.", 400

        except Exception as e:
            return f"An error occurred: {str(e)}", 500  # Return error message on exception
    
    return render_template('upload.html')

def is_heading(para):
    return para.style.name.startswith('Heading')

def is_subheading(para):
    return para.style.name.startswith('Heading') and not para.style.name.startswith('Heading 1')

def copy_paragraph_and_images(source_para, target_doc):
    target_para = target_doc.add_paragraph()
    for run in source_para.runs:
        new_run = target_para.add_run(run.text)
        new_run.bold = run.bold
        new_run.italic = run.italic
        new_run.underline = run.underline
        
        for drawing in run._element.xpath('.//w:drawing'):
            blip = drawing.xpath('.//a:blip')
            if blip:
                image_rId = blip[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                image_part = run.part.related_parts[image_rId]
                image_path = f"temp_image_{image_rId}.png"
                with open(image_path, "wb") as img_file:
                    img_file.write(image_part.blob)
                target_doc.add_picture(image_path, width=Inches(2))
                os.remove(image_path)

def set_cell_borders(cell):
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    tcBorders = docx.oxml.parse_xml(r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                                      r'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                      r'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                      r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                      r'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                      r'</w:tcBorders>')
    tcPr.append(tcBorders)

def copy_table(source_table, target_doc):
    new_table = target_doc.add_table(rows=len(source_table.rows), cols=len(source_table.columns))
    for row_idx, row in enumerate(source_table.rows):
        for col_idx, cell in enumerate(row.cells):
            new_table.cell(row_idx, col_idx).text = cell.text
            set_cell_borders(new_table.cell(row_idx, col_idx))

def get_header_text(table):
    return [cell.text.strip() for cell in table.rows[0].cells]

def merge_similar_tables(tables):
    combined_tables = []
    header_groups = {}

    for table in tables:
        header_text = tuple(get_header_text(table))  # Make header hashable
        if header_text not in header_groups:
            header_groups[header_text] = Document().add_table(rows=0, cols=len(table.columns))
            # Add header row only once with bold formatting
            header_row = header_groups[header_text].add_row().cells
            for idx, cell in enumerate(table.rows[0].cells):
                # Create a new paragraph for the header cell and apply bold formatting
                header_paragraph = header_row[idx].add_paragraph()  # Create a new paragraph for the header
                header_run = header_paragraph.add_run(cell.text)  # Add run to the new paragraph
                header_run.bold = True  # Apply bold formatting to header text
        
        # Add non-header rows from each table with matching headers
        for row in table.rows[1:]:  # Skip header row for merging
            new_row = header_groups[header_text].add_row().cells
            for idx, cell in enumerate(row.cells):
                new_row[idx].text = cell.text

    combined_tables.extend(header_groups.values())
    return combined_tables

if __name__ == '__main__':
    # Use serve(app, host='0.0.0.0', port=50100, threads=2) for production
    app.run(debug=True)  # Set debug=True for development
