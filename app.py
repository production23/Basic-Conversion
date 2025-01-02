import time
import os
import comtypes.client
from flask import Flask, request, send_file

app = Flask(__name__)

# Ensure directories exist
os.makedirs('uploads', exist_ok=True)
os.makedirs('downloads', exist_ok=True)

@app.route('/')
def index():
    return '''
    <h1>File Conversion Service</h1>
    <form action="/upload-pdf" method="post" enctype="multipart/form-data">
      <input type="file" name="file">
      <input type="submit" value="Upload and Convert to Word">
    </form>
    <form action="/upload-word" method="post" enctype="multipart/form-data">
      <input type="file" name="file">
      <input type="submit" value="Upload and Convert to PDF">
    </form>
    '''

@app.route('/upload-word', methods=['POST'])
def upload_word():
    file = request.files['file']
    if file:
        original_filename = file.filename.replace(" ", "_")  # Sanitize filename
        upload_path = os.path.join('uploads', original_filename)
        file.save(upload_path)
        print(f"Saved file to {upload_path}")

        converted_filename = f'{os.path.splitext(original_filename)[0]}_converted_to_pdf.pdf'
        download_path = os.path.join('downloads', converted_filename)

        convert_word_to_pdf(os.path.abspath(upload_path), os.path.abspath(download_path))
        
        if os.path.exists(download_path):
            return send_file(download_path, as_attachment=True)
        else:
            return "File conversion failed or the file does not exist", 500

def convert_word_to_pdf(input_docx, output_pdf):
    print(f"Attempting to convert: {input_docx} to {output_pdf}")
    comtypes.CoInitialize()

    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(input_docx)
        print(f"Opened document: {input_docx}")

        doc.SaveAs(output_pdf, FileFormat=17)
        doc.Close()
        print(f"Saved PDF: {output_pdf}")

    except Exception as e:
        print(f"Error processing file {input_docx}: {e}")

    finally:
        word.Quit()
        comtypes.CoUninitialize()

if __name__ == '__main__':
    app.run(debug=True)
