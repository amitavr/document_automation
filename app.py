from flask import Flask, request, send_file, redirect
import os
import src.document_automation.process as process

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['PROCESSED_FOLDER'] = 'processed'

# This is generic webpage build where the user can upload the file.
@app.route('/document_automation')
def document_automation():
    return '''
    <!doctype html>
    <html>
    <head>
        <title>TITLE</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 20px;
                background-color: #e8f5e9;
            }
            h1 {
                color: #388e3c;
                text-align: center;
            }
            h2 {
                color: #388e3c;
                text-align: center;
            }
            form {
                max-width: 500px;
                margin: 0 auto;
                background-color: #ffffff;
                padding: 20px;
                border-radius: 5px;
                box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
            }
            label {
                display: block;
                margin-bottom: 10px;
                font-weight: bold;
                color: #2e7d32;
            }
            input[type="file"], input[type="text"] {
                margin-bottom: 10px;
                width: 200px; /* Adjusted width */
                padding: 8px;
                box-sizing: border-box;
            }
            input[type="submit"] {
                background-color: #388e3c;
                color: white;
                padding: 10px 20px;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                float: right;
                margin-top: -70px;
            }
            input[type="submit"]:hover {
                background-color: #2e7d32;
            }
        </style>
    </head>
    <body>
        <h1>HEADER 1</h1>
        <h2>Upload File</h2>
        <form action="/document_automation" method="post" enctype="multipart/form-data">
            <label for="file">Select file:</label>
            <input type="file" name="file"><br><br>
            <input type="submit" value="Upload">
        </form>
    </body>
    </html>
    '''

# This method handling when there's a POST request on website, 'Upload' button clicked.
@app.route('/document_automation', methods=['POST'])
def upload_wspw_ramp():
    # Check if file is provided.
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    # Verify if file name is not null
    if file.filename == '':
        return redirect(request.url)
    # Make sure we handling with an expected file extension, that case is .xlsx
    if file and file.filename.endswith('.xlsx'):
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file_path = f"src\\document_automation\\{file_path}"
        file.save(file_path)
        
        try:
            document_automation.processing(file_path)
        except Exception as error:
            return redirect(request.error)

        return send_file('src\\document_automation\\processed\\output.xlsx', as_attachment=True)
