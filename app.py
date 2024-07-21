from flask import Flask, request, send_file, render_template, redirect, url_for, session
from flask_socketio import SocketIO, emit
import os
import zipfile
import shutil
import win32com.client as client
import pythoncom

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Replace with a secure key for session management
socketio = SocketIO(app, cors_allowed_origins="*")

# Ensure the directories exist (for storing uploaded & converted files)
os.makedirs('uploads', exist_ok=True)
os.makedirs('converted', exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    pythoncom.CoInitialize()  # Initialize COM library 
    try:
        # save uploaded zip file in 'uploads' directory
        uploaded_file = request.files['file']
        if uploaded_file.filename == '':
            return 'No file selected'
        zip_path = os.path.join('uploads', uploaded_file.filename)
        uploaded_file.save(zip_path)

        # Unzip the uploaded file
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall('uploads')

        # Initialize Excel application
        excel = client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False

        converted_files = []
        total_files = len([f for f in os.listdir('uploads') if f.endswith(('.xls', '.xlsx'))])
        processed_files = 0

        for file in os.listdir('uploads'):
            file_path = os.path.abspath(os.path.join('uploads', file))
            if file.endswith(('.xls', '.xlsx')):
                try:
                    if os.path.exists(file_path):
                        print(f"Processing file: {file_path}")
                        wb = excel.Workbooks.Open(file_path)
                        filename, fileextension = os.path.splitext(file)
                        output_path = os.path.abspath(os.path.join('converted', filename + '.xlsx'))
                        wb.SaveAs(output_path, FileFormat=51)  # 51 represents the XLSX format
                        wb.Close(False)
                        converted_files.append(output_path)

                        processed_files += 1
                        progress = (processed_files / total_files) * 100
                        print(f"Progress: {progress}%")  # Debugging print
                        socketio.emit('progress', {'progress': progress})
                    else:
                        print(f"File not found: {file_path}")
                except Exception as e:
                    print(f"Error processing file {file}: {str(e)}")

        excel.Quit()

        # Zip the converted files
        converted_zip_path = os.path.abspath(os.path.join('converted', 'converted_files.zip'))
        with zipfile.ZipFile(converted_zip_path, 'w') as zipf:
            for file in converted_files:
                zipf.write(file, os.path.basename(file))

        # Store the path in session
        session['converted_zip'] = converted_zip_path

        # Cleanup
        shutil.rmtree('uploads')
        os.makedirs('uploads', exist_ok=True)

        return redirect(url_for('download'))

    except Exception as e:
        return f"An error occurred: {str(e)}"
    finally:
        pythoncom.CoUninitialize()

@app.route('/download')
def download():
    if 'converted_zip' in session:
        return render_template('download.html')
    else:
        return redirect(url_for('index'))

@app.route('/download-file')
def download_file():
    if 'converted_zip' in session:
        return send_file(session['converted_zip'], as_attachment=True)
    else:
        return redirect(url_for('index'))

if __name__ == '__main__':
    # app.run(debug=True)
    socketio.run(app, debug=True)
