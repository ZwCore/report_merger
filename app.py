
from flask import Flask, request, send_file, jsonify, render_template
from flask_cors import CORS
import os
import shutil
import datetime
import traceback
from merger import merge_reports

import sys

app = Flask(__name__)
CORS(app)

# Use absolute path for uploads to be safe
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    BASE_DIR = os.path.dirname(sys.executable)
    # For templates/static, if packaged with --add-data, they might be in sys._MEIPASS
    # But FLASK looks for them relative to root_path.
    # We should set template_folder and static_folder explicitly if needed,
    # OR simpler: Keep templates/static along with the exe?
    
    # Actually, for --onefile, sys._MEIPASS is where internal assets are.
    # sys.executable is where the exe is.
    # We want UPLOADS to be near the EXE (user visible), not in temp _MEIPASS.
    EXE_DIR = os.path.dirname(sys.executable)
    UPLOAD_FOLDER = os.path.join(EXE_DIR, 'uploads')
    
    # For Flask to find templates inside the Meipass:
    # We rely on Flask's default behavior relative to app.root_path.
    # If app is run from packaged script, __file__ might point correctly or not.
    # A robust way for Flask in PyInstaller:
    app = Flask(__name__, template_folder=os.path.join(sys._MEIPASS, 'templates'), static_folder=os.path.join(sys._MEIPASS, 'static'))
    CORS(app)
else:
    # Running as script
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/process', methods=['POST'])
def process_reports():
    try:
        # Clean upload folder first to avoid conflicts
        # Be careful not to delete currently processing files in a race condition, 
        # but for single user demo this is fine.
        for f in os.listdir(UPLOAD_FOLDER):
            file_path = os.path.join(UPLOAD_FOLDER, f)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"Error deleting {file_path}: {e}")

        if 'template' not in request.files:
            return jsonify({"error": "No template uploaded"}), 400
        
        template = request.files['template']
        template_filename = "template.docx"
        template_path = os.path.join(UPLOAD_FOLDER, template_filename)
        template.save(template_path)
        
        files = request.files.getlist('reports')
        # Debug print
        print(f"Received template: {template.filename}")
        print(f"Received {len(files)} reports")

        saved_files = []
        for file in files:
            if file.filename:
                # Ensure we strictly use the filename expected by the merger (Key.docx)
                # The frontend should likely send filenames correctly, 
                # but if the user uploads 'SecureTest (1).docx', it might fail matching '{{SecureTest}}'.
                # For now, we assume the user names files correctly or we might need a mapping logic.
                save_path = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(save_path)
                saved_files.append(file.filename)
            
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        output_filename = f"Final_Report_{timestamp}.docx"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)
        
        # Verify template exists
        if not os.path.exists(template_path):
             return jsonify({"error": "Template failed to save"}), 500

        merge_reports(template_path, UPLOAD_FOLDER, output_path)
        
        if not os.path.exists(output_path):
            return jsonify({"error": "Merger failed to generate output file"}), 500
            
        return jsonify({"downloadUrl": f"/api/download/{output_filename}"})

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    try:
        return send_file(os.path.join(UPLOAD_FOLDER, filename), as_attachment=True)
    except Exception as e:
         return jsonify({"error": f"File not found: {e}"}), 404

import logging

# Configure logging to file
log_file = os.path.join(os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__)), 'app.log')
logging.basicConfig(filename=log_file, level=logging.DEBUG, 
                    format='%(asctime)s %(levelname)s: %(message)s')

# Redirect stdout/stderr to logging
class StreamToLogger(object):
   def __init__(self, logger, log_level=logging.INFO):
      self.logger = logger
      self.log_level = log_level
   def write(self, buf):
      for line in buf.rstrip().splitlines():
         self.logger.log(self.log_level, line.rstrip())
   def flush(self):
      pass

sys.stdout = StreamToLogger(logging.getLogger("STDOUT"), logging.INFO)
sys.stderr = StreamToLogger(logging.getLogger("STDERR"), logging.ERROR)

import webbrowser
from threading import Timer

def open_browser():
      webbrowser.open_new("http://127.0.0.1:5000")

if __name__ == '__main__':
    logging.info("Starting Flask server...")
    # Open browser after 1.5 seconds to allow server to start
    Timer(1.5, open_browser).start()
    try:
        app.run(debug=False, port=5000)
    except Exception as e:
        logging.critical(f"Server crashed: {e}")
