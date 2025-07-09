import os
import subprocess
import logging
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches, Cm, Pt
from .generate_presentation import generate_from_dataframe
import pandas as pd
import uuid
from datetime import datetime

app = Flask(__name__, static_folder='static', template_folder='templates')
CORS(app)

logging.basicConfig(level=logging.INFO)

REQUIRED_FIELDS = [
    "fontSizeLevel1", "fontSizeLevel2", "colorFillLevel1", "colorFillLevel2",
    "textColorLevel1", "textColorLevel2", "borderColor", "widthLevel2", "heightLevel2"
]

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    if request.content_type and request.content_type.startswith('multipart/form-data'):
        form = request.form
        files = request.files
        data = {key: form.get(key) for key in REQUIRED_FIELDS}
        excel_file = files.get('excelFile')
        if not excel_file or not excel_file.filename:
            return jsonify({"success": False, "message": "No Excel file uploaded."}), 400
        # Generiere eindeutigen Dateinamen
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_id = uuid.uuid4().hex[:8]
        filename = f"capability_map_{ts}_{unique_id}.pptx"
        static_folder = app.static_folder or os.path.join(os.path.dirname(__file__), 'static')
        output_dir = os.path.join(static_folder, 'generated')
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, filename)
        # Speichere die hochgeladene Excel-Datei temporär
        temp_excel_path = os.path.join(output_dir, f"uploaded_{ts}_{unique_id}.xlsx")
        excel_file.save(temp_excel_path)
        excel_path = temp_excel_path
    else:
        # JSON-Fallback (optional, kann entfernt werden)
        return jsonify({"success": False, "message": "No Excel file uploaded (JSON-Fallback not supported)."}), 400
    if not data:
        return jsonify({"success": False, "message": "No data provided."}), 400
    missing = [f for f in REQUIRED_FIELDS if f not in data or data[f] is None]
    if missing:
        return jsonify({"success": False, "message": f"Missing fields: {', '.join(missing)}"}), 400
    # Baue Argumentliste
    args = [
        "python3", os.path.join(os.path.dirname(__file__), "generate_presentation.py")
    ]
    for key in REQUIRED_FIELDS:
        args.append(f"--{key}")
        args.append(str(data[key]))
    args.append(f"--excelPath")
    args.append(excel_path)
    args.append(f"--outputPath")
    args.append(output_path)
    logging.info(f"Running: {' '.join(args)}")
    try:
        result = subprocess.run(args, capture_output=True, text=True, check=True)
        logging.info(f"stdout: {result.stdout}")
        logging.info(f"stderr: {result.stderr}")
        download_url = f"/static/generated/{filename}"
        return jsonify({
            "success": True,
            "download_url": download_url
        })
        
    except subprocess.CalledProcessError as e:
        logging.error(f"Error: {e.stderr}")
        return jsonify({
            "success": False,
            "message": "Error generating presentation.",
            "error": e.stderr.strip()
        }), 500

@app.route("/healthz")
def healthz():
    return "OK", 200

def main():
    app.run(host='0.0.0.0', port=5000, debug=True)

if __name__ == "__main__":
    main()
