import os
import subprocess
import logging
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches, Cm, Pt

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
    # Prüfe ob multipart/form-data (Datei-Upload)
    if request.content_type and request.content_type.startswith('multipart/form-data'):
        form = request.form
        files = request.files
        data = {key: form.get(key) for key in REQUIRED_FIELDS}
        excel_file = files.get('excelFile')
        excel_path = os.path.join(os.path.dirname(__file__), 'excel_data', 'bcm_uploaded.xlsx')
        if excel_file:
            excel_file.save(excel_path)
        else:
            excel_path = os.path.join(os.path.dirname(__file__), 'excel_data', 'bcm_test_source.xlsx')
    else:
        # JSON-Fallback
        data = request.get_json()
        excel_path = os.path.join(os.path.dirname(__file__), 'excel_data', 'bcm_test_source.xlsx')
    if not data:
        return jsonify({"message": "No data provided."}), 400
    missing = [f for f in REQUIRED_FIELDS if f not in data or data[f] is None]
    if missing:
        return jsonify({"message": f"Missing fields: {', '.join(missing)}"}), 400

    # Baue Argumentliste
    args = [
        "python3", os.path.join(os.path.dirname(__file__), "generate_presentation.py")
    ]
    for key in REQUIRED_FIELDS:
        args.append(f"--{key}")
        args.append(str(data[key]))
    args.append(f"--excelPath")
    args.append(excel_path)

    logging.info(f"Running: {' '.join(args)}")
    try:
        result = subprocess.run(args, capture_output=True, text=True, check=True)
        logging.info(f"stdout: {result.stdout}")
        logging.info(f"stderr: {result.stderr}")
        return jsonify({
            "message": "Presentation generated successfully.",
            "output": result.stdout.strip()
        })
    except subprocess.CalledProcessError as e:
        logging.error(f"Error: {e.stderr}")
        return jsonify({
            "message": "Error generating presentation.",
            "error": e.stderr.strip()
        }), 500


def main():
    app.run(host='0.0.0.0', port=5000, debug=True)

if __name__ == "__main__":
    main()
