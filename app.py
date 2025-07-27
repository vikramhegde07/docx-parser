# app.py
from flask import Flask, request, jsonify
from werkzeug.utils import secure_filename
import os
from dotenv import load_dotenv
from parser.docx_processor import parse_docx
from flask_cors import CORS

load_dotenv()

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.getenv("UPLOAD_FOLDER", "uploads")
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

CORS(app)

@app.route('/',methods=["GET"])
def serverTest():
    return "Server is up and running"

@app.route("/parse-docx", methods=["POST"])
def parse_docx_route():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    try:
        upload_dir = os.path.join("static", "images")  # or any directory where you want to store extracted images
        os.makedirs(upload_dir, exist_ok=True)

        result = parse_docx(filepath, upload_dir)  
        return jsonify(result), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

port = int(os.getenv('PORT',5000))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=port,debug=True)
