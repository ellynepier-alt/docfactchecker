import os, tempfile, uuid
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from werkzeug.utils import secure_filename
from factcheck_engine import SUPPORTED_EXTS, load_kb, run_checks, make_report

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
KB_PATH = os.path.join(BASE_DIR, 'doc_guidelines_kb.json')
MAX_UPLOAD_MB = 25

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'change-this-before-production')
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_MB * 1024 * 1024


def allowed(filename):
    return os.path.splitext(filename)[1].lower() in SUPPORTED_EXTS


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('material')
        if not file or file.filename == '':
            flash('Please choose a file to check.')
            return redirect(url_for('index'))
        if not allowed(file.filename):
            flash('Unsupported file type. Use PDF, DOCX, TXT, or MD.')
            return redirect(url_for('index'))

        safe_name = secure_filename(file.filename)
        temp_dir = tempfile.mkdtemp(prefix='doc_factcheck_')
        in_path = os.path.join(temp_dir, safe_name)
        file.save(in_path)
        try:
            kb = load_kb(KB_PATH)
            result = run_checks(in_path, kb)
            out_path = os.path.join(temp_dir, f"DoC_factcheck_{uuid.uuid4().hex[:8]}.docx")
            make_report(result, kb, out_path)
            return send_file(out_path, as_attachment=True, download_name='DoC_factcheck_report.docx')
        except Exception as e:
            flash(f'Could not process this file: {e}')
            return redirect(url_for('index'))

    return render_template('index.html', max_mb=MAX_UPLOAD_MB)


if __name__ == '__main__':
    app.run(debug=False, host='127.0.0.1', port=5000)
