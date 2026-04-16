import os
import uuid
import tempfile
import shutil
from flask import Flask, request, jsonify, send_file
from processor import BudgetProcessor

app = Flask(__name__, static_folder='static', static_url_path='')

# In-memory store: token -> {dir, skill_path, excel_path}
_STORE: dict = {}


@app.route('/')
def index():
    return app.send_static_file('index.html')


@app.post('/api/generate')
def generate():
    basis_file    = request.files.get('basis')
    doc_file      = request.files.get('doc')
    template_file = request.files.get('template')
    school_name   = request.form.get('school_name', '').strip()
    fiscal_year   = request.form.get('fiscal_year', '').strip()

    missing = [
        name for name, f in [('基準表', basis_file), ('文件', doc_file), ('概算表', template_file)]
        if not f
    ]
    if missing:
        return jsonify({'error': f'缺少檔案：{", ".join(missing)}'}), 400

    tmpdir = tempfile.mkdtemp()
    try:
        basis_path    = os.path.join(tmpdir, 'basis.md')
        doc_path      = os.path.join(tmpdir, 'doc.docx')
        template_path = os.path.join(tmpdir, 'template.xlsx')

        basis_file.save(basis_path)
        doc_file.save(doc_path)
        template_file.save(template_path)

        processor = BudgetProcessor(
            basis_path, doc_path, template_path,
            school_name=school_name,
            fiscal_year=fiscal_year,
        )
        result = processor.process()

        # Persist skill file
        skill_path = os.path.join(tmpdir, 'SKILL.md')
        with open(skill_path, 'w', encoding='utf-8') as f:
            f.write(result['skill_content'])

        token = str(uuid.uuid4())
        _STORE[token] = {
            'dir':        tmpdir,
            'skill_path': skill_path,
            'excel_path': result.get('excel_path'),
        }

        return jsonify({
            'token':         token,
            'skill_content': result['skill_content'],
            'summary':       result.get('summary', {}),
            'has_excel':     bool(result.get('excel_path')),
        })

    except Exception as e:
        shutil.rmtree(tmpdir, ignore_errors=True)
        return jsonify({'error': str(e)}), 500


@app.get('/api/download/<token>/<file_type>')
def download(token: str, file_type: str):
    entry = _STORE.get(token)
    if not entry:
        return jsonify({'error': '檔案不存在或已過期'}), 404

    if file_type == 'skill':
        return send_file(
            entry['skill_path'],
            as_attachment=True,
            download_name='SKILL.md',
            mimetype='text/markdown',
        )
    if file_type == 'excel' and entry.get('excel_path'):
        return send_file(
            entry['excel_path'],
            as_attachment=True,
            download_name='概算表.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    return jsonify({'error': '檔案不存在'}), 404


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
