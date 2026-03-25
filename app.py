"""
厦门大学论文格式自动化工具 - Web 服务
"""

from flask import Flask, request, send_file, render_template, jsonify
import os
import uuid
import tempfile
import traceback
from formatter import format_thesis

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB 上限

UPLOAD_FOLDER = tempfile.mkdtemp()


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/format', methods=['POST'])
def format_document():
    if 'file' not in request.files:
        return jsonify({'error': '请上传文件'}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({'error': '文件名为空'}), 400

    if not file.filename.lower().endswith('.docx'):
        return jsonify({'error': '仅支持 .docx 格式（Word 文档）'}), 400

    options = {
        'fix_page_setup':       request.form.get('fix_page_setup', 'true') == 'true',
        'fix_styles':           request.form.get('fix_styles', 'true') == 'true',
        'fix_heading_direct':   request.form.get('fix_heading_direct', 'true') == 'true',
        'auto_detect_headings': request.form.get('auto_detect_headings', 'false') == 'true',
        'fix_body_fonts':       request.form.get('fix_body_fonts', 'true') == 'true',
        'add_headers':          request.form.get('add_headers', 'true') == 'true',
        'add_page_numbers':     request.form.get('add_page_numbers', 'true') == 'true',
        'regenerate_toc':       request.form.get('regenerate_toc', 'false') == 'true',
        'thesis_title':         request.form.get('thesis_title', ''),
    }

    uid = uuid.uuid4().hex
    input_path = os.path.join(UPLOAD_FOLDER, f'{uid}_input.docx')
    output_path = os.path.join(UPLOAD_FOLDER, f'{uid}_output.docx')

    try:
        file.save(input_path)
        changes = format_thesis(input_path, output_path, options)

        # 构建下载文件名
        base_name = file.filename[:-5]  # 去掉 .docx
        download_name = f'{base_name}_厦大格式.docx'

        return send_file(
            output_path,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'处理失败：{str(e)}'}), 500

    finally:
        if os.path.exists(input_path):
            os.remove(input_path)


@app.route('/preview', methods=['POST'])
def preview_changes():
    """预览将要进行的修改（不返回文件，只返回修改列表）"""
    if 'file' not in request.files:
        return jsonify({'error': '请上传文件'}), 400

    file = request.files['file']
    if not file.filename.lower().endswith('.docx'):
        return jsonify({'error': '仅支持 .docx 格式'}), 400

    options = {
        'fix_page_setup':       request.form.get('fix_page_setup', 'true') == 'true',
        'fix_styles':           request.form.get('fix_styles', 'true') == 'true',
        'fix_heading_direct':   request.form.get('fix_heading_direct', 'true') == 'true',
        'auto_detect_headings': request.form.get('auto_detect_headings', 'false') == 'true',
        'fix_body_fonts':       request.form.get('fix_body_fonts', 'true') == 'true',
        'add_headers':          request.form.get('add_headers', 'true') == 'true',
        'add_page_numbers':     request.form.get('add_page_numbers', 'true') == 'true',
        'regenerate_toc':       request.form.get('regenerate_toc', 'false') == 'true',
        'thesis_title':         request.form.get('thesis_title', ''),
    }

    uid = uuid.uuid4().hex
    input_path = os.path.join(UPLOAD_FOLDER, f'{uid}_input.docx')
    output_path = os.path.join(UPLOAD_FOLDER, f'{uid}_preview.docx')

    try:
        file.save(input_path)
        changes = format_thesis(input_path, output_path, options)
        return jsonify({'changes': changes, 'count': len(changes)})
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'处理失败：{str(e)}'}), 500
    finally:
        for p in [input_path, output_path]:
            if os.path.exists(p):
                os.remove(p)


if __name__ == '__main__':
    print('厦门大学论文格式工具已启动')
    print('请在浏览器打开: http://localhost:5001')
    app.run(debug=True, port=5001)
