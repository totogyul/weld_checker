import sys
import os
import threading
import webbrowser
import socket
import io

from flask import Flask, request, send_file, render_template, jsonify
from weld_checker import parse_table_pdf, parse_diagram_pdf, build_report

# PyInstaller로 패키징된 경우 리소스 경로 처리
if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__, template_folder=os.path.join(base_dir, 'templates'))
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB


def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/check', methods=['POST'])
def check():
    table_file   = request.files.get('table_pdf')
    diagram_file = request.files.get('diagram_pdf')

    if not table_file or not diagram_file:
        return jsonify(error='PDF 파일 두 개를 모두 업로드해 주세요.'), 400

    try:
        table_pairs   = parse_table_pdf(table_file.stream)
        diagram_pairs = parse_diagram_pdf(diagram_file.stream)
        excel_bytes   = build_report(table_pairs, diagram_pairs)
    except Exception as e:
        return jsonify(error=f'처리 중 오류가 발생했습니다: {e}'), 500

    return send_file(
        io.BytesIO(excel_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='오매칭_검사결과.xlsx',
    )


if __name__ == '__main__':
    port = find_free_port()
    url  = f'http://127.0.0.1:{port}'
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print(f'브라우저가 자동으로 열립니다: {url}')
    app.run(port=port, debug=False)
