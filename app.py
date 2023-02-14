import os.path
from io import BytesIO
from waitress import serve
from converter import convert_xlsx_to_apkg
from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, send_file

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    input_file = request.files.get('input_file')

    if not input_file:
        return 'Server error: File not uploaded'

    deck_name, _ = os.path.splitext(input_file.filename)
    tmp_input_file = BytesIO(input_file.read())

    output_file = BytesIO()
    try:
        convert_xlsx_to_apkg(tmp_input_file, output_file, deck_name)
    except Exception as e:
        return f'Converter error: {e} ({type(e).__name__})'
    output_file.seek(0)

    deck_filename = f'{secure_filename(deck_name)}.apkg'

    return send_file(
        output_file,
        mimetype='application/zip',
        download_name=deck_filename,
        as_attachment=True,
    )


if __name__ == '__main__':
    serve(app, port=1502)
