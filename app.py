"""
BMKG Rainfall Matrix Converter — Flask Web Server
Menyediakan web GUI untuk upload gambar dan download Excel.
"""

from flask import Flask, render_template, request, jsonify, send_file
import os
import uuid
import time
from convert import BMKGRainfallExtractor

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'outputs')

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

extractor = BMKGRainfallExtractor()

ALLOWED_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp'}


def allowed_file(filename):
    ext = os.path.splitext(filename)[1].lower()
    return ext in ALLOWED_EXTENSIONS


def cleanup_old_files(directory, max_age_seconds=3600):
    """Hapus file yang lebih lama dari max_age_seconds."""
    now = time.time()
    try:
        for f in os.listdir(directory):
            path = os.path.join(directory, f)
            if os.path.isfile(path):
                age = now - os.path.getmtime(path)
                if age > max_age_seconds:
                    os.remove(path)
    except Exception:
        pass


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    # Cleanup file lama
    cleanup_old_files(UPLOAD_DIR)
    cleanup_old_files(OUTPUT_DIR)

    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'Tidak ada file yang di-upload'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'Tidak ada file yang dipilih'}), 400

    if not allowed_file(file.filename):
        return jsonify({
            'success': False,
            'error': 'Format file tidak didukung. Gunakan JPG atau PNG.'
        }), 400

    # Simpan file upload
    unique_id = str(uuid.uuid4())[:8]
    ext = os.path.splitext(file.filename)[1].lower()
    input_filename = f"input_{unique_id}{ext}"
    output_filename = f"output_{unique_id}.xlsx"

    input_path = os.path.join(UPLOAD_DIR, input_filename)
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    file.save(input_path)

    try:
        start_time = time.time()

        # Ekstraksi dan konversi
        table_data = extractor.extract(input_path)
        extractor.generate_excel(table_data, output_path)

        elapsed = time.time() - start_time

        # Build preview data untuk frontend
        preview_rows = []
        for row in table_data['rows']:
            cells = []
            for cell in row['cells']:
                cells.append({
                    'text': cell.get('text', ''),
                    'color': cell.get('color_css', '#FFFFFF'),
                    'colorName': cell.get('color_name', 'putih'),
                })
            preview_rows.append({
                'cells': cells,
                'isHeader': row.get('is_header', False),
            })

        return jsonify({
            'success': True,
            'preview': {
                'title': table_data['title'],
                'rows': preview_rows,
                'footer': table_data.get('footer', ''),
            },
            'downloadUrl': f'/download/{output_filename}',
            'stats': {
                'rows': len(table_data['rows']),
                'cols': table_data.get('num_cols', 11),
                'time': round(elapsed, 2),
            }
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

    finally:
        # Hapus file input
        try:
            os.remove(input_path)
        except Exception:
            pass


@app.route('/download/<filename>')
def download(filename):
    # Sanitasi nama file
    filename = os.path.basename(filename)
    path = os.path.join(OUTPUT_DIR, filename)

    if not os.path.exists(path):
        return "File tidak ditemukan", 404

    return send_file(
        path,
        as_attachment=True,
        download_name=filename
    )


if __name__ == '__main__':
    print("=" * 50)
    print("[INFO] BMKG Rainfall Matrix Converter")
    print("=" * 50)
    print(f"[INFO] Upload dir: {UPLOAD_DIR}")
    print(f"[INFO] Output dir: {OUTPUT_DIR}")
    print(f"[INFO] Buka http://localhost:5000 di browser")
    print("=" * 50)
    app.run(debug=True, port=5000)
