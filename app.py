"""
BMKG Rainfall Matrix Converter — Flask Web Server
Menyediakan web GUI untuk upload gambar dan download Excel.
"""

import os
import cv2
import numpy as np
import xlsxwriter
import pytesseract
import uuid
import time
import re
from difflib import SequenceMatcher
from flask import Flask, render_template, request, jsonify, send_file

# ──────────────────────────────────────────────
# CORE LOGIC: BMKGRainfallExtractor
# ──────────────────────────────────────────────

class BMKGRainfallExtractor:
    """
    Mengekstrak tabel Matriks Curah Hujan dari gambar BMKG
    dan mengkonversi ke Excel dengan preservasi warna 100%.
    """

    # 38 Kabupaten/Kota di Jawa Timur (urutan tetap)
    KABUPATEN_LIST = [
        "Pacitan", "Ponorogo", "Trenggalek", "Tulungagung",
        "Blitar", "Kediri", "Malang", "Lumajang",
        "Jember", "Banyuwangi", "Bondowoso", "Situbondo",
        "Probolinggo", "Pasuruan", "Sidoarjo", "Mojokerto",
        "Jombang", "Nganjuk", "Madiun", "Magetan",
        "Ngawi", "Bojonegoro", "Tuban", "Lamongan",
        "Gresik", "Bangkalan", "Sampang", "Pamekasan",
        "Sumenep", "Kota Kediri", "Kota Blitar", "Kota Malang",
        "Kota Probolinggo", "Kota Pasuruan", "Kota Mojokerto",
        "Kota Madiun", "Kota Surabaya", "Kota Batu"
    ]

    # Warna standar BMKG
    COLOR_DEFS = {
        'hijau':  {'hex': '00FF00', 'css': '#00FF00', 'label': 'Ringan'},
        'kuning': {'hex': 'FFFF00', 'css': '#FFFF00', 'label': 'Sedang'},
        'oranye': {'hex': 'FF8C00', 'css': '#FF8C00', 'label': 'Lebat'},
        'merah':  {'hex': 'FF0000', 'css': '#FF0000', 'label': 'Sangat Lebat'},
        'abu':    {'hex': 'C0C0C0', 'css': '#C0C0C0', 'label': 'Belum Ada Data'},
        'putih':  {'hex': 'FFFFFF', 'css': '#FFFFFF', 'label': 'Kosong'},
    }

    def __init__(self, tesseract_cmd=None):
        if tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    def _detect_grid(self, img):
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (3, 3), 0)
        _, binary = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY_INV)
        h, w = binary.shape
        h_kernel_size = max(w // 15, 40)
        h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (h_kernel_size, 1))
        h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel)
        h_lines = cv2.dilate(h_lines, np.ones((3, 1), np.uint8), iterations=1)
        h_projection = np.sum(h_lines, axis=1).astype(float)
        h_positions = self._find_peaks(h_projection, min_gap=8)
        v_kernel_size = max(h // 30, 20)
        v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_kernel_size))
        v_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel)
        v_lines = cv2.dilate(v_lines, np.ones((1, 3), np.uint8), iterations=1)
        v_projection = np.sum(v_lines, axis=0).astype(float)
        v_positions = self._find_peaks(v_projection, min_gap=8)
        return h_positions, v_positions

    def _find_peaks(self, projection, min_gap=10, threshold_ratio=0.15):
        max_val = np.max(projection)
        if max_val == 0: return []
        threshold = max_val * threshold_ratio
        above = projection > threshold
        positions = []
        in_peak = False
        start = 0
        for i in range(len(above)):
            if above[i] and not in_peak:
                in_peak = True
                start = i
            elif not above[i] and in_peak:
                in_peak = False
                center = (start + i) // 2
                if not positions or (center - positions[-1]) >= min_gap:
                    positions.append(center)
        if in_peak:
            center = (start + len(projection) - 1) // 2
            if not positions or (center - positions[-1]) >= min_gap:
                positions.append(center)
        return positions

    def _classify_color(self, roi):
        h_roi, w_roi = roi.shape[:2]
        margin_y = max(4, h_roi // 4)
        margin_x = max(4, w_roi // 4)
        if h_roi > margin_y * 2 + 4 and w_roi > margin_x * 2 + 4:
            sample = roi[margin_y:-margin_y, margin_x:-margin_x]
        else:
            sample = roi
        if sample.size == 0: return 'putih'
        hsv = cv2.cvtColor(sample, cv2.COLOR_BGR2HSV)
        h_med = float(np.median(hsv[:, :, 0]))
        s_med = float(np.median(hsv[:, :, 1]))
        v_med = float(np.median(hsv[:, :, 2]))
        if s_med < 40:
            if v_med > 200: return 'putih'
            else: return 'abu'
        else:
            if h_med < 8 or h_med > 168: return 'merah'
            elif h_med < 23: return 'oranye'
            elif h_med < 38: return 'kuning'
            else: return 'hijau'

    def _ocr_cell(self, roi, whitelist=None, psm=7):
        if roi.size == 0: return ""
        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        h, w = gray.shape[:2]
        min_height = 80
        if h < min_height:
            scale = min_height / h
            gray = cv2.resize(gray, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC)
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        padded = cv2.copyMakeBorder(binary, 15, 15, 15, 15, cv2.BORDER_CONSTANT, value=255)
        config = f'--oem 3 --psm {psm}'
        if whitelist: config += f' -c tessedit_char_whitelist={whitelist}'
        try:
            return pytesseract.image_to_string(padded, config=config).strip()
        except: return ""

    def _ocr_header_numbers(self, img, h_positions, v_positions, header_row=0):
        rows_to_try = [header_row]
        if len(h_positions) > 3: rows_to_try.append(2)
        for try_row in rows_to_try:
            numbers = self._try_ocr_header_row(img, h_positions, v_positions, try_row)
            if sum(1 for n in numbers if n) > 0: return numbers
        num_data_cols = len(v_positions) - 2
        return [str(i).zfill(2) for i in range(1, num_data_cols + 1)]

    def _try_ocr_header_row(self, img, h_positions, v_positions, row_idx):
        numbers = []
        if row_idx + 1 >= len(h_positions): return numbers
        y1, y2 = h_positions[row_idx], h_positions[row_idx + 1]
        for col in range(1, len(v_positions) - 1):
            x1, x2 = v_positions[col], v_positions[col + 1]
            pad = 2
            cy1, cy2 = max(0, y1 + pad), max(0, y2 - pad)
            cx1, cx2 = max(0, x1 + pad), max(0, x2 - pad)
            if cy2 <= cy1 or cx2 <= cx1:
                numbers.append("")
                continue
            roi = img[cy1:cy2, cx1:cx2]
            text = ""
            for psm in [8, 7, 13]:
                text = self._ocr_cell(roi, whitelist='0123456789', psm=psm)
                text = re.sub(r'[^0-9]', '', text)
                if text: break
            numbers.append(text if text else "")
        valid_nums = [(i, int(n)) for i, n in enumerate(numbers) if n]
        if valid_nums:
            first_idx, first_num = valid_nums[0]
            filled = []
            for i in range(len(numbers)): filled.append(str(first_num - first_idx + i).zfill(2))
            numbers = filled
        return numbers

    def _extract_footer(self, img, h_positions):
        if not h_positions: return ""
        last_y, img_h = h_positions[-1], img.shape[0]
        if last_y + 10 < img_h:
            footer_roi = img[last_y + 2:img_h, :]
            if footer_roi.size > 0:
                fh, fw = footer_roi.shape[:2]
                if fh < 60:
                    scale = 60.0 / fh
                    footer_roi = cv2.resize(footer_roi, (int(fw * scale), int(fh * scale)), interpolation=cv2.INTER_CUBIC)
                try:
                    text = pytesseract.image_to_string(footer_roi, config='--oem 3 --psm 6').strip()
                    return re.sub(r'\s+', ' ', text)
                except: pass
        return ""

    def extract(self, image_path):
        img = cv2.imread(image_path)
        if img is None: raise ValueError(f"Gagal membaca gambar: {image_path}")
        h_positions, v_positions = self._detect_grid(img)
        if len(h_positions) < 4 or len(v_positions) < 3:
            raise ValueError("Gagal mendeteksi grid tabel. Pastikan gambar jelas.")
        num_rows, num_cols = len(h_positions) - 1, len(v_positions) - 1
        header_numbers = self._ocr_header_numbers(img, h_positions, v_positions, header_row=0)
        footer = self._extract_footer(img, h_positions)
        table = {
            'title': 'Matriks Curah Hujan Kabupaten di Jawa Timur',
            'headers': header_numbers,
            'num_cols': num_cols,
            'rows': [],
            'footer': footer,
        }
        data_row_index = 0
        for row in range(num_rows):
            y1, y2 = h_positions[row], h_positions[row + 1]
            row_data = {'cells': [], 'is_header': False, 'name_bold': False}
            for col in range(num_cols):
                x1, x2 = v_positions[col], v_positions[col + 1]
                color_name = 'putih'
                if col >= 1:
                    roi = img[y1+2:y2-2, x1+2:x2-2]
                    color_name = self._classify_color(roi)
                row_data['cells'].append({
                    'color_name': color_name,
                    'color_hex': self.COLOR_DEFS[color_name]['hex'],
                    'color_css': self.COLOR_DEFS[color_name]['css'],
                    'text': '',
                })
            if row == 0:
                row_data.update({'is_header': True, 'name_bold': True})
                if row_data['cells']:
                    row_data['cells'][0].update({'text': 'Provinsi', 'color_name': 'putih', 'color_hex': 'FFFFFF', 'color_css': '#FFFFFF'})
                for i, num in enumerate(header_numbers):
                    if i + 1 < len(row_data['cells']): row_data['cells'][i + 1]['text'] = num
            elif row == 1:
                if row_data['cells']: row_data['cells'][0]['text'] = 'Jawa Timur'
            elif row == 2:
                row_data.update({'is_header': True, 'name_bold': True})
                if row_data['cells']:
                    row_data['cells'][0].update({'text': 'Kab/Kota', 'color_name': 'putih', 'color_hex': 'FFFFFF', 'color_css': '#FFFFFF'})
                for i, num in enumerate(header_numbers):
                    if i + 1 < len(row_data['cells']): row_data['cells'][i + 1]['text'] = num
            else:
                if data_row_index < len(self.KABUPATEN_LIST):
                    name = self.KABUPATEN_LIST[data_row_index]
                else:
                    name_roi = img[y1+2:y2-2, v_positions[0]+2:v_positions[1]-2]
                    name = self._ocr_cell(name_roi) if name_roi.size > 0 else ""
                if row_data['cells']: row_data['cells'][0]['text'] = name
                data_row_index += 1
            table['rows'].append(row_data)
        return table

    def generate_excel(self, table_data, output_path):
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('Matriks Curah Hujan')
        num_cols = table_data.get('num_cols', 11)
        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'font_name': 'Arial', 'align': 'center', 'valign': 'vcenter'})
        if num_cols > 1: worksheet.merge_range(0, 0, 0, num_cols - 1, table_data['title'], title_format)
        else: worksheet.write(0, 0, table_data['title'], title_format)
        start_row, format_cache = 2, {}
        for row_idx, row_data in enumerate(table_data['rows']):
            excel_row = start_row + row_idx
            for col_idx, cell_data in enumerate(row_data['cells']):
                color_hex = cell_data.get('color_hex', 'FFFFFF')
                rgb_hex = color_hex[-6:]
                r, g, b = int(rgb_hex[0:2], 16), int(rgb_hex[2:4], 16), int(rgb_hex[4:6], 16)
                brightness = (r * 299 + g * 587 + b * 114) / 1000
                text_color = '#000000' if brightness > 128 else '#FFFFFF'
                is_header = row_data.get('is_header', False)
                is_name_bold = row_data.get('name_bold', False) and col_idx == 0
                f_key = (color_hex, text_color, is_header or is_name_bold, col_idx == 0)
                if f_key not in format_cache:
                    format_cache[f_key] = workbook.add_format({
                        'bg_color': f"#{color_hex}", 'font_color': text_color, 'bold': is_header or is_name_bold,
                        'font_size': 11, 'font_name': 'Arial', 'border': 1,
                        'align': 'left' if col_idx == 0 else 'center', 'valign': 'vcenter', 'text_wrap': col_idx == 0
                    })
                worksheet.write(excel_row, col_idx, cell_data.get('text', ''), format_cache[f_key])
        worksheet.set_row(0, 30)
        worksheet.set_row(1, 5)
        for r in range(start_row, start_row + len(table_data['rows'])): worksheet.set_row(r, 22)
        if table_data.get('footer'):
            f_row, f_format = start_row + len(table_data['rows']) + 1, workbook.add_format({'italic': True, 'font_size': 10, 'font_name': 'Arial', 'font_color': '#666666', 'align': 'center'})
            if num_cols > 1: worksheet.merge_range(f_row, 0, f_row, num_cols - 1, table_data['footer'], f_format)
            else: worksheet.write(f_row, 0, table_data['footer'], f_format)
        worksheet.set_column(0, 0, 25)
        if num_cols > 1: worksheet.set_column(1, num_cols - 1, 5)
        workbook.close()
        return output_path

# ──────────────────────────────────────────────
# WEB SERVER SETUP
# ──────────────────────────────────────────────

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Gunakan direktori /tmp untuk Vercel (Environment Read-Only)
if os.environ.get('VERCEL') or os.environ.get('VERCEL_ENV'):
    UPLOAD_DIR = '/tmp/uploads'
    OUTPUT_DIR = '/tmp/outputs'
else:
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
    now = time.time()
    try:
        for f in os.listdir(directory):
            path = os.path.join(directory, f)
            if os.path.isfile(path):
                if now - os.path.getmtime(path) > max_age_seconds:
                    os.remove(path)
    except: pass

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    cleanup_old_files(UPLOAD_DIR)
    cleanup_old_files(OUTPUT_DIR)
    if 'file' not in request.files: return jsonify({'success': False, 'error': 'Tidak ada file'}), 400
    file = request.files['file']
    if file.filename == '' or not allowed_file(file.filename):
        return jsonify({'success': False, 'error': 'File tidak valid'}), 400
    
    unique_id = str(uuid.uuid4())[:8]
    ext = os.path.splitext(file.filename)[1].lower()
    input_path = os.path.join(UPLOAD_DIR, f"input_{unique_id}{ext}")
    output_filename = f"output_{unique_id}.xlsx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)
    
    file.save(input_path)
    try:
        s_time = time.time()
        table_data = extractor.extract(input_path)
        extractor.generate_excel(table_data, output_path)
        elapsed = time.time() - s_time
        
        preview_rows = []
        for row in table_data['rows']:
            preview_rows.append({
                'cells': [{'text': c.get('text', ''), 'color': c.get('color_css', '#FFFFFF')} for c in row['cells']],
                'isHeader': row.get('is_header', False)
            })
            
        return jsonify({
            'success': True,
            'preview': {'title': table_data['title'], 'rows': preview_rows, 'footer': table_data.get('footer', '')},
            'downloadUrl': f'/download/{output_filename}',
            'stats': {'rows': len(table_data['rows']), 'cols': table_data.get('num_cols', 11), 'time': round(elapsed, 2)}
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
    finally:
        try: os.remove(input_path)
        except: pass

@app.route('/download/<filename>')
def download(filename):
    path = os.path.join(OUTPUT_DIR, os.path.basename(filename))
    if not os.path.exists(path): return "File tidak ditemukan", 404
    return send_file(path, as_attachment=True, download_name=filename)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
