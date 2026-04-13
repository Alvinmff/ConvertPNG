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

    # Warna standar BMKG + UNGU (Ekstrem)
    COLOR_DEFS = {
        'hijau':  {'hex': '00FF00', 'css': '#00FF00', 'label': 'Ringan'},
        'kuning': {'hex': 'FFFF00', 'css': '#FFFF00', 'label': 'Sedang'},
        'oranye': {'hex': 'FF8C00', 'css': '#FF8C00', 'label': 'Lebat'},
        'merah':  {'hex': 'FF0000', 'css': '#FF0000', 'label': 'Sangat Lebat'},
        'ungu':   {'hex': '8B00FF', 'css': '#8B00FF', 'label': 'Ekstrem'},  # Warna baru
        'abu':    {'hex': 'C0C0C0', 'css': '#C0C0C0', 'label': 'Belum Ada Data'},
        'putih':  {'hex': 'FFFFFF', 'css': '#FFFFFF', 'label': 'Kosong'},
    }

    # Target warna untuk ringkasan (Oranye, Merah, Ungu)
    HIGH_INDICATOR_COLORS = ['oranye', 'merah', 'ungu']

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
            # Deteksi warna berdasarkan Hue (HSV OpenCV: 0-179)
            if h_med < 8 or h_med > 172: 
                return 'merah'
            elif 8 <= h_med < 20: 
                return 'oranye'
            elif 20 <= h_med < 38: 
                return 'kuning'
            elif 100 <= h_med < 130: 
                return 'hijau'
            elif 130 <= h_med < 160:  # Range Ungu/Violet
                return 'ungu'
            else: 
                return 'hijau'  # Default fallback

    def _ocr_cell(self, roi, whitelist=None, psm=7):
        if roi.size == 0: return ""
        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        h, w = gray.shape[:2]
        min_height = 80
        if h < min_height:
            scale = min_height / h
            gray = cv2.resize(gray, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC)
            
        # Tambahkan blur untuk menghaluskan noise sebelum thresholding OTSU
        blurred = cv2.GaussianBlur(gray, (3, 3), 0)
        _, binary = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        # Cek apakah background hitam (jika hasil otsu terbalik). 
        # Area pinggir dipastikan adalah background setelah di-cropping aman.
        corners = [binary[0,0], binary[0,-1], binary[-1,0], binary[-1,-1]]
        if sum(corners) < 255 * 2: # Artinya lebih banyak piksel hitam di pojok
            binary = cv2.bitwise_not(binary)
            
        padded = cv2.copyMakeBorder(binary, 15, 15, 15, 15, cv2.BORDER_CONSTANT, value=255)
        config = f'--oem 3 --psm {psm}'
        if whitelist: config += f' -c tessedit_char_whitelist={whitelist}'
        try:
            return pytesseract.image_to_string(padded, config=config).strip()
        except: return ""



    def _get_best_header_numbers(self, img, h_positions, v_positions):
        nums0 = self._try_ocr_header_row(img, h_positions, v_positions, 0)
        nums2 = []
        if len(h_positions) > 3:
            nums2 = self._try_ocr_header_row(img, h_positions, v_positions, 2)
            
        valid0 = sum(1 for n in nums0 if n)
        valid2 = sum(1 for n in nums2 if n)
        
        # Ambil hasil OCR terbaik dari baris 0 atau baris 2
        if valid0 > 0 or valid2 > 0:
            return nums0 if valid0 >= valid2 else nums2
            
        # Fallback jika OCR sama sekali tidak berhasil (kasus foto sangat buram/berbeda)
        num_data_cols = len(v_positions) - 2
        return [str(i).zfill(2) for i in range(1, num_data_cols + 1)]

    def _try_ocr_header_row(self, img, h_positions, v_positions, row_idx):
        numbers = []
        if row_idx + 1 >= len(h_positions): return numbers
        y1, y2 = h_positions[row_idx], h_positions[row_idx + 1]
        for col in range(1, len(v_positions) - 1):
            x1, x2 = v_positions[col], v_positions[col + 1]
            # Tingkatkan margin cropping menjadi 15% dari tinggi/lebar cell.
            # Ini akan membuang garis tepi hitam kolom excel, sehingga Otsu threshold murni membaca background dan teks saja.
            pad_y = int((y2 - y1) * 0.15)
            pad_x = int((x2 - x1) * 0.15)
            cy1, cy2 = y1 + pad_y, y2 - pad_y
            cx1, cx2 = x1 + pad_x, x2 - pad_x
            
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

    def _analyze_high_indicators(self, table_data):
        """
        Analisis data untuk kolom 1-3 (indeks 1, 2, 3 setelah kolom nama).
        Return: dict dengan struktur {col_idx: {'oranye': [...], 'merah': [...], 'ungu': [...]}}
        """
        summary = {1: {'oranye': [], 'merah': [], 'ungu': []},
                   2: {'oranye': [], 'merah': [], 'ungu': []},
                   3: {'oranye': [], 'merah': [], 'ungu': []}}
        
        for row_data in table_data['rows']:
            # Lewati header rows (row 0, 1, 2)
            if row_data.get('is_header', False):
                continue
            
            cells = row_data.get('cells', [])
            if len(cells) < 2:
                continue
            
            # Nama kota di kolom pertama (indeks 0)
            city_name = cells[0].get('text', '').strip()
            if not city_name or city_name.lower() in ['jawa timur', 'provinsi', 'kab/kota', 'jawa  timur']:
                continue
            
            # Analisis kolom 1, 2, 3 (indeks 1, 2, 3)
            for col_idx in [1, 2, 3]:
                if col_idx < len(cells):
                    color_name = cells[col_idx].get('color_name', 'putih')
                    if color_name in self.HIGH_INDICATOR_COLORS:
                        summary[col_idx][color_name].append(city_name)
        
        # Tambahkan header tanggal ke summary
        result = {}
        for col_idx in [1, 2, 3]:
            header_val = ""
            if len(table_data['rows']) > 0 and col_idx < len(table_data['rows'][0]['cells']):
                header_val = table_data['rows'][0]['cells'][col_idx].get('text', '')
            if not header_val and len(table_data['rows']) > 2 and col_idx < len(table_data['rows'][2]['cells']):
                header_val = table_data['rows'][2]['cells'][col_idx].get('text', '')
            if not header_val:
                header_val = str(col_idx).zfill(2)
                
            result[col_idx] = {
                'tanggal': header_val,
                'data': summary[col_idx],
                'total': len(summary[col_idx]['oranye']) + len(summary[col_idx]['merah']) + len(summary[col_idx]['ungu'])
            }
        return result

    def extract(self, image_path):
        img = cv2.imread(image_path)
        if img is None: raise ValueError(f"Gagal membaca gambar: {image_path}")
        h_positions, v_positions = self._detect_grid(img)
        if len(h_positions) < 4 or len(v_positions) < 3:
            raise ValueError("Gagal mendeteksi grid tabel. Pastikan gambar jelas.")
        num_rows, num_cols = len(h_positions) - 1, len(v_positions) - 1
        footer = self._extract_footer(img, h_positions)
        header_numbers = self._get_best_header_numbers(img, h_positions, v_positions)
        
        table = {
            'title': 'Matriks Curah Hujan Kabupaten di Jawa Timur',
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
        
        # Analisis untuk ringkasan
        table['high_indicators_summary'] = self._analyze_high_indicators(table)
        return table

    def generate_excel(self, table_data, output_path):
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('Matriks Curah Hujan')
        num_cols = table_data.get('num_cols', 11)
        
        # Format definitions
        title_format = workbook.add_format({
            'bold': True, 'font_size': 14, 'font_name': 'Arial', 
            'align': 'center', 'valign': 'vcenter'
        })
        
        # Format untuk ringkasan
        summary_title_format = workbook.add_format({
            'bold': True, 'font_size': 12, 'font_name': 'Arial',
            'bg_color': '#333333', 'font_color': '#FFFFFF',
            'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        
        summary_header_format = workbook.add_format({
            'bold': True, 'font_size': 11, 'font_name': 'Arial',
            'bg_color': '#4472C4', 'font_color': '#FFFFFF',
            'align': 'left', 'valign': 'vcenter', 'border': 1
        })
        
        summary_oranye_format = workbook.add_format({
            'font_size': 10, 'font_name': 'Arial',
            'bg_color': '#FFE699', 'font_color': '#000000',  # Kuning muda untuk Oranye
            'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1
        })
        
        summary_merah_format = workbook.add_format({
            'font_size': 10, 'font_name': 'Arial',
            'bg_color': '#FFB3B3', 'font_color': '#000000',  # Merah muda
            'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1
        })
        
        summary_ungu_format = workbook.add_format({
            'font_size': 10, 'font_name': 'Arial',
            'bg_color': '#D9B3FF', 'font_color': '#000000',  # Ungu muda
            'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1
        })
        
        summary_total_format = workbook.add_format({
            'bold': True, 'font_size': 11, 'font_name': 'Arial',
            'bg_color': '#70AD47', 'font_color': '#FFFFFF',
            'align': 'center', 'valign': 'vcenter', 'border': 1
        })

        # Write Title
        if num_cols > 1: 
            worksheet.merge_range(0, 0, 0, num_cols - 1, table_data['title'], title_format)
        else: 
            worksheet.write(0, 0, table_data['title'], title_format)
        
        # Write Main Table
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
                        'bg_color': f"#{color_hex}", 'font_color': text_color, 
                        'bold': is_header or is_name_bold,
                        'font_size': 11, 'font_name': 'Arial', 'border': 1,
                        'align': 'left' if col_idx == 0 else 'center', 
                        'valign': 'vcenter', 'text_wrap': col_idx == 0
                    })
                worksheet.write(excel_row, col_idx, cell_data.get('text', ''), format_cache[f_key])

        # Set row heights untuk tabel utama
        worksheet.set_row(0, 30)
        worksheet.set_row(1, 5)
        for r in range(start_row, start_row + len(table_data['rows'])): 
            worksheet.set_row(r, 22)
        
        # Write Footer jika ada
        footer_row = start_row + len(table_data['rows']) + 1
        if table_data.get('footer'):
            f_format = workbook.add_format({
                'italic': True, 'font_size': 10, 'font_name': 'Arial', 
                'font_color': '#666666', 'align': 'center'
            })
            if num_cols > 1: 
                worksheet.merge_range(footer_row, 0, footer_row, num_cols - 1, table_data['footer'], f_format)
            else: 
                worksheet.write(footer_row, 0, table_data['footer'], f_format)
            footer_row += 2
        else:
            footer_row += 1

        # Write Summary Section (Ringkasan Indikator Tinggi)
        summary_data = table_data.get('high_indicators_summary', {})
        
        # Header Ringkasan
        summary_title_row = footer_row
        worksheet.merge_range(summary_title_row, 0, summary_title_row, 2, 
                             'RINGKASAN INDIKATOR CURAH HUJAN TINGGI - KOLOM 1 SAMPAI 3', 
                             summary_title_format)
        worksheet.set_row(summary_title_row, 25)
        
        current_row = summary_title_row + 1
        
        # Data untuk 3 kolom (horizontal layout)
        col_widths = [30, 30, 30]  # Lebar masing-masing kolom ringkasan
        
        # Headers per kolom tanggal
        for i, col_idx in enumerate([1, 2, 3]):
            if col_idx in summary_data:
                tanggal = summary_data[col_idx]['tanggal']
                worksheet.write(current_row, i, f'KOLOM: Tanggal {tanggal}', summary_header_format)
                worksheet.set_column(i, i, col_widths[i])
        
        current_row += 1
        
        # Cari tinggi maksimum untuk alignment
        max_entries = 0
        for col_idx in [1, 2, 3]:
            if col_idx in summary_data:
                data = summary_data[col_idx]['data']
                entries = max(len(data['oranye']), len(data['merah']), len(data['ungu']))
                max_entries = max(max_entries, entries)
        
        # Tulis data per warna untuk setiap kolom
        # Baris Oranye
        for i, col_idx in enumerate([1, 2, 3]):
            if col_idx in summary_data:
                data = summary_data[col_idx]['data']
                cities = data['oranye']
                count = len(cities)
                text = f"Oranye ({count}): {', '.join(cities) if cities else '-'}"
                worksheet.write(current_row, i, text, summary_oranye_format)
        
        current_row += 1
        
        # Baris Merah
        for i, col_idx in enumerate([1, 2, 3]):
            if col_idx in summary_data:
                data = summary_data[col_idx]['data']
                cities = data['merah']
                count = len(cities)
                text = f"Merah ({count}): {', '.join(cities) if cities else '-'}"
                worksheet.write(current_row, i, text, summary_merah_format)
        
        current_row += 1
        
        # Baris Ungu
        for i, col_idx in enumerate([1, 2, 3]):
            if col_idx in summary_data:
                data = summary_data[col_idx]['data']
                cities = data['ungu']
                count = len(cities)
                text = f"Ungu ({count}): {', '.join(cities) if cities else '-'}"
                worksheet.write(current_row, i, text, summary_ungu_format)
        
        current_row += 1
        
        # Baris Total
        for i, col_idx in enumerate([1, 2, 3]):
            if col_idx in summary_data:
                total = summary_data[col_idx]['total']
                text = f"TOTAL KEJADIAN: {total}"
                worksheet.write(current_row, i, text, summary_total_format)
        
        worksheet.set_row(current_row, 22)
        
        # Set column widths untuk tabel utama
        worksheet.set_column(0, 0, 25)
        if num_cols > 1: 
            worksheet.set_column(1, num_cols - 1, 5)
        
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
        
        # Update preview untuk menampilkan ringkasan juga
        preview_rows = []
        for row in table_data['rows']:
            preview_rows.append({
                'cells': [{'text': c.get('text', ''), 'color': c.get('color_css', '#FFFFFF')} for c in row['cells']],
                'isHeader': row.get('is_header', False)
            })
        
        # Buat preview ringkasan
        summary_preview = {}
        summary_data = table_data.get('high_indicators_summary', {})
        for col_idx in [1, 2, 3]:
            if col_idx in summary_data:
                summary_preview[f"kolom_{col_idx}"] = {
                    'tanggal': summary_data[col_idx]['tanggal'],
                    'oranye_count': len(summary_data[col_idx]['data']['oranye']),
                    'merah_count': len(summary_data[col_idx]['data']['merah']),
                    'ungu_count': len(summary_data[col_idx]['data']['ungu']),
                    'total': summary_data[col_idx]['total'],
                    'oranye_cities': summary_data[col_idx]['data']['oranye'][:5],  # Limit untuk preview
                    'merah_cities': summary_data[col_idx]['data']['merah'][:5],
                    'ungu_cities': summary_data[col_idx]['data']['ungu'][:5]
                }
            
        return jsonify({
            'success': True,
            'preview': {
                'title': table_data['title'], 
                'rows': preview_rows, 
                'footer': table_data.get('footer', ''),
                'summary': summary_preview
            },
            'downloadUrl': f'/download/{output_filename}',
            'stats': {
                'rows': len(table_data['rows']), 
                'cols': table_data.get('num_cols', 11), 
                'time': round(elapsed, 2)
            }
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
