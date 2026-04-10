"""
BMKG Rainfall Matrix Image to Excel Converter
Khusus untuk "Matriks Curah Hujan Kabupaten di Jawa Timur"
Preservasi warna 100% dari gambar ke Excel.

Requirements: pip install opencv-python openpyxl pytesseract Pillow numpy flask
"""

import cv2
import numpy as np
import xlsxwriter
import pytesseract
import pytesseract
import os
import re
import time
from difflib import SequenceMatcher


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

    # ──────────────────────────────────────────────
    # Grid Detection
    # ──────────────────────────────────────────────

    def _detect_grid(self, img):
        """
        Deteksi grid tabel menggunakan morphological line detection.
        Returns: (h_positions, v_positions)
        """
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (3, 3), 0)

        # Binary threshold — garis tabel menjadi putih
        _, binary = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY_INV)

        h, w = binary.shape

        # --- Deteksi garis horizontal ---
        h_kernel_size = max(w // 15, 40)
        h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (h_kernel_size, 1))
        h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel)
        h_lines = cv2.dilate(h_lines, np.ones((3, 1), np.uint8), iterations=1)

        h_projection = np.sum(h_lines, axis=1).astype(float)
        h_positions = self._find_peaks(h_projection, min_gap=8)

        # --- Deteksi garis vertikal ---
        v_kernel_size = max(h // 30, 20)
        v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_kernel_size))
        v_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel)
        v_lines = cv2.dilate(v_lines, np.ones((1, 3), np.uint8), iterations=1)

        v_projection = np.sum(v_lines, axis=0).astype(float)
        v_positions = self._find_peaks(v_projection, min_gap=8)

        return h_positions, v_positions

    def _find_peaks(self, projection, min_gap=10, threshold_ratio=0.15):
        """Temukan posisi garis dari profil proyeksi."""
        max_val = np.max(projection)
        if max_val == 0:
            return []

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

    # ──────────────────────────────────────────────
    # Color Classification (HSV-based)
    # ──────────────────────────────────────────────

    def _classify_color(self, roi):
        """
        Klasifikasi warna sel menggunakan HSV color space.
        Sampling dari center untuk menghindari artefak border.
        """
        h_roi, w_roi = roi.shape[:2]

        # Margin dari border
        margin_y = max(4, h_roi // 4)
        margin_x = max(4, w_roi // 4)

        if h_roi > margin_y * 2 + 4 and w_roi > margin_x * 2 + 4:
            sample = roi[margin_y:-margin_y, margin_x:-margin_x]
        else:
            sample = roi

        if sample.size == 0:
            return 'putih'

        # Konversi ke HSV
        hsv = cv2.cvtColor(sample, cv2.COLOR_BGR2HSV)

        # Gunakan median (lebih robust dari mean)
        h_med = float(np.median(hsv[:, :, 0]))
        s_med = float(np.median(hsv[:, :, 1]))
        v_med = float(np.median(hsv[:, :, 2]))

        # Klasifikasi berdasarkan HSV
        # OpenCV HSV: H=0-179, S=0-255, V=0-255
        if s_med < 40:
            # Achromatic (abu/putih)
            if v_med > 200:
                return 'putih'
            else:
                return 'abu'
        else:
            # Chromatic — klasifikasi berdasarkan hue
            if h_med < 8 or h_med > 168:
                return 'merah'
            elif h_med < 23:
                return 'oranye'
            elif h_med < 38:
                return 'kuning'
            else:
                return 'hijau'

    # ──────────────────────────────────────────────
    # OCR
    # ──────────────────────────────────────────────

    def _ocr_cell(self, roi, whitelist=None, psm=7):
        """OCR satu sel dengan preprocessing dan upscaling."""
        if roi.size == 0:
            return ""

        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)

        # Upscale gambar kecil agar Tesseract bisa membaca
        h, w = gray.shape[:2]
        min_height = 80
        if h < min_height:
            scale = min_height / h
            gray = cv2.resize(gray, (int(w * scale), int(h * scale)),
                              interpolation=cv2.INTER_CUBIC)

        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        # Padding untuk akurasi OCR
        padded = cv2.copyMakeBorder(binary, 15, 15, 15, 15,
                                     cv2.BORDER_CONSTANT, value=255)

        config = f'--oem 3 --psm {psm}'
        if whitelist:
            config += f' -c tessedit_char_whitelist={whitelist}'

        try:
            text = pytesseract.image_to_string(padded, config=config)
            return text.strip()
        except Exception:
            return ""

    def _ocr_header_numbers(self, img, h_positions, v_positions, header_row=0):
        """OCR baris header untuk mendapatkan angka tanggal."""
        # Coba header_row dulu, kalau gagal coba row 2 (sub-header)
        rows_to_try = [header_row]
        if len(h_positions) > 3:
            rows_to_try.append(2)  # Sub-header row

        for try_row in rows_to_try:
            numbers = self._try_ocr_header_row(img, h_positions, v_positions, try_row)
            valid_count = sum(1 for n in numbers if n)
            if valid_count > 0:
                return numbers

        # Fallback: kembalikan default kosong
        num_data_cols = len(v_positions) - 2
        return [str(i).zfill(2) for i in range(1, num_data_cols + 1)]

    def _try_ocr_header_row(self, img, h_positions, v_positions, row_idx):
        """Coba OCR satu baris header."""
        numbers = []

        if row_idx + 1 >= len(h_positions):
            return numbers

        y1, y2 = h_positions[row_idx], h_positions[row_idx + 1]

        for col in range(1, len(v_positions) - 1):
            x1, x2 = v_positions[col], v_positions[col + 1]

            pad = 2
            cy1 = max(0, y1 + pad)
            cy2 = max(0, y2 - pad)
            cx1 = max(0, x1 + pad)
            cx2 = max(0, x2 - pad)

            if cy2 <= cy1 or cx2 <= cx1:
                numbers.append("")
                continue

            roi = img[cy1:cy2, cx1:cx2]

            # Coba PSM 8 (single word) dulu, lalu PSM 7 (single line)
            text = ""
            for psm in [8, 7, 13]:
                text = self._ocr_cell(roi, whitelist='0123456789', psm=psm)
                text = re.sub(r'[^0-9]', '', text)
                if text:
                    break

            numbers.append(text if text else "")

        # Isi yang kosong berdasarkan pola sekuensial
        valid_nums = [(i, int(n)) for i, n in enumerate(numbers) if n]
        if valid_nums:
            first_idx, first_num = valid_nums[0]
            filled = []
            for i in range(len(numbers)):
                filled.append(str(first_num - first_idx + i).zfill(2))
            numbers = filled

        return numbers

    def _extract_footer(self, img, h_positions):
        """Ekstrak teks footer dari bawah tabel."""
        if not h_positions:
            return ""

        last_y = h_positions[-1]
        img_h = img.shape[0]

        if last_y + 10 < img_h:
            footer_roi = img[last_y + 2:img_h, :]
            if footer_roi.size > 0:
                # Upscale jika terlalu kecil
                fh, fw = footer_roi.shape[:2]
                if fh < 60:
                    scale = 60.0 / fh
                    footer_roi = cv2.resize(
                        footer_roi, (int(fw * scale), int(fh * scale)),
                        interpolation=cv2.INTER_CUBIC
                    )

                try:
                    text = pytesseract.image_to_string(
                        footer_roi, config='--oem 3 --psm 6'
                    )
                    text = text.strip()
                    text = re.sub(r'\s+', ' ', text)
                    if text:
                        return text
                except Exception:
                    pass
        return ""

    # ──────────────────────────────────────────────
    # Main Extraction Pipeline
    # ──────────────────────────────────────────────

    def extract(self, image_path):
        """
        Pipeline utama: baca gambar → deteksi grid → ekstrak warna & teks.
        Returns: dict berisi data tabel lengkap.
        """
        img = cv2.imread(image_path)
        if img is None:
            raise ValueError(f"Tidak dapat membaca gambar: {image_path}")

        # Deteksi grid
        h_positions, v_positions = self._detect_grid(img)

        if len(h_positions) < 4 or len(v_positions) < 3:
            raise ValueError(
                f"Gagal mendeteksi grid tabel. "
                f"Ditemukan {len(h_positions)} garis horizontal, "
                f"{len(v_positions)} garis vertikal. "
                f"Pastikan gambar memiliki tabel dengan garis yang jelas."
            )

        num_rows = len(h_positions) - 1
        num_cols = len(v_positions) - 1

        print(f"  Grid terdeteksi: {num_rows} baris x {num_cols} kolom")
        print(f"  Garis H: {len(h_positions)}, Garis V: {len(v_positions)}")

        # Ekstrak angka header
        header_numbers = self._ocr_header_numbers(img, h_positions, v_positions, header_row=0)
        print(f"  Header: {header_numbers}")

        # Ekstrak footer
        footer = self._extract_footer(img, h_positions)
        print(f"  Footer: {footer}")

        # Bangun data tabel
        table = {
            'title': 'Matriks Curah Hujan Kabupaten di Jawa Timur',
            'headers': header_numbers,
            'num_cols': num_cols,
            'rows': [],
            'footer': footer,
        }

        data_row_index = 0  # Counter untuk baris kab/kota

        for row in range(num_rows):
            y1 = h_positions[row]
            y2 = h_positions[row + 1]

            row_data = {
                'cells': [],
                'is_header': False,
                'name_bold': False,
            }

            # Ekstrak warna setiap sel
            for col in range(num_cols):
                x1 = v_positions[col]
                x2 = v_positions[col + 1]

                pad = 2
                cy1 = max(0, y1 + pad)
                cy2 = max(0, y2 - pad)
                cx1 = max(0, x1 + pad)
                cx2 = max(0, x2 - pad)

                # Default: Putih untuk kolom 0 (nama), klasifikasi untuk kolom data
                color_name = 'putih'
                if col >= 1: # Semua kolom data curah hujan
                    roi = img[y1+2:y2-2, x1+2:x2-2]
                    color_name = self._classify_color(roi)
                
                row_data['cells'].append({
                    'color_name': color_name,
                    'color_hex': self.COLOR_DEFS[color_name]['hex'],
                    'color_css': self.COLOR_DEFS[color_name]['css'],
                    'text': '',
                })

            # Tentukan tipe baris berdasarkan indeks
            # Row 0: Header (Provinsi)
            # Row 1: Jawa Timur summary
            # Row 2: Sub-header (Kab/Kota)
            # Row 3+: Data kab/kota

            if row == 0:
                # Header: Provinsi | angka-angka
                row_data['is_header'] = True
                row_data['name_bold'] = True
                if row_data['cells']:
                    row_data['cells'][0]['text'] = 'Provinsi'
                    row_data['cells'][0]['color_name'] = 'putih'
                    row_data['cells'][0]['color_hex'] = 'FFFFFF'
                    row_data['cells'][0]['color_css'] = '#FFFFFF'
                for i, num in enumerate(header_numbers):
                    ci = i + 1
                    if ci < len(row_data['cells']):
                        row_data['cells'][ci]['text'] = num

            elif row == 1:
                # Jawa Timur summary
                if row_data['cells']:
                    row_data['cells'][0]['text'] = 'Jawa Timur'

            elif row == 2:
                # Sub-header: Kab/Kota | angka-angka
                row_data['is_header'] = True
                row_data['name_bold'] = True
                if row_data['cells']:
                    row_data['cells'][0]['text'] = 'Kab/Kota'
                    row_data['cells'][0]['color_name'] = 'putih'
                    row_data['cells'][0]['color_hex'] = 'FFFFFF'
                    row_data['cells'][0]['color_css'] = '#FFFFFF'
                for i, num in enumerate(header_numbers):
                    ci = i + 1
                    if ci < len(row_data['cells']):
                        row_data['cells'][ci]['text'] = num

            else:
                # Baris data kab/kota
                if data_row_index < len(self.KABUPATEN_LIST):
                    name = self.KABUPATEN_LIST[data_row_index]
                else:
                    # Fallback OCR jika lebih dari 38
                    name_x1 = v_positions[0] + 2
                    name_x2 = v_positions[1] - 2
                    name_roi = img[max(0, y1+2):max(0, y2-2),
                                   max(0, name_x1):max(0, name_x2)]
                    name = self._ocr_cell(name_roi) if name_roi.size > 0 else ""

                if row_data['cells']:
                    row_data['cells'][0]['text'] = name
                data_row_index += 1

            table['rows'].append(row_data)

        return table

    # ──────────────────────────────────────────────
    # Excel Generation
    # ──────────────────────────────────────────────

    def generate_excel(self, table_data, output_path):
        """Generate file Excel dari data tabel yang diekstrak menggunakan xlsxwriter."""
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('Matriks Curah Hujan')

        num_cols = table_data.get('num_cols', 11)

        # --- Judul ---
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'font_name': 'Arial',
            'align': 'center',
            'valign': 'vcenter'
        })
        if num_cols > 1:
            worksheet.merge_range(0, 0, 0, num_cols - 1, table_data['title'], title_format)
        else:
            worksheet.write(0, 0, table_data['title'], title_format)

        # --- Data tabel mulai baris 3 (index 2) ---
        start_row = 2

        format_cache = {}

        for row_idx, row_data in enumerate(table_data['rows']):
            excel_row = start_row + row_idx

            for col_idx, cell_data in enumerate(row_data['cells']):
                cell_text = cell_data.get('text', '')

                # Background color
                color_hex = cell_data.get('color_hex', 'FFFFFF')
                rgb_hex = color_hex[-6:]
                
                r = int(rgb_hex[0:2], 16)
                g = int(rgb_hex[2:4], 16)
                b = int(rgb_hex[4:6], 16)
                brightness = (r * 299 + g * 587 + b * 114) / 1000
                text_color = '#000000' if brightness > 128 else '#FFFFFF'

                is_header = row_data.get('is_header', False)
                is_name_bold = row_data.get('name_bold', False) and col_idx == 0

                format_key = (color_hex, text_color, is_header or is_name_bold, col_idx == 0)
                if format_key not in format_cache:
                    cell_format = workbook.add_format({
                        'bg_color': f"#{color_hex}",
                        'font_color': text_color,
                        'bold': is_header or is_name_bold,
                        'font_size': 11,
                        'font_name': 'Arial',
                        'border': 1,
                        'align': 'left' if col_idx == 0 else 'center',
                        'valign': 'vcenter',
                        'text_wrap': col_idx == 0
                    })
                    format_cache[format_key] = cell_format

                worksheet.write(excel_row, col_idx, cell_text, format_cache[format_key])

        # --- Tinggi baris ---
        worksheet.set_row(0, 30)   # Judul
        worksheet.set_row(1, 5)    # Gap
        for row in range(start_row, start_row + len(table_data['rows'])):
            worksheet.set_row(row, 22)

        # --- Footer ---
        if table_data.get('footer'):
            footer_row = start_row + len(table_data['rows']) + 1
            footer_format = workbook.add_format({
                'italic': True,
                'font_size': 10,
                'font_name': 'Arial',
                'font_color': '#666666',
                'align': 'center'
            })
            if num_cols > 1:
                worksheet.merge_range(footer_row, 0, footer_row, num_cols - 1, table_data['footer'], footer_format)
            else:
                worksheet.write(footer_row, 0, table_data['footer'], footer_format)

        # --- Lebar kolom ---
        worksheet.set_column(0, 0, 25)
        if num_cols > 1:
            worksheet.set_column(1, num_cols - 1, 5)

        # Simpan
        workbook.close()
        return output_path

    # ──────────────────────────────────────────────
    # Full Pipeline
    # ──────────────────────────────────────────────

    def process(self, image_path, output_path):
        """Pipeline lengkap: gambar -> Excel."""
        print(f"[INFO] Membaca gambar: {image_path}")

        start_time = time.time()
        table_data = self.extract(image_path)

        print(f"[INFO] Membuat Excel: {output_path}")
        self.generate_excel(table_data, output_path)

        elapsed = time.time() - start_time
        print(f"[OK] Selesai dalam {elapsed:.1f}s! Disimpan: {output_path}")

        return table_data


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description='BMKG Rainfall Matrix Image to Excel Converter'
    )
    parser.add_argument('input', help='Path file gambar (jpg, png)')
    parser.add_argument('-o', '--output', default='output.xlsx',
                        help='Path output Excel')
    parser.add_argument('--tesseract',
                        help='Path ke executable tesseract')

    args = parser.parse_args()

    try:
        extractor = BMKGRainfallExtractor(tesseract_cmd=args.tesseract)
        extractor.process(args.input, args.output)
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
        return 1

    return 0


if __name__ == "__main__":
    exit(main())