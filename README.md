<p align="center">
  <h1 align="center">🌧️ BMKG Rainfall Matrix Converter</h1>
  <p align="center">
    <strong>Automated image-to-Excel converter for BMKG East Java Rainfall Matrix tables with 100% color preservation.</strong>
  </p>
  <p align="center">
    <img src="https://img.shields.io/badge/python-3.9+-blue?style=flat-square&logo=python&logoColor=white" alt="Python">
    <img src="https://img.shields.io/badge/flask-2.0+-green?style=flat-square&logo=flask&logoColor=white" alt="Flask">
    <img src="https://img.shields.io/badge/opencv-4.5+-red?style=flat-square&logo=opencv&logoColor=white" alt="OpenCV">
    <img src="https://img.shields.io/badge/tesseract-OCR-orange?style=flat-square" alt="Tesseract">
    <img src="https://img.shields.io/badge/deploy-vercel-black?style=flat-square&logo=vercel&logoColor=white" alt="Vercel">
  </p>
</p>

---

## 📋 Overview

**BMKG Rainfall Matrix Converter** is a web-based tool designed specifically for [BMKG (Badan Meteorologi, Klimatologi, dan Geofisika)](https://www.bmkg.go.id/) East Java to automate the conversion of rainfall matrix table images into structured, color-coded Excel spreadsheets.

The application uses **Computer Vision (OpenCV)** for grid detection and color classification, combined with **Tesseract OCR** for text recognition, to extract tabular data from screenshots or photographs of the *"Matriks Curah Hujan Kabupaten di Jawa Timur"* — while preserving the original BMKG color-coding system with 100% accuracy.

### ✨ Key Features

- **Intelligent Grid Detection** — Automatic table structure recognition using morphological line detection
- **HSV-Based Color Classification** — Accurate identification of BMKG rainfall severity colors
- **OCR Text Extraction** — Reads date headers and region names from the image
- **100% Color Preservation** — Generated Excel files retain the exact BMKG color scheme
- **Modern Web Interface** — Dark-themed, glassmorphism UI with drag-and-drop upload
- **Built-in Usage Guide** — Interactive step-by-step tutorial modal for new users
- **Cloud Ready** — Optimized for serverless deployment on Vercel

---

## 🎨 BMKG Color Legend

| Color | Label | Description |
|:---:|:---|:---|
| 🟩 `#00FF00` | **Ringan** | Light rainfall |
| 🟨 `#FFFF00` | **Sedang** | Moderate rainfall |
| 🟧 `#FF8C00` | **Lebat** | Heavy rainfall |
| 🟥 `#FF0000` | **Sangat Lebat** | Very heavy rainfall |
| ⬜ `#C0C0C0` | **Belum Ada Data** | No data available |
| ⬜ `#FFFFFF` | **Kosong** | Empty / No rain |

---

## 📁 Project Structure

```
ConvertPNG/
├── app.py                  # Main application (Flask server + extraction engine)
├── requirements.txt        # Python dependencies
├── static/
│   ├── css/
│   │   └── style.css       # UI stylesheet (glassmorphism theme)
│   └── js/
│       └── script.js       # Client-side logic (upload, preview, conversion)
├── templates/
│   └── index.html          # Main HTML template
├── uploads/                # Temporary upload storage (auto-created)
├── outputs/                # Temporary Excel output storage (auto-created)
└── README.md
```

---

## 🚀 Getting Started

### Prerequisites

- **Python** 3.9 or higher
- **Tesseract OCR** installed on your system

#### Installing Tesseract OCR

<details>
<summary><strong>Windows</strong></summary>

1. Download the installer from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki).
2. Run the installer and note the installation path (default: `C:\Program Files\Tesseract-OCR`).
3. Add the installation path to your system `PATH` environment variable.
</details>

<details>
<summary><strong>macOS</strong></summary>

```bash
brew install tesseract
```
</details>

<details>
<summary><strong>Linux (Ubuntu/Debian)</strong></summary>

```bash
sudo apt update && sudo apt install tesseract-ocr
```
</details>

### Installation

1. **Clone the repository**

   ```bash
   git clone https://github.com/Alvinmff/ConvertPNG.git
   cd convert-png
   ```

2. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**

   ```bash
   python app.py
   ```

4. **Open in your browser**

   Navigate to [http://localhost:5000](http://localhost:5000)

---

## 💡 How to Use

1. **Prepare your image** — Ensure the rainfall matrix table image is clear, well-lit, and not cropped or skewed.
2. **Upload** — Drag and drop the image onto the upload zone, or click to browse your files.
3. **Convert** — Click the **Konversi Sekarang** button. The system will automatically detect the grid, classify colors, and extract text using OCR.
4. **Download** — Review the table preview, then click **Download Excel** to save the `.xlsx` file with full color formatting.

---

## ☁️ Deployment (Vercel)

This application is optimized for serverless deployment on **Vercel**.

### Key Cloud Adaptations

| Concern | Solution |
|:---|:---|
| Read-only filesystem | Upload/output directories auto-redirect to `/tmp` |
| GUI dependencies | Uses `opencv-python-headless` (no X11/GUI required) |
| Tesseract OCR | Requires buildpack or system-level installation |

### Deploy to Vercel

1. Push your repository to GitHub.
2. Import the project from [vercel.com/new](https://vercel.com/new).
3. Vercel will auto-detect the Python (Flask) framework.
4. The application will be live at your assigned `.vercel.app` domain.

> **Note:** Tesseract OCR must be available in the Vercel build environment. You may need to configure a custom build step or use a community buildpack to install the `tesseract-ocr` binary.

---

## 🛠️ Tech Stack

| Layer | Technology |
|:---|:---|
| **Backend** | Python, Flask |
| **Computer Vision** | OpenCV (Headless) |
| **OCR Engine** | Tesseract (via pytesseract) |
| **Excel Generation** | XlsxWriter |
| **Frontend** | HTML5, Vanilla CSS, Vanilla JS |
| **Typography** | Inter (Google Fonts) |
| **Deployment** | Vercel (Serverless) |

---

## 🗺️ Coverage

This tool is preconfigured with the **38 Kabupaten/Kota** (regencies/cities) of **East Java Province**, Indonesia:

<details>
<summary>View full list</summary>

| # | Kabupaten / Kota |
|:---:|:---|
| 1 | Pacitan |
| 2 | Ponorogo |
| 3 | Trenggalek |
| 4 | Tulungagung |
| 5 | Blitar |
| 6 | Kediri |
| 7 | Malang |
| 8 | Lumajang |
| 9 | Jember |
| 10 | Banyuwangi |
| 11 | Bondowoso |
| 12 | Situbondo |
| 13 | Probolinggo |
| 14 | Pasuruan |
| 15 | Sidoarjo |
| 16 | Mojokerto |
| 17 | Jombang |
| 18 | Nganjuk |
| 19 | Madiun |
| 20 | Magetan |
| 21 | Ngawi |
| 22 | Bojonegoro |
| 23 | Tuban |
| 24 | Lamongan |
| 25 | Gresik |
| 26 | Bangkalan |
| 27 | Sampang |
| 28 | Pamekasan |
| 29 | Sumenep |
| 30 | Kota Kediri |
| 31 | Kota Blitar |
| 32 | Kota Malang |
| 33 | Kota Probolinggo |
| 34 | Kota Pasuruan |
| 35 | Kota Mojokerto |
| 36 | Kota Madiun |
| 37 | Kota Surabaya |
| 38 | Kota Batu |
</details>

---

## 📄 License

This project was developed as part of an internship program at **BMKG Stasiun Meteorologi Juanda, Surabaya** (2026).

---

<p align="center">
  <sub>Built with ☁️ by <strong>Muhammad Alvino Firmansyah</strong> — BMKG Juanda Internship 2026</sub>
</p>
