# 🖼️ Automated MRTG Image to Excel Report (Image Only)

**Script otomatis yang difokuskan pada penempatan gambar MRTG secara presisi ke dalam *template* laporan Excel bulanan tanpa proses OCR.**

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://python.org)
[![Pillow](https://img.shields.io/badge/Pillow-10.x-green)](https://python-pillow.org/)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-3.x-yellow)](https://openpyxl.readthedocs.io/)

---

## 📌 Fitur Utama

- ✅ **Penempatan Gambar Super Cepat** – Memasukkan gambar ke dalam Excel tanpa overhead *Optical Character Recognition* (OCR).
- ✅ **Resize Proporsional Dinamis** – Menyesuaikan ukuran gambar (*stretch*/*resize*) secara presisi agar pas dengan area sel Excel yang telah ditentukan di konfigurasi.
- ✅ **Mapping Area Fleksibel** – Menggunakan file teks sederhana untuk memetakan ID gambar ke sel awal dan akhir (misal: `B12-L23`).
- ✅ **Multi-Sheet Harian** – Otomatis mendeteksi folder tanggal dan membuat *sheet* harian (1-31) di dalam *file* Excel laporan Anda.

---

## 🛠️ Prasyarat

| Software | Keterangan |
|----------|-------------|
| **Python 3.8+** | [Download](https://www.python.org/downloads/) |
| **Template Excel** | File `MENTAHAN FORMAT DAILY MRTG.xlsx` |

*(Catatan: Versi ini tidak membutuhkan instalasi Tesseract).*

---

## 📦 Instalasi

1. **Clone repository** (atau pindah ke folder proyek)
   ```bash
   cd Automated-MRTG-to-Excel-Report-just_img
   ```

2. **Buat virtual environment (opsional tapi disarankan)**
   ```bash
   python -m venv venv
   venv\Scripts\activate      # Windows
   ```

3. **Install library Python yang dibutuhkan**
   ```bash
   pip install openpyxl pillow
   ```

---

## 📁 Persiapan File & Struktur Folder

Sebelum menjalankan script, pastikan file dan folder berikut tersedia di dalam folder yang sama dengan script `script_ini.py`:

1. **`list_mrtg_data.txt`**  
   Berisi daftar urutan SID atau Graph Title yang akan di-*insert*.
2. **`sid_image-position-excel.txt`**  
   File konfigurasi yang mendefinisikan lokasi sel (area letak gambar) di Excel untuk setiap ID. Contoh format:
   ```text
   SID : 4700001-0021497479
   -> B12-L23
   ```
3. **`MENTAHAN FORMAT DAILY MRTG.xlsx`**  
   File *template* laporan Excel kosong.
4. **Folder `MRTG-Data/`**  
   Folder utama yang di dalamnya terdapat folder berformat tanggal `YYYYMMDD` (contoh: `20260101`) berisi gambar-gambar MRTG.

---

## 🚀 Cara Penggunaan

1. Pastikan semua gambar sudah tersimpan rapi berdasarkan tanggal di folder `MRTG-Data/`.
2. Jalankan script di terminal:
   ```bash
   python script_ini.py
   ```
3. Script akan langsung memetakan semua gambar dari setiap folder tanggal ke sheet masing-masing di file Excel.
4. **Hasil akhir** akan tersimpan dalam *file* baru bernama `Daily_Report_Complete.xlsx`.

---

**Cepat dan Presisi! Selamat menyelesaikan pelaporan harian/bulanan Anda! 🚀**