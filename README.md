# 🧠 Pengenalan Emosi Wajah Secara Real-Time (saat Belajar)
### Menggunakan Python & OpenCV — Tutorial dengan Kode Sumber (2025)

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge&logo=python&logoColor=white"/>
  <img src="https://img.shields.io/badge/OpenCV-4.10-green?style=for-the-badge&logo=opencv&logoColor=white"/>
  <img src="https://img.shields.io/badge/DeepFace-0.0.93-orange?style=for-the-badge"/>
  <img src="https://img.shields.io/badge/MediaPipe-0.10-red?style=for-the-badge&logo=google&logoColor=white"/>
  <img src="https://img.shields.io/badge/Flask-3.0-lightgrey?style=for-the-badge&logo=flask&logoColor=white"/>
</p>

<p align="center">
  Sistem monitoring ekspresi wajah mahasiswa secara real-time saat mengikuti perkuliahan,
  menggunakan <i>computer vision</i> untuk mendeteksi dan mengklasifikasikan 7 emosi dasar
  langsung dari kamera laptop maupun webcam.
</p>

---

## 📸 Tampilan Sistem

> Webcam stream tampil besar di kiri · Grafik emosi & log real-time di kanan · Pengenalan wajah otomatis

---

## 🎯 Tujuan

Menggunakan *computer vision* untuk mengenali ekspresi wajah secara real-time guna membantu dosen memahami kondisi emosional mahasiswa selama proses pembelajaran berlangsung.

**Cara Kerja Singkat:**
```
Webcam → OpenCV (deteksi wajah)
       → DeepFace/ArcFace (klasifikasi emosi + pengenalan wajah)
       → MediaPipe Pose (deteksi angkat tangan / bertanya)
       → Flask (stream MJPEG ke browser)
       → SQLite (simpan log otomatis tiap 5 detik)
       → Ekspor Excel (laporan harian / mingguan / bulanan / semester)
```

---

## ✨ Fitur Utama

| Fitur | Keterangan |
|---|---|
| 📡 **Deteksi 7 Emosi Real-Time** | Marah, Jijik, Takut, Senang, Sedih, Terkejut, Netral |
| 🧑‍🎓 **Pengenalan Wajah Otomatis** | Identifikasi mahasiswa dari foto referensi (model ArcFace) |
| ✋ **Deteksi Angkat Tangan** | Mendeteksi mahasiswa yang sedang bertanya ke dosen (MediaPipe Pose) |
| 📊 **Grafik Live** | Bar distribusi emosi + Pie chart kumulatif per sesi |
| 📋 **Log Deteksi** | Riwayat lengkap emosi + waktu + confidence score |
| 💾 **Auto-Save ke Database** | Disimpan ke SQLite setiap 5 detik otomatis |
| 📥 **Ekspor Excel** | Laporan lengkap (harian / mingguan / bulanan / 3 bulan / 1 semester) |
| 👥 **Manajemen Data Siswa** | Upload foto wajah mahasiswa untuk pengenalan otomatis |
| 🎓 **Filter per Mata Kuliah** | Laporan bisa difilter per mata kuliah tertentu |

---

## 🗂️ Struktur Folder

```
emosi-v3/
├── app.py                  # Aplikasi utama Flask
├── requirements.txt        # Dependensi library
├── README.md
├── templates/
│   ├── index.html          # Halaman monitoring real-time
│   └── mahasiswa.html      # Halaman kelola data siswa
├── static/                 # Aset statis (opsional)
└── data/
    ├── log_emosi.db        # Database SQLite (otomatis dibuat)
    └── wajah_db/           # Folder foto wajah mahasiswa
```

---

## ⚙️ Cara Kerja Teknis

### Deteksi Wajah
Menggunakan **OpenCV Haar Cascade** (`haarcascade_frontalface_default.xml`) untuk mendeteksi lokasi wajah di setiap frame kamera secara ringan dan cepat.

### Klasifikasi Emosi
Menggunakan **DeepFace** dengan backend OpenCV. Hasil analisis berupa probabilitas untuk 7 emosi dasar. Sistem menggunakan **Majority Vote (window 10 frame)** untuk menstabilkan hasil deteksi agar tidak berubah-ubah terlalu cepat.

### Threshold Confidence per Emosi
Setiap emosi memiliki ambang batas berbeda untuk mencegah *false positive*:

| Emosi | Threshold | Boost Multiplier |
|---|---|---|
| 😠 Marah | 28% | ×1.35 |
| 🤢 Jijik | 25% | ×1.40 |
| 😨 Takut | 30% | ×1.30 |
| 😊 Senang | 38% | ×1.00 |
| 😢 Sedih | 78% | ×0.75 (diperketat) |
| 😲 Terkejut | 28% | ×1.30 |
| 😐 Netral | 18% | ×1.00 (fallback) |

### Pengenalan Wajah (Face Recognition)
Menggunakan **DeepFace.find()** dengan model **ArcFace** dan metric jarak *cosine* (threshold ≤ 0.45). Sebelum dicocokkan, frame dipreprocess terlebih dahulu:
1. Crop area wajah menggunakan Haar Cascade
2. CLAHE (Contrast Limited Adaptive Histogram Equalization) untuk normalisasi pencahayaan
3. Resize ke 224×224 piksel (input standar ArcFace)
4. Sharpening ringan

### Deteksi Angkat Tangan
Menggunakan **MediaPipe Pose** untuk mendeteksi pose tulang tubuh. Tangan dianggap terangkat jika koordinat Y pergelangan tangan lebih tinggi dari bahu (threshold 0.08 dalam koordinat normalized).

---

## 🚀 Instalasi & Menjalankan

### 1. Clone Repositori
```bash
git clone https://github.com/username/nama-repo.git
cd nama-repo
```

### 2. Buat Virtual Environment
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Mac / Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependensi
```bash
pip install -r requirements.txt
```
> ⏳ Proses install pertama kali membutuhkan waktu **10–20 menit** karena DeepFace akan mengunduh model AI secara otomatis.

### 4. Jalankan Aplikasi
```bash
python app.py
```

### 5. Buka Browser
```
http://localhost:5000
```

---

## 📖 Panduan Penggunaan

### Halaman Monitoring (`/`)

1. Isi **Nama Mahasiswa**, **Kelas**, dan **Mata Kuliah**
2. Klik **▶ Mulai** — kamera akan aktif dan deteksi berjalan otomatis
3. Panel kanan menampilkan:
   - **Tab Grafik** — distribusi emosi live + pie chart kumulatif
   - **Tab Log** — riwayat deteksi per waktu
   - **Tab Ekspor** — unduh laporan Excel dengan filter periode & mata kuliah
4. Klik **⏹ Stop** untuk menghentikan kamera
5. Data tersimpan otomatis ke database SQLite setiap **5 detik**

### Halaman Data Siswa (`/siswa`)

1. Isi **Nama**, **Kelas**, **NIM/NIS**
2. Upload **foto wajah** yang jelas (JPEG/PNG)
   - Foto akan otomatis diproses: crop wajah → CLAHE enhance → resize 224px
3. Klik **✅ Simpan Siswa**
4. Saat monitoring aktif, sistem akan mengenali wajah mahasiswa secara otomatis dan menampilkan kartu identitas beserta **% match accuracy**

> ⚠️ Jika ada foto lama di folder `data/wajah_db/`, hapus file `.pkl` di folder tersebut agar sistem membuat ulang embedding wajah dengan model ArcFace terbaru.

### Ekspor Laporan Excel
Laporan Excel terdiri dari **3 sheet**:

| Sheet | Isi |
|---|---|
| **Ringkasan** | Total deteksi per emosi + persentase + keterangan otomatis |
| **Data Log** | Semua rekaman deteksi dengan warna emosi per baris |
| **Per Mahasiswa** | Breakdown per mahasiswa + kesimpulan otomatis (antusias / fokus / kurang semangat) |

Filter yang tersedia: Hari Ini · 1 Minggu · 1 Bulan · 3 Bulan · **6 Bulan (1 Semester)**

---

## 🛠️ Teknologi yang Digunakan

| Library | Versi | Fungsi |
|---|---|---|
| **Flask** | 3.0.3 | Web framework & MJPEG streaming |
| **OpenCV** | 4.10.0 | Deteksi wajah & pengolahan frame |
| **DeepFace** | 0.0.93 | Klasifikasi emosi & pengenalan wajah (ArcFace) |
| **MediaPipe** | 0.10.14 | Deteksi pose / angkat tangan |
| **TF-Keras** | 2.17.0 | Backend model deep learning |
| **NumPy** | 1.26.4 | Operasi array & preprocessing |
| **OpenPyXL** | 3.1.5 | Ekspor laporan Excel |
| **SQLite** | bawaan Python | Penyimpanan log emosi |

---

## ❗ Troubleshooting

| Masalah | Solusi |
|---|---|
| **Kamera tidak muncul** | Pastikan tidak ada aplikasi lain yang menggunakan kamera |
| **FPS rendah / lambat** | Naikkan nilai `INTERVAL` di `app.py` (misalnya dari 6 ke 10) |
| **Sedih muncul terus padahal netral** | Sudah diperketat (threshold 78%), pastikan versi terbaru |
| **Wajah mahasiswa tidak dikenali** | (1) Pastikan foto jelas & wajah tampak penuh; (2) Hapus `.pkl` di `data/wajah_db/`; (3) Naikkan `DIST_MAX` sedikit di class `FaceRecognizer` |
| **Error saat install** | Coba `pip install tensorflow-cpu` sebelum install requirements |
| **Port sudah dipakai** | Ubah `port=5000` menjadi `port=8080` di bagian bawah `app.py` |
| **MediaPipe tidak terinstall** | Fitur angkat tangan dinonaktifkan otomatis, program tetap berjalan normal |

---

## 📋 Requirements Sistem

- **Python** 3.10 atau lebih baru
- **RAM** minimal 4 GB (disarankan 8 GB)
- **Webcam** (kamera laptop / webcam eksternal)
- **Koneksi internet** (hanya saat install pertama untuk mengunduh model AI)
- **OS**: Windows 10/11 · macOS 12+ · Ubuntu 20.04+

---

## 👨‍💻 Dibuat Untuk

Skripsi / Tugas Akhir — Program Studi Informatika / Teknik Informatika  
Topik: *Computer Vision · Pengenalan Emosi Wajah · Pembelajaran Online*

---

## 📄 Lisensi

Proyek ini dibuat untuk keperluan akademik. Bebas digunakan dan dimodifikasi dengan mencantumkan sumber.
