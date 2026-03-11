"""
=============================================================
  DETEKTOR EMOSI v3
  - DeepFace emotion dengan koreksi bias & threshold
  - MediaPipe Pose → deteksi angkat tangan
  - DeepFace Face Recognition → kenali wajah mahasiswa
=============================================================
"""

import cv2
import numpy as np
import threading
import time
import os
import glob
from collections import deque, Counter
from deepface import DeepFace
import mediapipe as mp

# ─── LABEL EMOSI ───────────────────────────────────────────
EMOSI = {
    'angry':    {'label': 'Marah',    'emoji': '😠', 'warna': '#ef4444'},
    'disgust':  {'label': 'Jijik',    'emoji': '🤢', 'warna': '#8b5cf6'},
    'fear':     {'label': 'Takut',    'emoji': '😨', 'warna': '#6366f1'},
    'happy':    {'label': 'Senang',   'emoji': '😊', 'warna': '#22c55e'},
    'sad':      {'label': 'Sedih',    'emoji': '😢', 'warna': '#3b82f6'},
    'surprise': {'label': 'Terkejut', 'emoji': '😲', 'warna': '#f59e0b'},
    'neutral':  {'label': 'Netral',   'emoji': '😐', 'warna': '#94a3b8'},
}

# ─── THRESHOLD MINIMUM PER EMOSI ──────────────────────────
# Emosi hanya diterima jika confidence-nya MELEBIHI threshold ini.
# Jika tidak, fallback ke NETRAL.
# Sedih sering false-positive pada wajah relaxed → threshold tinggi.
THRESHOLD_EMOSI = {
    'angry':    52,   # Harus yakin > 52%
    'disgust':  62,   # Sangat jarang → perlu keyakinan tinggi
    'fear':     58,   # Jarang di kelas → perlu keyakinan tinggi
    'happy':    38,   # Mudah dikenali → threshold rendah
    'sad':      60,   # FALSE POSITIVE TINGGI → threshold paling ketat
    'surprise': 45,
    'neutral':  20,   # Fallback default → threshold sangat rendah
}

# ─── MEDIAPIPE SETUP ───────────────────────────────────────
mp_pose   = mp.solutions.pose
mp_hands  = mp.solutions.hands
mp_draw   = mp.solutions.drawing_utils


class Detektor:
    def __init__(self):
        self.kamera    = None
        self.aktif     = False
        self.lock      = threading.Lock()
        self.frame_no  = 0
        self.INTERVAL  = 10   # Analisis DeepFace setiap N frame

        # ── Smoothing: majority vote 20 frame terakhir (lebih stabil)
        self.WINDOW  = 20
        self.antrian = deque(maxlen=self.WINDOW)

        # ── Haar Cascade untuk deteksi wajah (ringan)
        cascade = cv2.data.haarcascades + 'haarcascade_frontalface_default.xml'
        self.cascade = cv2.CascadeClassifier(cascade)

        # ── MediaPipe Pose (deteksi angkat tangan)
        self.pose = mp_pose.Pose(
            static_image_mode=False,
            model_complexity=0,          # 0=ringan, 1=normal, 2=akurat
            enable_segmentation=False,
            min_detection_confidence=0.5,
            min_tracking_confidence=0.5,
        )

        # ── State angkat tangan
        self.angkat_tangan       = False
        self.angkat_tangan_hold  = 0
        self.ANGKAT_HOLD         = 45    # Tahan flag selama N frame setelah turun

        # ── Face Recognition: daftar wajah mahasiswa yang terdaftar
        # Format: [{'id':..., 'nama':..., 'embedding':...}, ...]
        self.wajah_terdaftar = []
        self.RECOGNITION_INTERVAL = 30   # Kenali wajah setiap N frame
        self.mahasiswa_terdeteksi = None  # Dict mahasiswa yang sedang terdeteksi

        # ── Hasil terkini
        self.emosi_kini  = 'neutral'
        self.label_kini  = 'Netral'
        self.conf_kini   = 0.0
        self.semua_kini  = {k: 0.0 for k in EMOSI}

        # ── Info sesi (diisi dari UI)
        self.mata_kuliah = '-'
        self.kelas       = '-'
        self.nama_manual = 'Tidak Dikenal'  # Nama fallback jika tidak ada face recog

        # ── Auto-save
        self.last_save     = time.time()
        self.SAVE_INTERVAL = 5

        # ── DB callback (diset dari app.py)
        self.simpan_log_cb = None

    # ────────────────────────────────────────────
    # LOAD WAJAH TERDAFTAR dari folder mahasiswa
    # ────────────────────────────────────────────
    def muat_wajah(self, daftar_mhs: list):
        """
        daftar_mhs: list dict dari DB -> [{'id':1,'nama':'Budi','foto_path':'...'}]
        Ekstrak embedding DeepFace dari setiap foto.
        """
        self.wajah_terdaftar = []
        for mhs in daftar_mhs:
            foto = mhs.get('foto_path')
            if not foto or not os.path.exists(foto):
                continue
            try:
                hasil = DeepFace.represent(
                    img_path=foto,
                    model_name='Facenet',
                    enforce_detection=False,
                    detector_backend='opencv',
                )
                if isinstance(hasil, list):
                    hasil = hasil[0]
                emb = np.array(hasil['embedding'])
                self.wajah_terdaftar.append({
                    'id':        mhs['id'],
                    'nama':      mhs['nama'],
                    'nim':       mhs.get('nim', ''),
                    'kelas':     mhs.get('kelas', ''),
                    'embedding': emb,
                    'foto_path': foto,
                })
                print(f"  ✅ Wajah dimuat: {mhs['nama']}")
            except Exception as e:
                print(f"  ⚠️ Gagal muat wajah {mhs['nama']}: {e}")

        print(f"[Face DB] {len(self.wajah_terdaftar)} wajah terdaftar siap.")

    # ────────────────────────────────────────────
    # KENALI WAJAH di frame
    # ────────────────────────────────────────────
    def _kenali_wajah(self, frame):
        """Bandingkan wajah di frame dengan database. Return dict mahasiswa atau None."""
        if not self.wajah_terdaftar:
            return None
        try:
            hasil = DeepFace.represent(
                img_path=frame,
                model_name='Facenet',
                enforce_detection=False,
                detector_backend='opencv',
            )
            if isinstance(hasil, list):
                hasil = hasil[0]
            emb_kini = np.array(hasil['embedding'])

            best_dist = float('inf')
            best_mhs  = None
            THRESHOLD_DIST = 0.4   # Threshold jarak Euclidean (semakin kecil=semakin mirip)

            for mhs in self.wajah_terdaftar:
                emb_ref = mhs['embedding']
                # Cosine distance
                dot     = np.dot(emb_kini, emb_ref)
                norm    = np.linalg.norm(emb_kini) * np.linalg.norm(emb_ref)
                dist    = 1 - (dot / (norm + 1e-6))
                if dist < best_dist:
                    best_dist = dist
                    best_mhs  = mhs

            if best_dist <= THRESHOLD_DIST:
                return best_mhs
            return None
        except:
            return None

    # ────────────────────────────────────────────
    # KOREKSI BIAS EMOSI — kunci akurasi lebih baik
    # ────────────────────────────────────────────
    def _koreksi_emosi(self, semua: dict):
        """
        Terapkan threshold minimum per emosi.
        Jika emosi dominan tidak melewati threshold → fallback ke netral.
        Menghilangkan false positive 'sedih' pada ekspresi relaxed/netral.
        """
        if not semua:
            return 'neutral', 0.0

        # Cari emosi dengan nilai tertinggi
        dominan = max(semua, key=semua.get)
        conf    = semua.get(dominan, 0)

        threshold = THRESHOLD_EMOSI.get(dominan, 45)

        if conf < threshold:
            # Tidak yakin cukup → anggap netral
            return 'neutral', semua.get('neutral', conf)

        return dominan, conf

    # ────────────────────────────────────────────
    # ANALISIS EMOSI via DeepFace
    # ────────────────────────────────────────────
    def _analisis_emosi(self, frame):
        try:
            hasil = DeepFace.analyze(
                img_path=frame,
                actions=['emotion'],
                enforce_detection=False,
                detector_backend='opencv',
                silent=True,
            )
            if isinstance(hasil, list):
                hasil = hasil[0]
            semua = hasil.get('emotion', {k: 0.0 for k in EMOSI})
            return semua
        except:
            return {k: 0.0 for k in EMOSI}

    # ────────────────────────────────────────────
    # SMOOTHING — majority vote berboboti
    # ────────────────────────────────────────────
    def _smooth(self, emosi_baru: str) -> str:
        self.antrian.append(emosi_baru)
        # Bobot lebih tinggi untuk entri yang lebih baru
        n = len(self.antrian)
        bobot_total = {}
        for i, e in enumerate(self.antrian):
            bobot = (i + 1) / n    # entry terbaru = bobot tertinggi
            bobot_total[e] = bobot_total.get(e, 0) + bobot
        return max(bobot_total, key=bobot_total.get)

    # ────────────────────────────────────────────
    # DETEKSI ANGKAT TANGAN via MediaPipe Pose
    # ────────────────────────────────────────────
    def _cek_angkat_tangan(self, frame_rgb):
        """
        Kondisi angkat tangan:
        - Pergelangan tangan (WRIST) berada DI ATAS bahu (SHOULDER) dalam frame
        - Dalam koordinat gambar: y lebih kecil = posisi lebih atas
        """
        try:
            hasil = self.pose.process(frame_rgb)
            if not hasil.pose_landmarks:
                if self.angkat_tangan_hold > 0:
                    self.angkat_tangan_hold -= 1
                self.angkat_tangan = (self.angkat_tangan_hold > 0)
                return

            lm = hasil.pose_landmarks.landmark

            # Landmark kiri: SHOULDER=11, WRIST=15
            # Landmark kanan: SHOULDER=12, WRIST=16
            L_shoulder = lm[mp_pose.PoseLandmark.LEFT_SHOULDER]
            L_wrist    = lm[mp_pose.PoseLandmark.LEFT_WRIST]
            R_shoulder = lm[mp_pose.PoseLandmark.RIGHT_SHOULDER]
            R_wrist    = lm[mp_pose.PoseLandmark.RIGHT_WRIST]

            MARGIN = 0.05  # Wrist harus minimal 5% di atas bahu

            angkat_kiri  = (L_wrist.y < L_shoulder.y - MARGIN and
                            L_wrist.visibility > 0.5 and L_shoulder.visibility > 0.5)
            angkat_kanan = (R_wrist.y < R_shoulder.y - MARGIN and
                            R_wrist.visibility > 0.5 and R_shoulder.visibility > 0.5)

            if angkat_kiri or angkat_kanan:
                self.angkat_tangan_hold = self.ANGKAT_HOLD
            elif self.angkat_tangan_hold > 0:
                self.angkat_tangan_hold -= 1

            self.angkat_tangan = (self.angkat_tangan_hold > 0)
        except:
            self.angkat_tangan = False

    # ────────────────────────────────────────────
    # GAMBAR ANOTASI DI FRAME
    # ────────────────────────────────────────────
    def _anotasi(self, frame, wajah_list, emosi, conf, mhs):
        info = EMOSI.get(emosi, EMOSI['neutral'])
        hex_w = info['warna'].lstrip('#')
        r, g, b = tuple(int(hex_w[i:i+2], 16) for i in (0, 2, 4))
        bgr = (b, g, r)

        nama_tampil = mhs['nama'] if mhs else self.nama_manual

        for (x, y, w, h) in wajah_list:
            cv2.rectangle(frame, (x, y), (x+w, y+h), bgr, 2)

            if self.angkat_tangan:
                teks = f"Angkat Tangan! ({conf:.0f}%)"
                t_color = (0, 220, 80)
            else:
                teks = f"{info['label']} ({conf:.0f}%)"
                t_color = (255, 255, 255)

            (tw, th), _ = cv2.getTextSize(teks, cv2.FONT_HERSHEY_SIMPLEX, 0.60, 2)
            cv2.rectangle(frame, (x, y-th-30), (x+max(tw,len(nama_tampil)*9)+8, y), bgr, -1)
            cv2.putText(frame, teks,         (x+4, y-18),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.60, t_color, 2)
            cv2.putText(frame, nama_tampil,  (x+4, y-4),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.45, (220, 220, 220), 1)

        h_fr = frame.shape[0]

        # Baris bawah kiri
        cv2.putText(frame, f"Emosi: {info['label']}",
                    (10, h_fr-36), cv2.FONT_HERSHEY_SIMPLEX, 0.60, bgr, 2)
        cv2.putText(frame, f"{self.mata_kuliah} | {self.kelas}",
                    (10, h_fr-12), cv2.FONT_HERSHEY_SIMPLEX, 0.48, (180,180,180), 1)

        # Banner angkat tangan
        if self.angkat_tangan:
            cv2.rectangle(frame, (0, 0), (frame.shape[1], 38), (0, 100, 30), -1)
            cv2.putText(frame, "✋ ANGKAT TANGAN — Mahasiswa Ingin Bertanya",
                        (10, 26), cv2.FONT_HERSHEY_SIMPLEX, 0.65, (0, 255, 80), 2)

        return frame

    # ────────────────────────────────────────────
    # BUKA / TUTUP KAMERA
    # ────────────────────────────────────────────
    def buka(self, idx=0):
        if self.kamera and self.kamera.isOpened():
            return True
        self.kamera = cv2.VideoCapture(idx)
        if self.kamera.isOpened():
            self.kamera.set(cv2.CAP_PROP_FRAME_WIDTH,  640)
            self.kamera.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
            self.aktif = True
            return True
        return False

    def tutup(self):
        self.aktif = False
        if self.kamera:
            self.kamera.release()
            self.kamera = None

    # ────────────────────────────────────────────
    # GENERATOR MJPEG (main loop)
    # ────────────────────────────────────────────
    def generate(self):
        if not self.buka():
            return

        semua_emosi   = {k: 0.0 for k in EMOSI}
        emosi_raw     = 'neutral'

        while self.aktif:
            ok, frame = self.kamera.read()
            if not ok:
                break

            self.frame_no += 1
            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

            # ── Deteksi wajah (Haar, tiap frame)
            gray  = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            wajah = self.cascade.detectMultiScale(gray, 1.1, 5, minSize=(60,60))

            # ── Analisis emosi setiap INTERVAL frame
            if self.frame_no % self.INTERVAL == 0:
                semua_emosi = self._analisis_emosi(frame)
                emosi_raw   = max(semua_emosi, key=semua_emosi.get)

            # ── Koreksi bias + threshold
            emosi_c, conf_c = self._koreksi_emosi(semua_emosi)

            # ── Smoothing berboboti
            emosi_s = self._smooth(emosi_c)
            conf_s  = semua_emosi.get(emosi_s, conf_c)

            # ── Cek angkat tangan (setiap frame, ringan)
            self._cek_angkat_tangan(frame_rgb)

            # ── Kenali wajah setiap RECOGNITION_INTERVAL frame
            if self.frame_no % self.RECOGNITION_INTERVAL == 0:
                self.mahasiswa_terdeteksi = self._kenali_wajah(frame)

            mhs = self.mahasiswa_terdeteksi

            # ── Update state
            with self.lock:
                self.emosi_kini  = emosi_s
                self.label_kini  = EMOSI.get(emosi_s, EMOSI['neutral'])['label']
                self.conf_kini   = round(conf_s, 2)
                self.semua_kini  = semua_emosi

            # ── Auto-save ke DB setiap SAVE_INTERVAL detik
            now = time.time()
            if now - self.last_save >= self.SAVE_INTERVAL and self.simpan_log_cb:
                mid   = mhs['id']   if mhs else None
                nama  = mhs['nama'] if mhs else self.nama_manual
                kelas = mhs['kelas'] if mhs else self.kelas
                self.simpan_log_cb(
                    mid, nama, kelas, self.mata_kuliah,
                    emosi_s, EMOSI[emosi_s]['label'],
                    conf_s, self.angkat_tangan
                )
                self.last_save = now

            # ── Render anotasi
            frame = self._anotasi(frame.copy(), wajah, emosi_s, conf_s, mhs)
            _, buf = cv2.imencode('.jpg', frame, [cv2.IMWRITE_JPEG_QUALITY, 85])
            yield (b'--frame\r\nContent-Type: image/jpeg\r\n\r\n'
                   + buf.tobytes() + b'\r\n')

        self.tutup()

    # ────────────────────────────────────────────
    # STATUS JSON
    # ────────────────────────────────────────────
    def get_status(self):
        with self.lock:
            mhs = self.mahasiswa_terdeteksi
            return {
                'emosi':          self.emosi_kini,
                'label':          self.label_kini,
                'emoji':          EMOSI.get(self.emosi_kini, EMOSI['neutral'])['emoji'],
                'conf':           round(self.conf_kini, 2),
                'angkat_tangan':  self.angkat_tangan,
                'mahasiswa': {
                    'id':   mhs['id']    if mhs else None,
                    'nama': mhs['nama']  if mhs else self.nama_manual,
                    'nim':  mhs.get('nim','') if mhs else '',
                } if True else None,
                'semua': {
                    k: {
                        'label': EMOSI[k]['label'],
                        'emoji': EMOSI[k]['emoji'],
                        'nilai': round(self.semua_kini.get(k, 0), 2),
                        'warna': EMOSI[k]['warna'],
                    } for k in EMOSI
                }
            }


# Instance global
cam = Detektor()
