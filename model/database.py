"""
Database helper — SQLite untuk log emosi & data mahasiswa
"""
import sqlite3, os
from datetime import datetime, timedelta

DB_PATH = 'data/emosi.db'

def get_db():
    con = sqlite3.connect(DB_PATH, check_same_thread=False)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    os.makedirs('data/wajah_mahasiswa', exist_ok=True)
    con = get_db()
    con.executescript('''
        CREATE TABLE IF NOT EXISTS mahasiswa (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nama        TEXT    NOT NULL,
            nim         TEXT    NOT NULL DEFAULT '',
            kelas       TEXT    NOT NULL DEFAULT '',
            foto_path   TEXT    DEFAULT NULL,
            created_at  TEXT    NOT NULL
        );

        CREATE TABLE IF NOT EXISTS log_emosi (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            mahasiswa_id  INTEGER DEFAULT NULL,
            nama          TEXT    NOT NULL DEFAULT 'Tidak Dikenal',
            kelas         TEXT    NOT NULL DEFAULT '-',
            mata_kuliah   TEXT    NOT NULL DEFAULT '-',
            emosi         TEXT    NOT NULL,
            label         TEXT    NOT NULL,
            confidence    REAL    NOT NULL DEFAULT 0,
            angkat_tangan INTEGER NOT NULL DEFAULT 0,
            waktu         TEXT    NOT NULL,
            FOREIGN KEY(mahasiswa_id) REFERENCES mahasiswa(id)
        );

        CREATE INDEX IF NOT EXISTS idx_waktu  ON log_emosi(waktu);
        CREATE INDEX IF NOT EXISTS idx_mapel  ON log_emosi(mata_kuliah);
        CREATE INDEX IF NOT EXISTS idx_mhs_id ON log_emosi(mahasiswa_id);
    ''')
    con.commit()
    con.close()

# ── MAHASISWA ──────────────────────────────
def tambah_mahasiswa(nama, nim, kelas, foto_path=None):
    con = get_db()
    cur = con.execute(
        'INSERT INTO mahasiswa (nama,nim,kelas,foto_path,created_at) VALUES (?,?,?,?,?)',
        (nama, nim, kelas, foto_path, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    )
    mid = cur.lastrowid
    con.commit(); con.close()
    return mid

def daftar_mahasiswa():
    con = get_db()
    rows = con.execute('SELECT * FROM mahasiswa ORDER BY nama').fetchall()
    con.close()
    return [dict(r) for r in rows]

def get_mahasiswa(mid):
    con = get_db()
    r = con.execute('SELECT * FROM mahasiswa WHERE id=?', (mid,)).fetchone()
    con.close()
    return dict(r) if r else None

def hapus_mahasiswa(mid):
    con = get_db()
    con.execute('DELETE FROM mahasiswa WHERE id=?', (mid,))
    con.commit(); con.close()

def update_foto(mid, foto_path):
    con = get_db()
    con.execute('UPDATE mahasiswa SET foto_path=? WHERE id=?', (foto_path, mid))
    con.commit(); con.close()

# ── LOG EMOSI ──────────────────────────────
def simpan_log(mahasiswa_id, nama, kelas, mata_kuliah,
               emosi, label, confidence, angkat_tangan=False):
    waktu = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    con = get_db()
    con.execute(
        '''INSERT INTO log_emosi
           (mahasiswa_id,nama,kelas,mata_kuliah,emosi,label,confidence,angkat_tangan,waktu)
           VALUES (?,?,?,?,?,?,?,?,?)''',
        (mahasiswa_id, nama, kelas, mata_kuliah,
         emosi, label, round(confidence, 2), int(angkat_tangan), waktu)
    )
    con.commit(); con.close()

def query_log(limit=60, mata_kuliah=None, tanggal_mulai=None, tanggal_selesai=None,
              mahasiswa_id=None):
    con = get_db()
    sql, args = 'SELECT * FROM log_emosi WHERE 1=1', []
    if mata_kuliah and mata_kuliah != 'semua':
        sql += ' AND mata_kuliah=?'; args.append(mata_kuliah)
    if tanggal_mulai:
        sql += ' AND waktu>=?'; args.append(tanggal_mulai)
    if tanggal_selesai:
        sql += ' AND waktu<=?'; args.append(tanggal_selesai + ' 23:59:59')
    if mahasiswa_id:
        sql += ' AND mahasiswa_id=?'; args.append(mahasiswa_id)
    sql += ' ORDER BY id DESC'
    if limit:
        sql += f' LIMIT {int(limit)}'
    rows = con.execute(sql, args).fetchall()
    con.close()
    return [dict(r) for r in rows]

def daftar_mata_kuliah():
    con = get_db()
    rows = con.execute(
        "SELECT DISTINCT mata_kuliah FROM log_emosi WHERE mata_kuliah!='-' ORDER BY mata_kuliah"
    ).fetchall()
    con.close()
    return [r[0] for r in rows]

def hapus_log_semua():
    con = get_db()
    con.execute('DELETE FROM log_emosi')
    con.commit(); con.close()
