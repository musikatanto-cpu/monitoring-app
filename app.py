import streamlit as st
import datetime
import sqlite3
import pandas as pd
import hashlib
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# ============================================================
# KONFIGURASI APLIKASI
# ============================================================
st.set_page_config(
    page_title="Sistem Monitoring Binpres",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded"
)

DB_NAME = "monitoring.db"


# ============================================================
# DATABASE
# ============================================================
def get_conn():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def hash_password(password):
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def verify_password(password, password_hash):
    return hash_password(password) == password_hash


def init_db():
    conn = get_conn()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user',
            nama_lengkap TEXT DEFAULT '',
            aktif INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cabor (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT UNIQUE NOT NULL,
            aktif INTEGER DEFAULT 1
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS laporan_monitoring (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal DATE,
            cabor TEXT,
            lokasi TEXT,
            petugas TEXT,
            fisik_parameter TEXT,
            fisik_peaking TEXT,
            fisik_recovery TEXT,
            fisik_cedera TEXT,
            taktis_lawan TEXT,
            taktis_instruksi TEXT,
            taktis_ujicoba TEXT,
            mental_cemas TEXT,
            mental_fokus TEXT,
            mental_rutinitas TEXT,
            mental_psikolog TEXT,
            nutrisi_bb TEXT,
            nutrisi_asupan TEXT,
            nutrisi_hidrasi TEXT,
            nutrisi_tidur TEXT,
            medis_rekam TEXT,
            medis_doping TEXT,
            medis_alat TEXT,
            medis_nonteknis TEXT,
            status TEXT DEFAULT 'Belum Ditindaklanjuti',
            catatan_admin TEXT DEFAULT '',
            foto_data BLOB,
            koordinat_gps TEXT DEFAULT '',
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Migrasi database jika ada kolom baru
    existing_reports = [r["name"] for r in cursor.execute("PRAGMA table_info(laporan_monitoring)").fetchall()]
    report_columns = {
        "status": "TEXT DEFAULT 'Belum Ditindaklanjuti'",
        "catatan_admin": "TEXT DEFAULT ''",
        "foto_data": "BLOB",
        "koordinat_gps": "TEXT DEFAULT ''",
        "updated_at": "TEXT DEFAULT CURRENT_TIMESTAMP"
    }
    for col, definition in report_columns.items():
        if col not in existing_reports:
            cursor.execute(f"ALTER TABLE laporan_monitoring ADD COLUMN {col} {definition}")

    # Account default
    count = cursor.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]
    if count == 0:
        cursor.execute(
            "INSERT INTO users (username, password, role, nama_lengkap, aktif) VALUES (?, ?, ?, ?, 1)",
            ("admin", hash_password("admin123"), "admin", "Administrator")
        )
        cursor.execute(
            "INSERT INTO users (username, password, role, nama_lengkap, aktif) VALUES (?, ?, ?, ?, 1)",
            ("petugas", hash_password("petugas123"), "user", "Petugas Monitoring")
        )

    default_cabor = [
        'ANGGAR', 'ANGKAT BERAT', 'ANGKAT BESI', 'AQUATIC/RENANG', 'ARUNG JERAM',
        'ATLETIK', 'BALAP SEPEDA', 'BARONGSAI', 'BERMOTOR', 'BILLIARD', 'BINARAGA',
        'BOLA BASKET', 'BOLA TANGAN', 'BOLA VOLI', 'BOWLING', 'BRIDGE', 'BULUTANGKIS',
        'CATUR', 'DAYUNG', 'DRUMBAND', 'E-SPORT', 'FLOOR BALL', 'FUTSAL', 'GATEBALL',
        'GOLF', 'GULAT', 'GYMNASTIC/SENAM', 'HOKI', 'IBCA MMA', 'JU JITSU', 'JUDO',
        'KARATE', 'KEMPO', 'MENEMBAK', 'MUAYTHAI', 'PANAHAN', 'PANJAT TEBING',
        'PENCAK SILAT', 'PETANQUE', 'PICKLEBALL', 'RUGBY', 'SAMBO', 'SELAM',
        'SEPAK BOLA', 'SEPAK TAKRAW', 'SEPATU RODA', 'SOFTBALL', 'SQUASH',
        'TAEKWONDO', 'TARUNG DERAJAT', 'TENIS LAPANG', 'TENIS MEJA', 'TINJU',
        'WOODBALL', 'WUSHU'
    ]
    if cursor.execute("SELECT COUNT(*) AS n FROM cabor").fetchone()["n"] == 0:
        cursor.executemany("INSERT OR IGNORE INTO cabor (nama) VALUES (?)", [(x,) for x in default_cabor])

    conn.commit()
    conn.close()


# ============================================================
# HELPER WATERMARK FOTO (TIME MARK & REL LOKASI)
# ============================================================
def beri_watermark_foto(image_bytes, timestamp_str, lokasi_str, gps_str=""):
    """Menambahkan timestamp dan lokasi real secara otomatis ke atas gambar"""
    img = Image.open(BytesIO(image_bytes)).convert("RGB")
    
    # Orientasi & Ukuran
    w, h = img.size
    
    # Buat overlay semi transparan di bagian bawah
    overlay = Image.new("RGBA", img.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)
    
    box_height = int(h * 0.14)
    if box_height < 70:
        box_height = 70
        
    # Rectangle hitam transparan di bawah foto
    draw.rectangle([(0, h - box_height), (w, h)], fill=(0, 0, 0, 180))
    
    # Gabungkan overlay ke gambar utama
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    draw = ImageDraw.Draw(img)
    
    # Tentukan teks watermark
    teks_waktu = f"⏰ Waktu: {timestamp_str}"
    teks_lokasi = f"📍 Lokasi: {lokasi_str}"
    if gps_str:
        teks_lokasi += f" ({gps_str})"
    teks_watermark = f"{teks_waktu}\n{teks_lokasi}\n🛡️ VERIFIED MONEV BINPRES KONI KAB. TANGERANG"
    
    # Font size responsif
    font_size = max(14, int(h * 0.022))
    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except IOError:
        font = ImageFont.load_default()
        
    # Cetak teks warna putih
    padding_x = int(w * 0.03)
    padding_y = h - box_height + 10
    draw.text((padding_x, padding_y), teks_watermark, fill=(255, 255, 0), font=font)
    
    out_io = BytesIO()
    img.save(out_io, format="JPEG", quality=92)
    return out_io.getvalue()


def logout():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()


def get_cabor_list():
    conn = get_conn()
    rows = conn.execute("SELECT nama FROM cabor WHERE aktif=1 ORDER BY nama").fetchall()
    conn.close()
    return [r["nama"] for r in rows]


def get_current_user():
    if "username" not in st.session_state:
        return None
    conn = get_conn()
    row = conn.execute("SELECT * FROM users WHERE username=?", (st.session_state["username"],)).fetchone()
    conn.close()
    return row


# ============================================================
# GENERATE WORD REPORT (RAPI, KORPORAT & PROFESIONAL)
# ============================================================
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def generate_word_report(data_list, is_all=False):
    doc = Document()
    
    # Set Margin Standar Laporan
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    PRIMARY_COLOR = RGBColor(0, 51, 102) # Navy Blue
    
    for idx, data in enumerate(data_list):
        # Header Laporan
        title_p = doc.add_paragraph()
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_title = title_p.add_run("LAPORAN MONITORING & EVALUASI (MONEV)\nBIDANG PEMBINAAN DAN PRESTASI")
        run_title.font.name = 'Arial'
        run_title.font.size = Pt(14)
        run_title.font.bold = True
        run_title.font.color.rgb = PRIMARY_COLOR
        
        sub_title = doc.add_paragraph()
        sub_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_sub = sub_title.add_run("KONI KABUPATEN TANGERANG")
        run_sub.font.name = 'Arial'
        run_sub.font.size = Pt(12)
        run_sub.font.bold = True
        
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

        # Tabel Informasi Utama
        table_info = doc.add_table(rows=5, cols=2)
        table_info.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_info.autofit = False
        
        info_items = [
            ("Cabang Olahraga", data.get("cabor", "-")),
            ("Tanggal Monitoring", str(data.get("tanggal", "-"))),
            ("Lokasi Latihan / Event", data.get("lokasi", "-")),
            ("Petugas Monev", data.get("petugas", "-")),
            ("Status Tindak Lanjut", data.get("status", "Belum Ditindaklanjuti"))
        ]
        
        for row_idx, (label, val) in enumerate(info_items):
            cell_lbl = table_info.cell(row_idx, 0)
            cell_val = table_info.cell(row_idx, 1)
            
            cell_lbl.width = Inches(2.2)
            cell_val.width = Inches(4.6)
            
            p0 = cell_lbl.paragraphs[0]
            r0 = p0.add_run(label)
            r0.bold = True
            r0.font.size = Pt(10)
            
            p1 = cell_val.paragraphs[0]
            r1 = p1.add_run(val)
            r1.font.size = Pt(10)
            
            set_cell_background(cell_lbl, "F2F4F7")

        doc.add_paragraph().paragraph_format.space_after = Pt(12)

        # Fungsi Helper Tambah Seksi Indikator
        def add_section_table(section_title, items):
            h_p = doc.add_paragraph()
            h_run = h_p.add_run(section_title)
            h_run.font.name = 'Arial'
            h_run.font.size = Pt(11)
            h_run.font.bold = True
            h_run.font.color.rgb = PRIMARY_COLOR
            h_p.paragraph_format.space_before = Pt(8)
            h_p.paragraph_format.space_after = Pt(4)

            t = doc.add_table(rows=len(items) + 1, cols=2)
            t.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            # Header Table
            hdr_cells = t.rows[0].cells
            hdr_cells[0].width = Inches(2.8)
            hdr_cells[1].width = Inches(4.0)
            
            hdr_cells[0].paragraphs[0].add_run("Indikator Evaluasi").bold = True
            hdr_cells[1].paragraphs[0].add_run("Catatan / Temuan Lapangan").bold = True
            
            set_cell_background(hdr_cells[0], "1F4E78")
            set_cell_background(hdr_cells[1], "1F4E78")
            hdr_cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
            hdr_cells[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

            for idx_i, (q_text, val_text) in enumerate(items):
                row_cells = t.rows[idx_i + 1].cells
                row_cells[0].width = Inches(2.8)
                row_cells[1].width = Inches(4.0)
                
                p_q = row_cells[0].paragraphs[0]
                rq = p_q.add_run(q_text)
                rq.font.size = Pt(9.5)
                
                p_v = row_cells[1].paragraphs[0]
                rv = p_v.add_run(val_text if val_text else "-")
                rv.font.size = Pt(9.5)
                
                if idx_i % 2 == 1:
                    set_cell_background(row_cells[0], "F9FAFB")
                    set_cell_background(row_cells[1], "F9FAFB")
            
            doc.add_paragraph().paragraph_format.space_after = Pt(6)

        # 1. Fisik
        add_section_table("1. Performa Fisik & Kebugaran", [
            ("Capaian Parameter Fisik vs Target", data.get("fisik_parameter")),
            ("Grafik Performa Puncak (Peaking)", data.get("fisik_peaking")),
            ("Tingkat Pemulihan (Recovery)", data.get("fisik_recovery")),
            ("Keluhan / Indikasi Cedera", data.get("fisik_cedera"))
        ])

        # 2. Taktis
        add_section_table("2. Kesiapan Taktis & Strategi", [
            ("Pemetaan Kekuatan Lawan", data.get("taktis_lawan")),
            ("Instruksi Teknis Bertekanan", data.get("taktis_instruksi")),
            ("Hasil Try-out / Sparing", data.get("taktis_ujicoba"))
        ])

        # 3. Mental
        add_section_table("3. Mental & Psikologis Atlet", [
            ("Pengendalian Stres & Kecemasan", data.get("mental_cemas")),
            ("Fokus, Motivasi & Self-Confidence", data.get("mental_fokus")),
            ("Rutinitas Mental Khusus", data.get("mental_rutinitas")),
            ("Koordinasi Psikolog Olahraga", data.get("mental_psikolog"))
        ])

        # 4. Nutrisi
        add_section_table("4. Nutrisi, Berat Badan & Gaya Hidup", [
            ("Progres Penyesuaian Berat Badan", data.get("nutrisi_bb")),
            ("Asupan Nutrisi & Suplemen", data.get("nutrisi_asupan")),
            ("Status Hidrasi Harian", data.get("nutrisi_hidrasi")),
            ("Kualitas & Cukup Tidur", data.get("nutrisi_tidur"))
        ])

        # 5. Medis & Logistik
        add_section_table("5. Medis, Bebas Doping & Logistik", [
            ("Rekam Medis & Fisioterapis", data.get("medis_rekam")),
            ("Keamanan Obat & Suplemen (Doping)", data.get("medis_doping")),
            ("Kesiapan Alat / Kostum Tanding", data.get("medis_alat")),
            ("Kendala Non-Teknis", data.get("medis_nonteknis"))
        ])

        # Catatan Admin jika ada
        if data.get("catatan_admin"):
            add_section_table("Catatan & Tindak Lanjut Admin", [
                ("Rekomendasi Admin BINPRES", data.get("catatan_admin"))
            ])

        # FOTO DOKUMENTASI TER-WATERMARK
        if data.get("foto_data"):
            doc.add_page_break()
            f_title = doc.add_paragraph()
            f_run = f_title.add_run("DOKUMENTASI FOTO MONITORING (TIME MARK & REL LOKASI)")
            f_run.font.name = 'Arial'
            f_run.font.size = Pt(11)
            f_run.font.bold = True
            f_run.font.color.rgb = PRIMARY_COLOR
            f_title.paragraph_format.space_after = Pt(8)

            img_stream = BytesIO(data["foto_data"])
            doc.add_picture(img_stream, width=Inches(5.8))
            
            cap_p = doc.add_paragraph()
            cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            c_run = cap_p.add_run(f"Gambar 1: Dokumentasi Lapangan Cabor {data.get('cabor')} - {data.get('tanggal')}")
            c_run.font.size = Pt(8.5)
            c_run.font.italic = True

        if idx < len(data_list) - 1 and not data.get("foto_data"):
            doc.add_page_break()

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ============================================================
# LOGIN
# ============================================================
def halaman_login():
    st.markdown(
        """
        <div style="text-align:center;padding:30px 0 15px 0">
            <div style="font-size:52px">🏆</div>
            <h1 style="margin-bottom:5px">Sistem Monitoring BINPRES</h1>
            <p style="color:#777">KONI Kabupaten Tangerang</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    with st.form("login_form"):
        username = st.text_input("👤 Username")
        password = st.text_input("🔒 Password", type="password")
        submit = st.form_submit_button("Masuk ke Sistem", use_container_width=True, type="primary")

        if submit:
            conn = get_conn()
            user = conn.execute("SELECT * FROM users WHERE username=? AND aktif=1", (username.strip(),)).fetchone()
            conn.close()

            valid = False
            if user:
                valid = verify_password(password, user["password"]) or password == user["password"]

            if valid:
                st.session_state["logged_in"] = True
                st.session_state["username"] = user["username"]
                st.session_state["role"] = user["role"]
                st.session_state["nama_lengkap"] = user["nama_lengkap"] or user["username"]
                st.rerun()
            else:
                st.error("Username atau Password salah / akun tidak aktif.")


def sidebar_app():
    user = get_current_user()
    with st.sidebar:
        st.markdown("## 🏆 BINPRES")
        st.caption("Monitoring & Evaluasi Cabor")
        st.markdown("---")
        st.write("👤 **" + (user["nama_lengkap"] if user else st.session_state.get("username", "")) + "**")
        st.caption("Role: " + st.session_state.get("role", "").upper())
        st.markdown("---")

        if st.button("🚪 Logout", use_container_width=True):
            logout()


# ============================================================
# DASHBOARD ADMIN & KELOLA
# ============================================================
def dashboard_admin():
    conn = get_conn()
    total = conn.execute("SELECT COUNT(*) AS n FROM laporan_monitoring").fetchone()["n"]
    bulan_ini = conn.execute("SELECT COUNT(*) AS n FROM laporan_monitoring WHERE strftime('%Y-%m', tanggal)=strftime('%Y-%m','now')").fetchone()["n"]
    cabor_termonitor = conn.execute("SELECT COUNT(DISTINCT cabor) AS n FROM laporan_monitoring").fetchone()["n"]
    tindak = conn.execute("SELECT COUNT(*) AS n FROM laporan_monitoring WHERE status='Belum Ditindaklanjuti' OR status IS NULL").fetchone()["n"]
    conn.close()

    st.markdown("## 📊 Dashboard Monitoring")
    a, b, c, d = st.columns(4)
    a.metric("Total Laporan", total)
    b.metric("Laporan Bulan Ini", bulan_ini)
    c.metric("Cabor Termonitor", cabor_termonitor)
    d.metric("Perlu Tindak Lanjut", tindak)
    st.markdown("---")


def kelola_user():
    st.subheader("👥 Manajemen User")
    conn = get_conn()
    df = pd.read_sql_query("SELECT id, username, nama_lengkap, role, CASE WHEN aktif=1 THEN 'Aktif' ELSE 'Nonaktif' END AS status, created_at FROM users ORDER BY id DESC", conn)
    conn.close()
    st.dataframe(df, use_container_width=True, hide_index=True)

    with st.form("tambah_user"):
        c1, c2 = st.columns(2)
        with c1:
            username = st.text_input("Username baru")
            nama = st.text_input("Nama lengkap")
        with c2:
            password = st.text_input("Password", type="password")
            role = st.selectbox("Hak akses", ["user", "admin"])

        submit = st.form_submit_button("Simpan User", use_container_width=True, type="primary")

        if submit:
            if not username.strip() or not password:
                st.error("Username dan password wajib diisi.")
            else:
                conn = get_conn()
                try:
                    conn.execute("INSERT INTO users (username,password,role,nama_lengkap,aktif) VALUES (?,?,?,?,1)",
                                 (username.strip(), hash_password(password), role, nama.strip()))
                    conn.commit()
                    st.success("User berhasil ditambahkan.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Username sudah digunakan.")
                finally:
                    conn.close()


def kelola_cabor():
    st.subheader("🏅 Kelola Cabang Olahraga")
    conn = get_conn()
    df = pd.read_sql_query("SELECT id, nama, CASE WHEN aktif=1 THEN 'Aktif' ELSE 'Nonaktif' END AS status FROM cabor ORDER BY nama", conn)
    conn.close()
    st.dataframe(df, use_container_width=True, hide_index=True)

    with st.form("tambah_cabor"):
        nama_cabor = st.text_input("Nama Cabang Olahraga Baru")
        submit = st.form_submit_button("➕ Tambah Cabor", use_container_width=True)

        if submit:
            if nama_cabor.strip():
                conn = get_conn()
                try:
                    conn.execute("INSERT INTO cabor (nama) VALUES (?)", (nama_cabor.strip().upper(),))
                    conn.commit()
                    st.success("Cabor berhasil ditambahkan.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Cabor tersebut sudah ada.")
                finally:
                    conn.close()


def halaman_laporan_admin():
    st.subheader("📄 Data Laporan Monitoring")

    conn = get_conn()
    df = pd.read_sql_query("SELECT id, tanggal, cabor, lokasi, petugas, status FROM laporan_monitoring ORDER BY tanggal DESC, id DESC", conn)
    conn.close()

    if df.empty:
        st.info("Belum ada laporan masuk.")
        return

    f1, f2, f3 = st.columns(3)
    with f1:
        cabor_filter = st.selectbox("Filter Cabor", ["Semua"] + sorted(df["cabor"].dropna().unique().tolist()))
    with f2:
        status_filter = st.selectbox("Filter Status", ["Semua"] + sorted(df["status"].fillna("Belum Ditindaklanjuti").unique().tolist()))
    with f3:
        keyword = st.text_input("🔎 Cari lokasi / petugas")

    filtered = df.copy()
    if cabor_filter != "Semua":
        filtered = filtered[filtered["cabor"] == cabor_filter]
    if status_filter != "Semua":
        filtered = filtered[filtered["status"].fillna("Belum Ditindaklanjuti") == status_filter]
    if keyword:
        mask = filtered["lokasi"].fillna("").str.contains(keyword, case=False) | filtered["petugas"].fillna("").str.contains(keyword, case=False)
        filtered = filtered[mask]

    st.dataframe(filtered, use_container_width=True, hide_index=True)

    if filtered.empty:
        return

    st.markdown("---")
    st.markdown("### 📥 Export & Tindak Lanjut")

    pilihan_id = st.selectbox("Pilih ID laporan untuk diunduh / ditindaklanjuti", filtered["id"].tolist())

    conn = get_conn()
    row = conn.execute("SELECT * FROM laporan_monitoring WHERE id=?", (int(pilihan_id),)).fetchone()
    conn.close()

    if row:
        c1, c2 = st.columns(2)
        with c1:
            word_file = generate_word_report([dict(row)], is_all=False)
            st.download_button(
                "📄 Download Laporan Word (Rapi & Profesional)",
                data=word_file,
                file_name=f"Laporan_{row['cabor']}_{row['tanggal']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )

        with c2:
            with st.form("update_status_laporan"):
                status = st.selectbox("Status", ["Belum Ditindaklanjuti", "Sedang Ditindaklanjuti", "Selesai"], index=["Belum Ditindaklanjuti", "Sedang Ditindaklanjuti", "Selesai"].index(row["status"] or "Belum Ditindaklanjuti"))
                catatan = st.text_area("Catatan Admin", value=row["catatan_admin"] or "")
                simpan = st.form_submit_button("💾 Simpan Status", use_container_width=True)

                if simpan:
                    conn = get_conn()
                    conn.execute("UPDATE laporan_monitoring SET status=?, catatan_admin=?, updated_at=CURRENT_TIMESTAMP WHERE id=?", (status, catatan, int(pilihan_id)))
                    conn.commit()
                    conn.close()
                    st.success("Tindak lanjut diperbarui.")
                    st.rerun()


def halaman_admin():
    sidebar_app()
    dashboard_admin()

    tab1, tab2, tab3 = st.tabs(["📄 Laporan Monitoring", "👥 Manajemen User", "🏅 Kelola Cabor"])
    with tab1:
        halaman_laporan_admin()
    with tab2:
        kelola_user()
    with tab3:
        kelola_cabor()


# ============================================================
# HALAMAN USER / PETUGAS (INPUT & UNDUH)
# ============================================================
def halaman_user():
    sidebar_app()

    st.markdown("## 📝 Form Input Monitoring Lapangan")
    st.caption(f"Petugas: **{st.session_state.get('nama_lengkap', st.session_state['username'])}**")

    cabor_list = get_cabor_list()

    with st.form("form_monitoring", clear_on_submit=False):
        st.markdown("### 📌 Informasi Dasar & Lokasi Real")
        col1, col2 = st.columns(2)

        with col1:
            tanggal = st.date_input("Tanggal Monitoring", datetime.date.today())
            cabor = st.selectbox("Cabang Olahraga", ["Pilih Cabor..."] + cabor_list)

        with col2:
            lokasi = st.text_input("Lokasi Latihan / Stadion / Gelanggang (Real)", placeholder="Contoh: GOR Stadion Benteng Reborn")
            gps_input = st.text_input("Koordinat GPS / Link Google Maps (Opsional)", placeholder="-6.175392, 106.638210")

        st.markdown("---")
        st.markdown("### 📷 Upload Foto Dokumentasi (Wajib Time Mark & Lokasi)")
        foto_file = st.file_uploader("Unggah Foto Kegiatan Monev (JPG/PNG)", type=["jpg", "jpeg", "png"])

        st.markdown("---")

        with st.expander("💪 1. Performa Fisik & Kebugaran", expanded=True):
            fisik_1 = st.text_area("Capaian parameter fisik (vs target):", height=68)
            fisik_2 = st.text_area("Grafik performa puncak (peaking):", height=68)
            fisik_3 = st.text_area("Tingkat pemulihan fisik (recovery):", height=68)
            fisik_4 = st.text_area("Keluhan cedera lama / baru:", height=68)

        with st.expander("🎯 2. Kesiapan Taktis & Strategi"):
            taktis_1 = st.text_area("Pemetaan kekuatan lawan:", height=68)
            taktis_2 = st.text_area("Instruksi teknis di bawah tekanan:", height=68)
            taktis_3 = st.text_area("Hasil try-out / sparing:", height=68)

        with st.expander("🧠 3. Mental & Psikologis"):
            mental_1 = st.text_area("Pengendalian kecemasan & stres:", height=68)
            mental_2 = st.text_area("Fokus & self-confidence:", height=68)
            mental_3 = st.text_area("Rutinitas mental khusus:", height=68)
            mental_4 = st.text_area("Koordinasi dengan psikolog:", height=68)

        with st.expander("🥗 4. Nutrisi & Gaya Hidup"):
            nutrisi_1 = st.text_area("Penyesuaian berat badan:", height=68)
            nutrisi_2 = st.text_area("Asupan nutrisi & suplemen:", height=68)
            nutrisi_3 = st.text_area("Status hidrasi harian:", height=68)
            nutrisi_4 = st.text_area("Kualitas & waktu tidur:", height=68)

        with st.expander("⚕️ 5. Medis, Bebas Doping & Logistik"):
            medis_1 = st.text_area("Status rekam medis & fisioterapis:", height=68)
            medis_2 = st.text_area("Keamanan obat/suplemen (Bebas Doping):", height=68)
            medis_3 = st.text_area("Kesiapan alat tanding:", height=68)
            medis_4 = st.text_area("Kendala non-teknis:", height=68)

        st.markdown("---")
        submitted = st.form_submit_button("💾 Simpan Laporan & Buat Dokumen Word", use_container_width=True, type="primary")

        if submitted:
            if cabor == "Pilih Cabor...":
                st.error("⚠️ Harap pilih Cabang Olahraga.")
            elif not lokasi.strip():
                st.error("⚠️ Lokasi wajib diisi.")
            elif not foto_file:
                st.error("⚠️ Foto Dokumentasi wajib diunggah untuk mencetak Time Mark & Real Lokasi.")
            else:
                # Waktu real time mark saat diunggah
                now_str = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S WIB")
                
                # Buat Watermark pada Foto
                raw_bytes = foto_file.read()
                watermarked_photo = beri_watermark_foto(
                    image_bytes=raw_bytes,
                    timestamp_str=now_str,
                    lokasi_str=lokasi.strip(),
                    gps_str=gps_input.strip()
                )

                conn = get_conn()
                cursor = conn.cursor()
                cursor.execute(
                    """INSERT INTO laporan_monitoring (
                        tanggal, cabor, lokasi, petugas,
                        fisik_parameter, fisik_peaking, fisik_recovery, fisik_cedera,
                        taktis_lawan, taktis_instruksi, taktis_ujicoba,
                        mental_cemas, mental_fokus, mental_rutinitas, mental_psikolog,
                        nutrisi_bb, nutrisi_asupan, nutrisi_hidrasi, nutrisi_tidur,
                        medis_rekam, medis_doping, medis_alat, medis_nonteknis,
                        status, catatan_admin, foto_data, koordinat_gps
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        tanggal, cabor, lokasi.strip(), st.session_state["username"],
                        fisik_1, fisik_2, fisik_3, fisik_4,
                        taktis_1, taktis_2, taktis_3,
                        mental_1, mental_2, mental_3, mental_4,
                        nutrisi_1, nutrisi_2, nutrisi_3, nutrisi_4,
                        medis_1, medis_2, medis_3, medis_4,
                        "Belum Ditindaklanjuti", "", watermarked_photo, gps_input.strip()
                    )
                )
                conn.commit()
                last_id = cursor.lastrowid
                conn.close()

                st.session_state["last_saved_id"] = last_id
                st.success("✅ Laporan berhasil disimpan beserta Watermark Time Mark & Lokasi!")

    # JIKA LAPORAN BARU SAJA DISIMPAN -> TAMPILKAN TOMBOL UNDUH WORD LANGSUNG
    if "last_saved_id" in st.session_state:
        conn = get_conn()
        saved_row = conn.execute("SELECT * FROM laporan_monitoring WHERE id=?", (st.session_state["last_saved_id"],)).fetchone()
        conn.close()

        if saved_row:
            st.markdown("### 📄 Unduh Laporan Siap Cetak")
            word_file = generate_word_report([dict(saved_row)], is_all=False)
            st.download_button(
                "⬇️ Unduh Laporan Word (.docx) Rapi & Profesional",
                data=word_file,
                file_name=f"Laporan_Monev_{saved_row['cabor']}_{saved_row['tanggal']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )


# ============================================================
# ROUTER
# ============================================================
init_db()

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    halaman_login()
else:
    if st.session_state.get("role") == "admin":
        halaman_admin()
    else:
        halaman_user()
