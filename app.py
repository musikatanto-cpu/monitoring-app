import streamlit as st
import datetime
import sqlite3
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches

# --- 1. INISIALISASI DATABASE ---
def init_db():
    conn = sqlite3.connect('monitoring.db')
    cursor = conn.cursor()
    
    # Tabel Pengguna (Admin & User)
    cursor.execute('''CREATE TABLE IF NOT EXISTS users (
                        id INTEGER PRIMARY KEY, username TEXT, password TEXT, role TEXT)''')
    
    # Tabel Cabang Olahraga
    cursor.execute('''CREATE TABLE IF NOT EXISTS cabor (
                        id INTEGER PRIMARY KEY, nama TEXT)''')
    
    # Tabel Laporan (Diperbarui dengan pembuat laporan)
    cursor.execute('''CREATE TABLE IF NOT EXISTS laporan_monitoring (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        tanggal DATE, cabor TEXT, lokasi TEXT, petugas TEXT,
                        fisik_parameter TEXT, fisik_peaking TEXT, fisik_recovery TEXT, fisik_cedera TEXT,
                        taktis_lawan TEXT, taktis_instruksi TEXT, taktis_ujicoba TEXT,
                        mental_cemas TEXT, mental_fokus TEXT, mental_rutinitas TEXT, mental_psikolog TEXT,
                        nutrisi_bb TEXT, nutrisi_asupan TEXT, nutrisi_hidrasi TEXT, nutrisi_tidur TEXT,
                        medis_rekam TEXT, medis_doping TEXT, medis_alat TEXT, medis_nonteknis TEXT)''')
    
    # Insert Default Users (Jika kosong)
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO users (username, password, role) VALUES ('admin', 'admin123', 'admin')")
        cursor.execute("INSERT INTO users (username, password, role) VALUES ('petugas', 'petugas123', 'user')")
        
    # Insert Default Cabor (Jika kosong)
    cursor.execute("SELECT COUNT(*) FROM cabor")
    if cursor.fetchone()[0] == 0:
        default_cabor = ['Panahan', 'Taekwondo', 'Sepatu Roda', 'Judo', 'Sepak Takraw', 'Catur', 'Sepak Bola']
        for c in default_cabor:
            cursor.execute("INSERT INTO cabor (nama) VALUES (?)", (c,))
            
    conn.commit()
    conn.close()

# --- 2. GENERATE WORD DOCUMENT ---
def generate_word_report(data):
    doc = Document()
    
    # Judul Dokumen
    title = doc.add_heading('LAPORAN MONEV BINPRES', 0)
    title.alignment = 1 # Center
    
    # Header Info
    doc.add_paragraph(f"Cabang Olahraga\t: {data['cabor']}")
    doc.add_paragraph(f"Tanggal\t\t: {data['tanggal']}")
    doc.add_paragraph(f"Lokasi\t\t: {data['lokasi']}")
    doc.add_paragraph(f"Petugas Monev\t: {data['petugas']}")
    doc.add_paragraph("-" * 50)
    
    # Fungsi pembantu untuk membuat sub-bab
    def add_section(title, questions_answers):
        doc.add_heading(title, level=2)
        for q, a in questions_answers:
            p = doc.add_paragraph()
            p.add_run(q).bold = True
            doc.add_paragraph(a if a else "-")
    
    add_section('1. Performa Fisik & Kebugaran', [
        ('Capaian parameter fisik (kekuatan, daya tahan, kecepatan, kelincahan):', data['fisik_parameter']),
        ('Apakah atlet mencapai grafik performa puncak (peaking)?', data['fisik_peaking']),
        ('Tingkat pemulihan fisik (recovery) harian pasca-latihan:', data['fisik_recovery']),
        ('Keluhan cedera lama/baru:', data['fisik_cedera'])
    ])
    
    add_section('2. Kesiapan Taktis & Penguasaan Strategi', [
        ('Pemetaan kekuatan calon lawan:', data['taktis_lawan']),
        ('Kemampuan mengikuti instruksi teknis di bawah tekanan:', data['taktis_instruksi']),
        ('Hasil try-out / sparing:', data['taktis_ujicoba'])
    ])
    
    add_section('3. Mental, Psikologis & Kesiapan Mental', [
        ('Tingkat kecemasan & pengendalian stres:', data['mental_cemas']),
        ('Fokus, motivasi, dan self-confidence:', data['mental_fokus']),
        ('Rutinitas mental khusus saat bertanding:', data['mental_rutinitas']),
        ('Koordinasi dengan psikolog olahraga:', data['mental_psikolog'])
    ])
    
    add_section('4. Nutrisi, Berat Badan & Gaya Hidup', [
        ('Progres penyesuaian berat badan:', data['nutrisi_bb']),
        ('Asupan nutrisi dan suplemen:', data['nutrisi_asupan']),
        ('Status hidrasi:', data['nutrisi_hidrasi']),
        ('Kualitas dan kecukupan tidur:', data['nutrisi_tidur'])
    ])
    
    add_section('5. Medis, Bebas Doping & Logistik', [
        ('Status rekam medis & kesiapan tim medis:', data['medis_rekam']),
        ('Keamanan obat, suplemen (Bebas Doping):', data['medis_doping']),
        ('Kesiapan perlengkapan tanding:', data['medis_alat']),
        ('Kendala non-teknis (akomodasi, transportasi):', data['medis_nonteknis'])
    ])
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 3. HALAMAN LOGIN ---
def halaman_login():
    st.title("🔐 Login Sistem Monitoring")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Login")
        
        if submit:
            conn = sqlite3.connect('monitoring.db')
            cursor = conn.cursor()
            cursor.execute("SELECT role FROM users WHERE username=? AND password=?", (username, password))
            user = cursor.fetchone()
            conn.close()
            
            if user:
                st.session_state['logged_in'] = True
                st.session_state['username'] = username
                st.session_state['role'] = user[0]
                st.rerun()
            else:
                st.error("Username atau Password salah!")

# --- 4. HALAMAN ADMIN ---
def halaman_admin():
    st.title("👨‍💼 Dashboard Admin")
    st.write(f"Selamat datang, **{st.session_state['username']}**!")
    if st.button("Logout"):
        st.session_state.clear()
        st.rerun()

    tab1, tab2 = st.tabs(["📄 Data Laporan Masuk", "⚙️ Kelola Cabang Olahraga"])
    
    with tab1:
        st.subheader("Semua Laporan Monitoring")
        conn = sqlite3.connect('monitoring.db')
        df = pd.read_sql_query("SELECT id, tanggal, cabor, lokasi, petugas FROM laporan_monitoring ORDER BY id DESC", conn)
        
        if df.empty:
            st.info("Belum ada laporan masuk.")
        else:
            st.dataframe(df, use_container_width=True)
            
            st.markdown("---")
            st.write("**Unduh Laporan ke Word (.docx)**")
            pilih_id = st.selectbox("Pilih ID Laporan untuk diunduh:", df['id'].tolist())
            
            if st.button("Generate Dokumen"):
                cursor = conn.cursor()
                # Ambil semua data berdasarkan ID
                cursor.execute("SELECT * FROM laporan_monitoring WHERE id=?", (pilih_id,))
                row = cursor.fetchone()
                # Mapping nama kolom agar sesuai fungsi docx
                col_names = [description[0] for description in cursor.description]
                data_dict = dict(zip(col_names, row))
                
                word_file = generate_word_report(data_dict)
                st.download_button(
                    label="📥 Download File Word",
                    data=word_file,
                    file_name=f"Laporan_{data_dict['cabor']}_{data_dict['tanggal']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        conn.close()

    with tab2:
        st.subheader("Daftar Cabang Olahraga")
        conn = sqlite3.connect('monitoring.db')
        cabor_df = pd.read_sql_query("SELECT * FROM cabor", conn)
        st.dataframe(cabor_df, use_container_width=True, hide_index=True)
        
        with st.form("tambah_cabor"):
            cabor_baru = st.text_input("Tambah Cabor Baru")
            if st.form_submit_button("Tambah"):
                if cabor_baru:
                    cursor = conn.cursor()
                    cursor.execute("INSERT INTO cabor (nama) VALUES (?)", (cabor_baru,))
                    conn.commit()
                    st.success(f"{cabor_baru} berhasil ditambahkan!")
                    st.rerun()
        conn.close()

# --- 5. HALAMAN USER (INPUT FORM) ---
def halaman_user():
    st.title("📝 Form Input Monitoring")
    st.write(f"Petugas: **{st.session_state['username']}**")
    if st.button("Logout"):
        st.session_state.clear()
        st.rerun()
        
    conn = sqlite3.connect('monitoring.db')
    cursor = conn.cursor()
    cursor.execute("SELECT nama FROM cabor ORDER BY nama")
    list_cabor = [row[0] for row in cursor.fetchall()]
    conn.close()

    with st.form("form_monitoring", clear_on_submit=True):
        st.subheader("Informasi Dasar")
        col1, col2 = st.columns(2)
        with col1:
            tanggal = st.date_input("Tanggal Monitoring", datetime.date.today())
            cabor = st.selectbox("Cabang Olahraga", ["Pilih Cabor..."] + list_cabor)
        with col2:
            lokasi = st.text_input("Lokasi Latihan / Try-out")

        # Expander Indikator (Sesuai Materi)
        with st.expander("💪 1. Performa Fisik & Kebugaran"):
            fisik_1 = st.text_area("Capaian parameter fisik (vs benchmark target):", height=68)
            fisik_2 = st.text_area("Apakah atlet mencapai grafik performa puncak (peaking)?", height=68)
            fisik_3 = st.text_area("Tingkat pemulihan fisik (recovery):", height=68)
            fisik_4 = st.text_area("Keluhan cedera lama / indikasi cedera baru:", height=68)

        with st.expander("🎯 2. Kesiapan Taktis & Strategi"):
            taktis_1 = st.text_area("Pemetaan kekuatan calon lawan:", height=68)
            taktis_2 = st.text_area("Kemampuan mengikuti instruksi teknis di bawah tekanan:", height=68)
            taktis_3 = st.text_area("Hasil try-out / sparing (peningkatan efektivitas):", height=68)

        with st.expander("🧠 3. Mental, Psikologis & Kesiapan Mental"):
            mental_1 = st.text_area("Tingkat kecemasan & kemampuan mengendalikan stres:", height=68)
            mental_2 = st.text_area("Fokus, motivasi, dan self-confidence:", height=68)
            mental_3 = st.text_area("Rutinitas mental khusus saat masuk lapangan:", height=68)
            mental_4 = st.text_area("Koordinasi dengan tim psikolog olahraga:", height=68)

        with st.expander("🥗 4. Nutrisi, Berat Badan & Gaya Hidup"):
            nutrisi_1 = st.text_area("Progres penyesuaian berat badan (weight management):", height=68)
            nutrisi_2 = st.text_area("Pemantauan asupan nutrisi dan suplemen harian:", height=68)
            nutrisi_3 = st.text_area("Status hidrasi harian atlet:", height=68)
            nutrisi_4 = st.text_area("Kualitas dan kecukupan waktu tidur:", height=68)

        with st.expander("⚕️ 5. Medis, Bebas Doping & Logistik"):
            medis_1 = st.text_area("Status rekam medis terkini & kesiapan fisioterapis:", height=68)
            medis_2 = st.text_area("Keamanan obat, suplemen (Bebas Doping):", height=68)
            medis_3 = st.text_area("Kesiapan perlengkapan khusus bertanding:", height=68)
            medis_4 = st.text_area("Kendala non-teknis (akomodasi, transportasi):", height=68)

        submitted = st.form_submit_button("💾 Simpan Laporan", use_container_width=True)
        
        if submitted:
            if cabor == "Pilih Cabor...":
                st.error("⚠️ Harap pilih Cabang Olahraga!")
            elif not lokasi:
                st.error("⚠️ Lokasi wajib diisi!")
            else:
                conn = sqlite3.connect('monitoring.db')
                cursor = conn.cursor()
                cursor.execute('''INSERT INTO laporan_monitoring (
                    tanggal, cabor, lokasi, petugas,
                    fisik_parameter, fisik_peaking, fisik_recovery, fisik_cedera,
                    taktis_lawan, taktis_instruksi, taktis_ujicoba,
                    mental_cemas, mental_fokus, mental_rutinitas, mental_psikolog,
                    nutrisi_bb, nutrisi_asupan, nutrisi_hidrasi, nutrisi_tidur,
                    medis_rekam, medis_doping, medis_alat, medis_nonteknis
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', 
                (tanggal, cabor, lokasi, st.session_state['username'],
                 fisik_1, fisik_2, fisik_3, fisik_4,
                 taktis_1, taktis_2, taktis_3,
                 mental_1, mental_2, mental_3, mental_4,
                 nutrisi_1, nutrisi_2, nutrisi_3, nutrisi_4,
                 medis_1, medis_2, medis_3, medis_4))
                conn.commit()
                conn.close()
                st.success("✅ Laporan berhasil disimpan! Admin kini bisa mengunduhnya dalam format Word.")

# --- 6. ROUTER APLIKASI ---
init_db()

if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    halaman_login()
else:
    if st.session_state['role'] == 'admin':
        halaman_admin()
    else:
        halaman_user()
