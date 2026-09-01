import streamlit as st
import sqlite3
import bcrypt
import pandas as pd
from docx import Document
from docx.shared import Inches
import io
import datetime

# ==========================================
# KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(
    page_title="Monitoring RIZAL DUDA",
    page_icon="🏅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 1. KONFIGURASI DATABASE & AUTHENTICATION
# ==========================================
DB_NAME = "monitoring.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    # Tabel Users
    c.execute('''CREATE TABLE IF NOT EXISTS users 
                 (username TEXT PRIMARY KEY, password TEXT, role TEXT)''')
    # Tabel Jadwal Monitoring
    c.execute('''CREATE TABLE IF NOT EXISTS schedules 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, cabor TEXT, tanggal TEXT, tempat TEXT)''')

    # Buat akun admin default jika belum ada
    c.execute("SELECT * FROM users WHERE username='admin'")
    if not c.fetchone():
        hashed_pw = bcrypt.hashpw('admin123'.encode('utf-8'), bcrypt.gensalt())
        c.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)", 
                  ('admin', hashed_pw, 'admin'))

    conn.commit()
    conn.close()

def authenticate(username, password):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT password, role FROM users WHERE username=?", (username,))
    result = c.fetchone()
    conn.close()

    if result:
        stored_password, role = result
        if bcrypt.checkpw(password.encode('utf-8'), stored_password):
            return True, role
    return False, None

# ==========================================
# 2. FUNGSI GENERATE WORD (.docx)
# ==========================================
def generate_word_report(petugas_text, cabor, tanggal, tempat, catatan, fotos):
    doc = Document()
    doc.add_heading('Laporan Monitoring Cabang Olahraga', 0)

    # Info Jadwal
    doc.add_paragraph(f"Cabang Olahraga\t: {cabor}")
    doc.add_paragraph(f"Tanggal\t\t: {tanggal}")
    doc.add_paragraph(f"Tempat\t\t: {tempat}")
    
    # Daftar Petugas (Dinamis)
    doc.add_heading('Daftar Petugas:', level=3)
    petugas_list = [p.strip() for p in petugas_text.split('\n') if p.strip()]
    for i, p in enumerate(petugas_list, 1):
        doc.add_paragraph(f"{i}. {p}")

    # Catatan
    doc.add_heading('Catatan Evaluasi / Hasil Monitoring:', level=3)
    doc.add_paragraph(catatan)

    # Multi Foto Dokumentasi
    if fotos:
        doc.add_heading('Foto Dokumentasi:', level=3)
        for idx, foto in enumerate(fotos):
            try:
                image_stream = io.BytesIO(foto.getvalue())
                # Menambahkan foto, diatur lebarnya agar tidak merusak margin kertas
                doc.add_picture(image_stream, width=Inches(4.5)) 
            except Exception as e:
                doc.add_paragraph(f"(Gagal memuat gambar ke-{idx+1}: {e})")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 3. HALAMAN & ANTARMUKA PENGGUNA
# ==========================================

init_db()

if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
    st.session_state['username'] = ''
    st.session_state['role'] = ''

# --- HALAMAN LOGIN ---
if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("<h1 style='text-align: center;'>🏆</h1>", unsafe_allow_html=True)
        st.markdown("<h2 style='text-align: center;'>APLIKASI MONITORING<br>RIZAL DUDA</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: gray;'>Binpres KONI Kabupaten Tangerang</p>", unsafe_allow_html=True)
        st.markdown("---")
        
        st.info("💡 **Info Login Admin:** Username: `admin` | Password: `admin123`")

        with st.form("login_form"):
            username_input = st.text_input("👤 Username")
            password_input = st.text_input("🔑 Password", type="password")
            submit_btn = st.form_submit_button("Masuk Sistem", use_container_width=True)

            if submit_btn:
                is_auth, role = authenticate(username_input, password_input)
                if is_auth:
                    st.session_state['logged_in'] = True
                    st.session_state['username'] = username_input
                    st.session_state['role'] = role
                    st.success("Login berhasil! Memuat dashboard...")
                    st.rerun()
                else:
                    st.error("🚨 Username atau password salah!")

# --- HALAMAN UTAMA (SETELAH LOGIN) ---
else:
    # Sidebar Menu
    st.sidebar.markdown("### 🏅 MONITORING RIZAL DUDA")
    st.sidebar.caption("Binpres KONI Kab. Tangerang")
    st.sidebar.markdown("---")
    
    st.sidebar.success(f"👋 Halo, **{st.session_state['username'].upper()}**\n\n🛡️ Akses: **{st.session_state['role'].upper()}**")
    
    st.sidebar.markdown("#### 📌 Menu Navigasi")
    
    # Logika Akses Menu
    menu = ["📝 Ambil Jadwal & Isi Form"]
    if st.session_state['role'] == 'admin':
        menu.insert(0, "📅 Kelola Jadwal (Admin)")
        menu.insert(1, "👥 Kelola User (Admin)")

    choice = st.sidebar.radio("", menu)
    
    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 Logout", use_container_width=True):
        st.session_state['logged_in'] = False
        st.session_state['username'] = ''
        st.session_state['role'] = ''
        st.rerun()


    # ----------------------------------------------------
    # MENU: KELOLA USER (ADMIN ONLY)
    # ----------------------------------------------------
    if choice == "👥 Kelola User (Admin)":
        st.title("👥 Manajemen Pengguna")
        
        tab1, tab2 = st.tabs(["Daftar Pengguna", "Tambah Pengguna Baru"])
        with tab1:
            conn = sqlite3.connect(DB_NAME)
            users_df = pd.read_sql_query("SELECT username, role FROM users", conn)
            conn.close()
            st.dataframe(users_df, use_container_width=True)

            with st.expander("🗑️ Hapus Pengguna", expanded=False):
                del_user = st.selectbox("Pilih pengguna yang akan dihapus", users_df['username'].tolist())
                if st.button("Hapus Akun", type="primary"):
                    if del_user == 'admin':
                        st.error("⚠️ Tidak bisa menghapus akun admin utama!")
                    else:
                        conn = sqlite3.connect(DB_NAME)
                        c = conn.cursor()
                        c.execute("DELETE FROM users WHERE username=?", (del_user,))
                        conn.commit()
                        conn.close()
                        st.success(f"✅ User **{del_user}** berhasil dihapus!")
                        st.rerun()

        with tab2:
            with st.form("add_user_form"):
                new_username = st.text_input("👤 Username Baru")
                new_password = st.text_input("🔑 Password Baru", type="password")
                new_role = st.selectbox("🛡️ Hak Akses (Role)", ["user", "admin"])
                submit_new_user = st.form_submit_button("💾 Simpan Pengguna")

                if submit_new_user:
                    if new_username and new_password:
                        try:
                            hashed_pw = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt())
                            conn = sqlite3.connect(DB_NAME)
                            c = conn.cursor()
                            c.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)", 
                                      (new_username, hashed_pw, new_role))
                            conn.commit()
                            conn.close()
                            st.success(f"✅ Pengguna baru **{new_username}** berhasil ditambahkan!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("⚠️ Username sudah terdaftar!")
                    else:
                        st.warning("⚠️ Username dan Password tidak boleh kosong!")


    # ----------------------------------------------------
    # MENU: KELOLA JADWAL (ADMIN ONLY)
    # ----------------------------------------------------
    elif choice == "📅 Kelola Jadwal (Admin)":
        st.title("📅 Kelola Jadwal Monitoring")
        st.markdown("Tambahkan jadwal cabor yang harus dimonitoring oleh petugas.")
        
        cabor_list = [
            "ANGGAR (IKASI)", "ATLETIK (PASI)", "ANGKAT BESI (PABSI)", "BOLA BASKET (PERBASI)", 
            "BOLA VOLI (PBVSI)", "BULU TANGKIS (PBSI)", "CATUR (PERCASI)", "E. SPORT",
            "KARATE (FORKI)", "PENCAK SILAT (IPSI)", "RENANG (PRSI)", "SEPAK BOLA (Askab-PSSI)", 
            "TAEKWONDO (TI)", "TENIS MEJA (PTMSI)", "TINJU (PERTINA)", "WUSHU (WI)"
            # Anda bisa menambahkan cabor lainnya ke dalam list ini
        ]
        
        col_form, col_data = st.columns([1, 2])
        
        with col_form:
            st.subheader("➕ Tambah Jadwal Baru")
            with st.form("form_jadwal"):
                new_cabor = st.selectbox("Cabang Olahraga", cabor_list)
                new_tanggal = st.date_input("Tanggal Kegiatan", datetime.date.today())
                new_tempat = st.text_input("Tempat / Lokasi")
                submit_jadwal = st.form_submit_button("Simpan Jadwal", use_container_width=True)
                
                if submit_jadwal:
                    if new_tempat:
                        conn = sqlite3.connect(DB_NAME)
                        c = conn.cursor()
                        c.execute("INSERT INTO schedules (cabor, tanggal, tempat) VALUES (?, ?, ?)", 
                                  (new_cabor, str(new_tanggal), new_tempat))
                        conn.commit()
                        conn.close()
                        st.success("✅ Jadwal berhasil ditambahkan!")
                        st.rerun()
                    else:
                        st.warning("⚠️ Tempat tidak boleh kosong!")
                        
        with col_data:
            st.subheader("📋 Daftar Jadwal Aktif")
            conn = sqlite3.connect(DB_NAME)
            jadwal_df = pd.read_sql_query("SELECT id, cabor as Cabor, tanggal as Tanggal, tempat as Tempat FROM schedules", conn)
            conn.close()
            
            if jadwal_df.empty:
                st.info("Belum ada jadwal yang dibuat.")
            else:
                st.dataframe(jadwal_df, use_container_width=True, hide_index=True)
                
                with st.expander("🗑️ Hapus Jadwal"):
                    del_id = st.selectbox("Pilih ID Jadwal yang akan dihapus", jadwal_df['id'].tolist())
                    if st.button("Hapus Jadwal Terpilih"):
                        conn = sqlite3.connect(DB_NAME)
                        c = conn.cursor()
                        c.execute("DELETE FROM schedules WHERE id=?", (del_id,))
                        conn.commit()
                        conn.close()
                        st.success("Jadwal berhasil dihapus!")
                        st.rerun()


    # ----------------------------------------------------
    # MENU: FORM LAPORAN (UNTUK USER & ADMIN)
    # ----------------------------------------------------
    elif choice == "📝 Ambil Jadwal & Isi Form":
        st.title("📝 Form Laporan Monitoring")
        
        # Mengambil data jadwal dari Database
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT id, cabor, tanggal, tempat FROM schedules")
        schedules_data = c.fetchall()
        conn.close()

        if not schedules_data:
            st.warning("⚠️ Belum ada jadwal monitoring yang tersedia. Harap hubungi Admin untuk membuat jadwal.")
        else:
            st.info("1️⃣ Pilih Jadwal yang tersedia di bawah ini, lalu isi daftar petugas dan catatan evaluasi.")
            
            # Format tampilan dropdown jadwal
            schedule_options = {}
            for s in schedules_data:
                label = f"{s[1]} | {s[2]} | {s[3]}"
                schedule_options[label] = s # Menyimpan raw data berdasarkan label

            selected_label = st.selectbox("Pilih Jadwal Monitoring", list(schedule_options.keys()))
            selected_schedule = schedule_options[selected_label]
            
            # Ekstrak data jadwal yang dipilih
            val_cabor = selected_schedule[1]
            val_tanggal = selected_schedule[2]
            val_tempat = selected_schedule[3]

            with st.form("form_laporan"):
                st.markdown("### 📋 Detail Petugas & Evaluasi")
                
                # Petugas menggunakan text_area agar bisa diisi banyak nama (1 per baris)
                petugas_text = st.text_area("👤 Daftar Petugas (Tulis 1 nama per baris)", 
                                           placeholder="Contoh:\nBudi Santoso\nAndi Saputra\nSiti Aminah", 
                                           help="Tulis satu nama per baris. Aplikasi otomatis akan menambahkan penomoran (1, 2, 3) di dokumen Word.")
                
                catatan = st.text_area("✍️ Catatan Evaluasi & Hasil Monitoring", height=150)
                
                st.markdown("### 📸 Upload Foto (Minimal 2, Maksimal 5)")
                fotos = st.file_uploader("Upload Foto Dokumentasi", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

                submit_laporan = st.form_submit_button("📄 Buat Laporan (Generate Word)", use_container_width=True)

            # --- VALIDASI & GENERATE LAPORAN ---
            if submit_laporan:
                if not petugas_text.strip():
                    st.warning("⚠️ Harap isi minimal 1 nama petugas!")
                elif not catatan.strip():
                    st.warning("⚠️ Catatan evaluasi tidak boleh kosong!")
                elif len(fotos) < 2:
                    st.error("🚨 Jumlah foto kurang! Harap unggah minimal 2 foto dokumentasi.")
                elif len(fotos) > 5:
                    st.error("🚨 Terlalu banyak foto! Batas maksimal adalah 5 foto.")
                else:
                    with st.spinner("⏳ Sedang menyusun dokumen Word..."):
                        word_file = generate_word_report(petugas_text, val_cabor, val_tanggal, val_tempat, catatan, fotos)
                        
                        st.success("🎉 Laporan berhasil dibuat! Silakan klik tombol di bawah untuk mengunduh.")
                        
                        st.download_button(
                            label="📥 Download Laporan Word (.docx)",
                            data=word_file,
                            file_name=f"Laporan_Monev_{val_cabor.split(' ')[0]}_{val_tanggal}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )