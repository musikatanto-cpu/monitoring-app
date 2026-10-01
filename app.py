import streamlit as st
import datetime
import sqlite3
import pandas as pd
import hashlib
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches

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
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Migrasi sederhana agar database lama tetap dapat dipakai.
    existing_users = [r["name"] for r in cursor.execute("PRAGMA table_info(users)").fetchall()]
    user_columns = {
        "nama_lengkap": "TEXT DEFAULT ''",
        "aktif": "INTEGER DEFAULT 1",
        "created_at": "TEXT DEFAULT CURRENT_TIMESTAMP"
    }
    for col, definition in user_columns.items():
        if col not in existing_users:
            cursor.execute(f"ALTER TABLE users ADD COLUMN {col} {definition}")

    existing_reports = [r["name"] for r in cursor.execute("PRAGMA table_info(laporan_monitoring)").fetchall()]
    report_columns = {
        "status": "TEXT DEFAULT 'Belum Ditindaklanjuti'",
        "catatan_admin": "TEXT DEFAULT ''",
        "updated_at": "TEXT DEFAULT CURRENT_TIMESTAMP"
    }
    for col, definition in report_columns.items():
        if col not in existing_reports:
            cursor.execute(f"ALTER TABLE laporan_monitoring ADD COLUMN {col} {definition}")

    # Default account. Password lama tetap bisa digunakan saat migrasi,
    # tetapi akun baru disimpan menggunakan hash.
    count = cursor.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]
    if count == 0:
        cursor.execute(
            """INSERT INTO users
               (username, password, role, nama_lengkap, aktif)
               VALUES (?, ?, ?, ?, 1)""",
            ("admin", hash_password("admin123"), "admin", "Administrator")
        )
        cursor.execute(
            """INSERT INTO users
               (username, password, role, nama_lengkap, aktif)
               VALUES (?, ?, ?, ?, 1)""",
            ("petugas", hash_password("petugas123"), "user", "Petugas Monitoring")
        )
    else:
        # Jika database berasal dari versi lama, upgrade password plaintext
        # untuk akun default yang masih menggunakan password lama.
        for username, old_password, role, nama in [
            ("admin", "admin123", "admin", "Administrator"),
            ("petugas", "petugas123", "user", "Petugas Monitoring")
        ]:
            row = cursor.execute(
                "SELECT id, password FROM users WHERE username=?",
                (username,)
            ).fetchone()
            if row and row["password"] == old_password:
                cursor.execute(
                    "UPDATE users SET password=?, role=?, nama_lengkap=? WHERE id=?",
                    (hash_password(old_password), role, nama, row["id"])
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
        cursor.executemany(
            "INSERT OR IGNORE INTO cabor (nama) VALUES (?)",
            [(x,) for x in default_cabor]
        )

    conn.commit()
    conn.close()


# ============================================================
# HELPER
# ============================================================
def logout():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()


def get_cabor_list():
    conn = get_conn()
    rows = conn.execute(
        "SELECT nama FROM cabor WHERE aktif=1 ORDER BY nama"
    ).fetchall()
    conn.close()
    return [r["nama"] for r in rows]


def get_current_user():
    if "username" not in st.session_state:
        return None
    conn = get_conn()
    row = conn.execute(
        "SELECT * FROM users WHERE username=?",
        (st.session_state["username"],)
    ).fetchone()
    conn.close()
    return row


# ============================================================
# WORD REPORT
# ============================================================
def generate_word_report(data_list, is_all=False):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    judul_utama = (
        "REKAPITULASI SEMUA LAPORAN MONEV BINPRES"
        if is_all else "LAPORAN MONEV BINPRES"
    )
    title = doc.add_heading(judul_utama, 0)
    title.alignment = 1

    for idx, data in enumerate(data_list):
        if is_all:
            doc.add_heading(
                f"Laporan {idx + 1}: {data['cabor']} - {data['tanggal']}",
                level=1
            )

        doc.add_paragraph(f"Cabang Olahraga\t: {data['cabor']}")
        doc.add_paragraph(f"Tanggal\t\t: {data['tanggal']}")
        doc.add_paragraph(f"Lokasi\t\t: {data['lokasi']}")
        doc.add_paragraph(f"Petugas Monev\t: {data['petugas']}")
        doc.add_paragraph(f"Status\t\t: {data.get('status', '-')}")
        doc.add_paragraph("-" * 70)

        def add_section(title_text, questions_answers):
            doc.add_heading(title_text, level=2)
            for q, a in questions_answers:
                p = doc.add_paragraph()
                p.add_run(q).bold = True
                doc.add_paragraph(a if a else "-")

        add_section("1. Performa Fisik & Kebugaran", [
            ("Capaian parameter fisik:", data["fisik_parameter"]),
            ("Grafik performa puncak (peaking):", data["fisik_peaking"]),
            ("Tingkat pemulihan fisik (recovery):", data["fisik_recovery"]),
            ("Keluhan cedera lama/baru:", data["fisik_cedera"])
        ])

        add_section("2. Kesiapan Taktis & Penguasaan Strategi", [
            ("Pemetaan kekuatan calon lawan:", data["taktis_lawan"]),
            ("Kemampuan mengikuti instruksi teknis:", data["taktis_instruksi"]),
            ("Hasil try-out / sparing:", data["taktis_ujicoba"])
        ])

        add_section("3. Mental, Psikologis & Kesiapan Mental", [
            ("Tingkat kecemasan & pengendalian stres:", data["mental_cemas"]),
            ("Fokus, motivasi, dan self-confidence:", data["mental_fokus"]),
            ("Rutinitas mental khusus:", data["mental_rutinitas"]),
            ("Koordinasi dengan psikolog olahraga:", data["mental_psikolog"])
        ])

        add_section("4. Nutrisi, Berat Badan & Gaya Hidup", [
            ("Progres penyesuaian berat badan:", data["nutrisi_bb"]),
            ("Asupan nutrisi dan suplemen:", data["nutrisi_asupan"]),
            ("Status hidrasi:", data["nutrisi_hidrasi"]),
            ("Kualitas dan kecukupan tidur:", data["nutrisi_tidur"])
        ])

        add_section("5. Medis, Bebas Doping & Logistik", [
            ("Status rekam medis & tim medis:", data["medis_rekam"]),
            ("Keamanan obat/suplemen:", data["medis_doping"]),
            ("Kesiapan perlengkapan tanding:", data["medis_alat"]),
            ("Kendala non-teknis:", data["medis_nonteknis"])
        ])

        if data.get("catatan_admin"):
            add_section("Catatan Admin / Tindak Lanjut", [
                ("Catatan:", data["catatan_admin"])
            ])

        if idx < len(data_list) - 1:
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
        submit = st.form_submit_button(
            "Masuk ke Sistem",
            use_container_width=True,
            type="primary"
        )

        if submit:
            conn = get_conn()
            user = conn.execute(
                "SELECT * FROM users WHERE username=? AND aktif=1",
                (username.strip(),)
            ).fetchone()
            conn.close()

            valid = False
            if user:
                # Mendukung hash baru dan password plaintext lama.
                valid = (
                    verify_password(password, user["password"])
                    or password == user["password"]
                )

            if valid:
                st.session_state["logged_in"] = True
                st.session_state["username"] = user["username"]
                st.session_state["role"] = user["role"]
                st.session_state["nama_lengkap"] = user["nama_lengkap"] or user["username"]
                st.rerun()
            else:
                st.error("Username atau Password salah / akun tidak aktif.")

    st.caption("Versi pengembangan • Data tersimpan pada database SQLite lokal")


# ============================================================
# SIDEBAR
# ============================================================
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
# DASHBOARD ADMIN
# ============================================================
def dashboard_admin():
    conn = get_conn()

    total = conn.execute(
        "SELECT COUNT(*) AS n FROM laporan_monitoring"
    ).fetchone()["n"]
    bulan_ini = conn.execute(
        """SELECT COUNT(*) AS n FROM laporan_monitoring
           WHERE strftime('%Y-%m', tanggal)=strftime('%Y-%m','now')"""
    ).fetchone()["n"]
    cabor_termonitor = conn.execute(
        "SELECT COUNT(DISTINCT cabor) AS n FROM laporan_monitoring"
    ).fetchone()["n"]
    tindak = conn.execute(
        """SELECT COUNT(*) AS n FROM laporan_monitoring
           WHERE status='Belum Ditindaklanjuti' OR status IS NULL"""
    ).fetchone()["n"]

    conn.close()

    st.markdown("## 📊 Dashboard Monitoring")
    st.caption("Ringkasan aktivitas monitoring BINPRES")

    a, b, c, d = st.columns(4)
    a.metric("Total Laporan", total)
    b.metric("Laporan Bulan Ini", bulan_ini)
    c.metric("Cabor Termonitor", cabor_termonitor)
    d.metric("Perlu Tindak Lanjut", tindak)

    st.markdown("---")


# ============================================================
# KELOLA USER
# ============================================================
def kelola_user():
    st.subheader("👥 Manajemen User")

    conn = get_conn()
    df = pd.read_sql_query(
        """SELECT id, username, nama_lengkap, role,
                  CASE WHEN aktif=1 THEN 'Aktif' ELSE 'Nonaktif' END AS status,
                  created_at
           FROM users ORDER BY id DESC""",
        conn
    )
    conn.close()

    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("### ➕ Tambah User")
    with st.form("tambah_user"):
        c1, c2 = st.columns(2)
        with c1:
            username = st.text_input("Username baru")
            nama = st.text_input("Nama lengkap")
        with c2:
            password = st.text_input("Password", type="password")
            role = st.selectbox("Hak akses", ["user", "admin"])

        submit = st.form_submit_button(
            "Simpan User",
            use_container_width=True,
            type="primary"
        )

        if submit:
            if not username.strip() or not password:
                st.error("Username dan password wajib diisi.")
            elif len(password) < 6:
                st.error("Password minimal 6 karakter.")
            else:
                conn = get_conn()
                try:
                    conn.execute(
                        """INSERT INTO users
                           (username,password,role,nama_lengkap,aktif)
                           VALUES (?,?,?,?,1)""",
                        (username.strip(), hash_password(password), role, nama.strip())
                    )
                    conn.commit()
                    st.success("User berhasil ditambahkan.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Username sudah digunakan.")
                finally:
                    conn.close()

    st.markdown("### ✏️ Edit / Nonaktifkan User")
    conn = get_conn()
    users = [
        dict(r) for r in conn.execute(
            "SELECT id, username, nama_lengkap, role, aktif FROM users ORDER BY username"
        ).fetchall()
    ]
    conn.close()

    if users:
        pilihan = st.selectbox(
            "Pilih user",
            users,
            format_func=lambda x: f"{x['username']} — {x['nama_lengkap'] or '-'}"
        )

        with st.form("edit_user"):
            nama_baru = st.text_input(
                "Nama lengkap",
                value=pilihan["nama_lengkap"] or ""
            )
            role_baru = st.selectbox(
                "Role",
                ["user", "admin"],
                index=0 if pilihan["role"] == "user" else 1
            )
            aktif_baru = st.checkbox(
                "Akun aktif",
                value=bool(pilihan["aktif"])
            )
            password_baru = st.text_input(
                "Password baru (kosongkan jika tidak diubah)",
                type="password"
            )

            simpan = st.form_submit_button(
                "💾 Simpan Perubahan",
                use_container_width=True
            )

            if simpan:
                conn = get_conn()
                if password_baru:
                    conn.execute(
                        """UPDATE users
                           SET nama_lengkap=?, role=?, aktif=?, password=?
                           WHERE id=?""",
                        (
                            nama_baru.strip(),
                            role_baru,
                            1 if aktif_baru else 0,
                            hash_password(password_baru),
                            pilihan["id"]
                        )
                    )
                else:
                    conn.execute(
                        """UPDATE users
                           SET nama_lengkap=?, role=?, aktif=?
                           WHERE id=?""",
                        (
                            nama_baru.strip(),
                            role_baru,
                            1 if aktif_baru else 0,
                            pilihan["id"]
                        )
                    )
                conn.commit()
                conn.close()
                st.success("Data user berhasil diperbarui.")
                st.rerun()


# ============================================================
# KELOLA CABOR
# ============================================================
def kelola_cabor():
    st.subheader("🏅 Kelola Cabang Olahraga")

    conn = get_conn()
    df = pd.read_sql_query(
        """SELECT id, nama,
                  CASE WHEN aktif=1 THEN 'Aktif' ELSE 'Nonaktif' END AS status
           FROM cabor ORDER BY nama""",
        conn
    )
    conn.close()
    st.dataframe(df, use_container_width=True, hide_index=True)

    with st.form("tambah_cabor"):
        nama_cabor = st.text_input("Nama Cabang Olahraga Baru")
        submit = st.form_submit_button("➕ Tambah Cabor", use_container_width=True)

        if submit:
            if not nama_cabor.strip():
                st.error("Nama cabor wajib diisi.")
            else:
                conn = get_conn()
                try:
                    conn.execute(
                        "INSERT INTO cabor (nama) VALUES (?)",
                        (nama_cabor.strip().upper(),)
                    )
                    conn.commit()
                    st.success("Cabor berhasil ditambahkan.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Cabor tersebut sudah ada.")
                finally:
                    conn.close()

    st.markdown("### ✏️ Edit / Nonaktifkan Cabor")
    conn = get_conn()
    rows = [
        dict(r) for r in conn.execute(
            "SELECT * FROM cabor ORDER BY nama"
        ).fetchall()
    ]
    conn.close()

    if rows:
        pilihan = st.selectbox(
            "Pilih cabor",
            rows,
            format_func=lambda x: x["nama"]
        )
        with st.form("edit_cabor"):
            nama_baru = st.text_input("Nama cabor", value=pilihan["nama"])
            aktif = st.checkbox("Cabor aktif", value=bool(pilihan["aktif"]))
            simpan = st.form_submit_button("💾 Simpan", use_container_width=True)
            if simpan:
                conn = get_conn()
                try:
                    conn.execute(
                        "UPDATE cabor SET nama=?, aktif=? WHERE id=?",
                        (nama_baru.strip().upper(), 1 if aktif else 0, pilihan["id"])
                    )
                    conn.commit()
                    st.success("Cabor berhasil diperbarui.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Nama cabor sudah digunakan.")
                finally:
                    conn.close()


# ============================================================
# LAPORAN ADMIN
# ============================================================
def halaman_laporan_admin():
    st.subheader("📄 Data Laporan Monitoring")

    conn = get_conn()
    df = pd.read_sql_query(
        """SELECT id, tanggal, cabor, lokasi, petugas, status
           FROM laporan_monitoring
           ORDER BY tanggal DESC, id DESC""",
        conn
    )
    conn.close()

    if df.empty:
        st.info("Belum ada laporan masuk.")
        return

    f1, f2, f3 = st.columns(3)
    with f1:
        cabor_filter = st.selectbox(
            "Filter Cabor",
            ["Semua"] + sorted(df["cabor"].dropna().unique().tolist())
        )
    with f2:
        status_filter = st.selectbox(
            "Filter Status",
            ["Semua"] + sorted(df["status"].fillna("Belum Ditindaklanjuti").unique().tolist())
        )
    with f3:
        keyword = st.text_input("🔎 Cari lokasi / petugas")

    filtered = df.copy()
    if cabor_filter != "Semua":
        filtered = filtered[filtered["cabor"] == cabor_filter]
    if status_filter != "Semua":
        filtered = filtered[filtered["status"].fillna("Belum Ditindaklanjuti") == status_filter]
    if keyword:
        mask = (
            filtered["lokasi"].fillna("").str.contains(keyword, case=False, na=False)
            | filtered["petugas"].fillna("").str.contains(keyword, case=False, na=False)
        )
        filtered = filtered[mask]

    st.caption(f"Menampilkan {len(filtered)} dari {len(df)} laporan.")
    st.dataframe(filtered, use_container_width=True, hide_index=True)

    if filtered.empty:
        return

    st.markdown("---")
    st.markdown("### 📥 Export Laporan")

    c1, c2 = st.columns(2)
    with c1:
        pilihan_id = st.selectbox(
            "Pilih ID laporan untuk Word",
            filtered["id"].tolist()
        )
        if st.button("📄 Generate Word Satuan", use_container_width=True):
            conn = get_conn()
            row = conn.execute(
                "SELECT * FROM laporan_monitoring WHERE id=?",
                (int(pilihan_id),)
            ).fetchone()
            conn.close()

            word_file = generate_word_report([dict(row)], is_all=False)
            st.download_button(
                "⬇️ Download Word",
                data=word_file,
                file_name=f"Laporan_{row['cabor']}_{row['tanggal']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    with c2:
        st.write("**Rekap hasil filter**")
        if st.button("📚 Generate Word Rekap", use_container_width=True):
            conn = get_conn()
            ids = tuple(int(x) for x in filtered["id"].tolist())
            placeholders = ",".join("?" * len(ids))
            rows = conn.execute(
                f"SELECT * FROM laporan_monitoring WHERE id IN ({placeholders}) ORDER BY tanggal ASC, id ASC",
                ids
            ).fetchall()
            conn.close()

            word_file = generate_word_report([dict(r) for r in rows], is_all=True)
            st.download_button(
                "⬇️ Download Rekap Word",
                data=word_file,
                file_name=f"Rekap_Monev_{datetime.date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    st.markdown("---")
    st.markdown("### ✏️ Tindak Lanjut Laporan")

    pilihan_edit = st.selectbox(
        "Pilih laporan yang akan diperbarui",
        filtered["id"].tolist(),
        key="edit_laporan_id"
    )

    conn = get_conn()
    row = conn.execute(
        "SELECT * FROM laporan_monitoring WHERE id=?",
        (int(pilihan_edit),)
    ).fetchone()
    conn.close()

    if row:
        with st.form("update_status_laporan"):
            status = st.selectbox(
                "Status tindak lanjut",
                [
                    "Belum Ditindaklanjuti",
                    "Sedang Ditindaklanjuti",
                    "Selesai"
                ],
                index=[
                    "Belum Ditindaklanjuti",
                    "Sedang Ditindaklanjuti",
                    "Selesai"
                ].index(row["status"] or "Belum Ditindaklanjuti")
            )
            catatan = st.text_area(
                "Catatan Admin / Rekomendasi",
                value=row["catatan_admin"] or ""
            )
            simpan = st.form_submit_button(
                "💾 Simpan Tindak Lanjut",
                use_container_width=True
            )

            if simpan:
                conn = get_conn()
                conn.execute(
                    """UPDATE laporan_monitoring
                       SET status=?, catatan_admin=?, updated_at=CURRENT_TIMESTAMP
                       WHERE id=?""",
                    (status, catatan, int(pilihan_edit))
                )
                conn.commit()
                conn.close()
                st.success("Tindak lanjut berhasil diperbarui.")
                st.rerun()


# ============================================================
# HALAMAN ADMIN
# ============================================================
def halaman_admin():
    sidebar_app()
    dashboard_admin()

    tab1, tab2, tab3 = st.tabs([
        "📄 Laporan Monitoring",
        "👥 Manajemen User",
        "🏅 Kelola Cabor"
    ])

    with tab1:
        halaman_laporan_admin()

    with tab2:
        kelola_user()

    with tab3:
        kelola_cabor()


# ============================================================
# HALAMAN USER / PETUGAS
# ============================================================
def halaman_user():
    sidebar_app()

    st.markdown("## 📝 Form Input Monitoring")
    st.caption(
        f"Petugas: **{st.session_state.get('nama_lengkap', st.session_state['username'])}**"
    )

    cabor_list = get_cabor_list()

    with st.form("form_monitoring", clear_on_submit=True):
        st.markdown("### 📌 Informasi Dasar")
        col1, col2 = st.columns(2)

        with col1:
            tanggal = st.date_input(
                "Tanggal Monitoring",
                datetime.date.today()
            )
            cabor = st.selectbox(
                "Cabang Olahraga",
                ["Pilih Cabor..."] + cabor_list
            )

        with col2:
            lokasi = st.text_input("Lokasi Latihan / Try-out")

        st.markdown("---")

        with st.expander("💪 1. Performa Fisik & Kebugaran", expanded=True):
            fisik_1 = st.text_area("Capaian parameter fisik (vs benchmark target):", height=68)
            fisik_2 = st.text_area("Apakah atlet mencapai grafik performa puncak (peaking)?", height=68)
            fisik_3 = st.text_area("Tingkat pemulihan fisik (recovery):", height=68)
            fisik_4 = st.text_area("Keluhan cedera lama / indikasi cedera baru:", height=68)

        with st.expander("🎯 2. Kesiapan Taktis & Strategi"):
            taktis_1 = st.text_area("Pemetaan kekuatan calon lawan:", height=68)
            taktis_2 = st.text_area("Kemampuan mengikuti instruksi teknis di bawah tekanan:", height=68)
            taktis_3 = st.text_area("Hasil try-out / sparing:", height=68)

        with st.expander("🧠 3. Mental, Psikologis & Kesiapan Mental"):
            mental_1 = st.text_area("Tingkat kecemasan & kemampuan mengendalikan stres:", height=68)
            mental_2 = st.text_area("Fokus, motivasi, dan self-confidence:", height=68)
            mental_3 = st.text_area("Rutinitas mental khusus saat bertanding:", height=68)
            mental_4 = st.text_area("Koordinasi dengan tim psikolog olahraga:", height=68)

        with st.expander("🥗 4. Nutrisi, Berat Badan & Gaya Hidup"):
            nutrisi_1 = st.text_area("Progres penyesuaian berat badan:", height=68)
            nutrisi_2 = st.text_area("Pemantauan asupan nutrisi dan suplemen:", height=68)
            nutrisi_3 = st.text_area("Status hidrasi harian atlet:", height=68)
            nutrisi_4 = st.text_area("Kualitas dan kecukupan waktu tidur:", height=68)

        with st.expander("⚕️ 5. Medis, Bebas Doping & Logistik"):
            medis_1 = st.text_area("Status rekam medis terkini & kesiapan fisioterapis:", height=68)
            medis_2 = st.text_area("Keamanan obat, suplemen (Bebas Doping):", height=68)
            medis_3 = st.text_area("Kesiapan perlengkapan khusus bertanding:", height=68)
            medis_4 = st.text_area("Kendala non-teknis (akomodasi, transportasi):", height=68)

        st.markdown("---")
        submitted = st.form_submit_button(
            "💾 Simpan Laporan Monitoring",
            use_container_width=True,
            type="primary"
        )

        if submitted:
            if cabor == "Pilih Cabor...":
                st.error("⚠️ Harap pilih Cabang Olahraga.")
            elif not lokasi.strip():
                st.error("⚠️ Lokasi wajib diisi.")
            else:
                conn = get_conn()
                conn.execute(
                    """INSERT INTO laporan_monitoring (
                        tanggal, cabor, lokasi, petugas,
                        fisik_parameter, fisik_peaking, fisik_recovery, fisik_cedera,
                        taktis_lawan, taktis_instruksi, taktis_ujicoba,
                        mental_cemas, mental_fokus, mental_rutinitas, mental_psikolog,
                        nutrisi_bb, nutrisi_asupan, nutrisi_hidrasi, nutrisi_tidur,
                        medis_rekam, medis_doping, medis_alat, medis_nonteknis,
                        status, catatan_admin
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        tanggal, cabor, lokasi.strip(), st.session_state["username"],
                        fisik_1, fisik_2, fisik_3, fisik_4,
                        taktis_1, taktis_2, taktis_3,
                        mental_1, mental_2, mental_3, mental_4,
                        nutrisi_1, nutrisi_2, nutrisi_3, nutrisi_4,
                        medis_1, medis_2, medis_3, medis_4,
                        "Belum Ditindaklanjuti", ""
                    )
                )
                conn.commit()
                conn.close()
                st.success("✅ Laporan berhasil disimpan. Admin dapat melihat dan menindaklanjuti laporan.")


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
