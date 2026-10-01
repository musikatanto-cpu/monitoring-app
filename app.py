"""
Sistem Monitoring BINPRES
KONI Kabupaten Tangerang

Aplikasi Streamlit untuk monitoring & evaluasi cabang olahraga.
"""

import streamlit as st
import datetime
import sqlite3
import hashlib
from io import BytesIO
from typing import Optional, List, Dict, Any

import pandas as pd
from docx import Document
from docx.shared import Inches

# ============================================================
# KONFIGURASI
# ============================================================
st.set_page_config(
    page_title="Sistem Monitoring Binpres",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded",
)

DB_NAME = "monitoring.db"

DEFAULT_STATUS = "Belum Ditindaklanjuti"
STATUS_OPTIONS = [
    "Belum Ditindaklanjuti",
    "Sedang Ditindaklanjuti",
    "Selesai",
]

DEFAULT_CABOR = [
    "ANGGAR", "ANGKAT BERAT", "ANGKAT BESI", "AQUATIC/RENANG", "ARUNG JERAM",
    "ATLETIK", "BALAP SEPEDA", "BARONGSAI", "BERMOTOR", "BILLIARD", "BINARAGA",
    "BOLA BASKET", "BOLA TANGAN", "BOLA VOLI", "BOWLING", "BRIDGE", "BULUTANGKIS",
    "CATUR", "DAYUNG", "DRUMBAND", "E-SPORT", "FLOOR BALL", "FUTSAL", "GATEBALL",
    "GOLF", "GULAT", "GYMNASTIC/SENAM", "HOKI", "IBCA MMA", "JU JITSU", "JUDO",
    "KARATE", "KEMPO", "MENEMBAK", "MUAYTHAI", "PANAHAN", "PANJAT TEBING",
    "PENCAK SILAT", "PETANQUE", "PICKLEBALL", "RUGBY", "SAMBO", "SELAM",
    "SEPAK BOLA", "SEPAK TAKRAW", "SEPATU RODA", "SOFTBALL", "SQUASH",
    "TAEKWONDO", "TARUNG DERAJAT", "TENIS LAPANG", "TENIS MEJA", "TINJU",
    "WOODBALL", "WUSHU",
]

# ============================================================
# DATABASE LAYER
# ============================================================
def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def verify_password(password: str, stored: str) -> bool:
    """Mendukung hash baru dan plaintext lama (migrasi)."""
    return hash_password(password) == stored or password == stored


def _add_column_if_missing(cursor: sqlite3.Cursor, table: str, column: str, definition: str) -> None:
    existing = [r["name"] for r in cursor.execute(f"PRAGMA table_info({table})").fetchall()]
    if column not in existing:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def init_db() -> None:
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

    # Migrasi kolom (kompatibel database lama)
    for col, defn in {
        "nama_lengkap": "TEXT DEFAULT ''",
        "aktif": "INTEGER DEFAULT 1",
        "created_at": "TEXT DEFAULT CURRENT_TIMESTAMP",
    }.items():
        _add_column_if_missing(cursor, "users", col, defn)

    for col, defn in {
        "status": "TEXT DEFAULT 'Belum Ditindaklanjuti'",
        "catatan_admin": "TEXT DEFAULT ''",
        "updated_at": "TEXT DEFAULT CURRENT_TIMESTAMP",
    }.items():
        _add_column_if_missing(cursor, "laporan_monitoring", col, defn)

    # Seed default users
    count = cursor.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]
    if count == 0:
        cursor.executemany(
            """INSERT INTO users (username, password, role, nama_lengkap, aktif)
               VALUES (?, ?, ?, ?, 1)""",
            [
                ("admin", hash_password("admin123"), "admin", "Administrator"),
                ("petugas", hash_password("petugas123"), "user", "Petugas Monitoring"),
            ],
        )
    else:
        # Upgrade password plaintext lama → hash
        for username, old_pw, role, nama in [
            ("admin", "admin123", "admin", "Administrator"),
            ("petugas", "petugas123", "user", "Petugas Monitoring"),
        ]:
            row = cursor.execute(
                "SELECT id, password FROM users WHERE username = ?", (username,)
            ).fetchone()
            if row and row["password"] == old_pw:
                cursor.execute(
                    "UPDATE users SET password = ?, role = ?, nama_lengkap = ? WHERE id = ?",
                    (hash_password(old_pw), role, nama, row["id"]),
                )

    # Seed cabor
    if cursor.execute("SELECT COUNT(*) AS n FROM cabor").fetchone()["n"] == 0:
        cursor.executemany(
            "INSERT OR IGNORE INTO cabor (nama) VALUES (?)",
            [(nama,) for nama in DEFAULT_CABOR],
        )

    conn.commit()
    conn.close()


# ============================================================
# HELPER / QUERY
# ============================================================
def logout() -> None:
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()


def get_cabor_list() -> List[str]:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT nama FROM cabor WHERE aktif = 1 ORDER BY nama"
        ).fetchall()
    return [r["nama"] for r in rows]


def get_current_user() -> Optional[sqlite3.Row]:
    username = st.session_state.get("username")
    if not username:
        return None
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM users WHERE username = ?", (username,)
        ).fetchone()


def fetch_all_users() -> pd.DataFrame:
    with get_conn() as conn:
        return pd.read_sql_query(
            """SELECT id, username, nama_lengkap, role,
                      CASE WHEN aktif = 1 THEN 'Aktif' ELSE 'Nonaktif' END AS status,
                      created_at
               FROM users ORDER BY id DESC""",
            conn,
        )


def fetch_all_cabor() -> pd.DataFrame:
    with get_conn() as conn:
        return pd.read_sql_query(
            """SELECT id, nama,
                      CASE WHEN aktif = 1 THEN 'Aktif' ELSE 'Nonaktif' END AS status
               FROM cabor ORDER BY nama""",
            conn,
        )


def fetch_laporan_summary() -> Dict[str, int]:
    with get_conn() as conn:
        total = conn.execute("SELECT COUNT(*) AS n FROM laporan_monitoring").fetchone()["n"]
        bulan_ini = conn.execute(
            """SELECT COUNT(*) AS n FROM laporan_monitoring
               WHERE strftime('%Y-%m', tanggal) = strftime('%Y-%m', 'now')"""
        ).fetchone()["n"]
        cabor_termonitor = conn.execute(
            "SELECT COUNT(DISTINCT cabor) AS n FROM laporan_monitoring"
        ).fetchone()["n"]
        perlu_tindak = conn.execute(
            """SELECT COUNT(*) AS n FROM laporan_monitoring
               WHERE status = ? OR status IS NULL""",
            (DEFAULT_STATUS,),
        ).fetchone()["n"]
    return {
        "total": total,
        "bulan_ini": bulan_ini,
        "cabor_termonitor": cabor_termonitor,
        "perlu_tindak": perlu_tindak,
    }


def fetch_laporan_list() -> pd.DataFrame:
    with get_conn() as conn:
        return pd.read_sql_query(
            """SELECT id, tanggal, cabor, lokasi, petugas, status
               FROM laporan_monitoring
               ORDER BY tanggal DESC, id DESC""",
            conn,
        )


def get_laporan_by_id(laporan_id: int) -> Optional[sqlite3.Row]:
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM laporan_monitoring WHERE id = ?", (laporan_id,)
        ).fetchone()


# ============================================================
# WORD REPORT GENERATOR
# ============================================================
def generate_word_report(data_list: List[Dict[str, Any]], is_all: bool = False) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    judul = "REKAPITULASI SEMUA LAPORAN MONEV BINPRES" if is_all else "LAPORAN MONEV BINPRES"
    title = doc.add_heading(judul, 0)
    title.alignment = 1  # center

    def add_section(title_text: str, items: List[tuple]) -> None:
        doc.add_heading(title_text, level=2)
        for question, answer in items:
            p = doc.add_paragraph()
            p.add_run(question).bold = True
            doc.add_paragraph(answer if answer else "-")

    for idx, data in enumerate(data_list):
        if is_all:
            doc.add_heading(
                f"Laporan {idx + 1}: {data['cabor']} - {data['tanggal']}", level=1
            )

        doc.add_paragraph(f"Cabang Olahraga\t: {data['cabor']}")
        doc.add_paragraph(f"Tanggal\t\t: {data['tanggal']}")
        doc.add_paragraph(f"Lokasi\t\t: {data['lokasi']}")
        doc.add_paragraph(f"Petugas Monev\t: {data['petugas']}")
        doc.add_paragraph(f"Status\t\t: {data.get('status', '-')}")
        doc.add_paragraph("-" * 70)

        add_section("1. Performa Fisik & Kebugaran", [
            ("Capaian parameter fisik:", data.get("fisik_parameter")),
            ("Grafik performa puncak (peaking):", data.get("fisik_peaking")),
            ("Tingkat pemulihan fisik (recovery):", data.get("fisik_recovery")),
            ("Keluhan cedera lama/baru:", data.get("fisik_cedera")),
        ])
        add_section("2. Kesiapan Taktis & Penguasaan Strategi", [
            ("Pemetaan kekuatan calon lawan:", data.get("taktis_lawan")),
            ("Kemampuan mengikuti instruksi teknis:", data.get("taktis_instruksi")),
            ("Hasil try-out / sparing:", data.get("taktis_ujicoba")),
        ])
        add_section("3. Mental, Psikologis & Kesiapan Mental", [
            ("Tingkat kecemasan & pengendalian stres:", data.get("mental_cemas")),
            ("Fokus, motivasi, dan self-confidence:", data.get("mental_fokus")),
            ("Rutinitas mental khusus:", data.get("mental_rutinitas")),
            ("Koordinasi dengan psikolog olahraga:", data.get("mental_psikolog")),
        ])
        add_section("4. Nutrisi, Berat Badan & Gaya Hidup", [
            ("Progres penyesuaian berat badan:", data.get("nutrisi_bb")),
            ("Asupan nutrisi dan suplemen:", data.get("nutrisi_asupan")),
            ("Status hidrasi:", data.get("nutrisi_hidrasi")),
            ("Kualitas dan kecukupan tidur:", data.get("nutrisi_tidur")),
        ])
        add_section("5. Medis, Bebas Doping & Logistik", [
            ("Status rekam medis & tim medis:", data.get("medis_rekam")),
            ("Keamanan obat/suplemen:", data.get("medis_doping")),
            ("Kesiapan perlengkapan tanding:", data.get("medis_alat")),
            ("Kendala non-teknis:", data.get("medis_nonteknis")),
        ])

        if data.get("catatan_admin"):
            add_section("Catatan Admin / Tindak Lanjut", [
                ("Catatan:", data["catatan_admin"]),
            ])

        if idx < len(data_list) - 1:
            doc.add_page_break()

    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# ============================================================
# UI COMPONENTS
# ============================================================
def render_sidebar() -> None:
    user = get_current_user()
    with st.sidebar:
        st.markdown("## 🏆 BINPRES")
        st.caption("Monitoring & Evaluasi Cabor")
        st.markdown("---")
        nama = (user["nama_lengkap"] if user and user["nama_lengkap"]
                else st.session_state.get("username", ""))
        st.write(f"👤 **{nama}**")
        st.caption(f"Role: {st.session_state.get('role', '').upper()}")
        st.markdown("---")
        if st.button("🚪 Logout", use_container_width=True):
            logout()


def halaman_login() -> None:
    st.markdown(
        """
        <div style="text-align:center;padding:30px 0 15px 0">
            <div style="font-size:52px">🏆</div>
            <h1 style="margin-bottom:5px">Sistem Monitoring BINPRES</h1>
            <p style="color:#777">KONI Kabupaten Tangerang</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.form("login_form"):
        username = st.text_input("👤 Username")
        password = st.text_input("🔒 Password", type="password")
        submit = st.form_submit_button(
            "Masuk ke Sistem", use_container_width=True, type="primary"
        )

        if submit:
            with get_conn() as conn:
                user = conn.execute(
                    "SELECT * FROM users WHERE username = ? AND aktif = 1",
                    (username.strip(),),
                ).fetchone()

            if user and verify_password(password, user["password"]):
                st.session_state["logged_in"] = True
                st.session_state["username"] = user["username"]
                st.session_state["role"] = user["role"]
                st.session_state["nama_lengkap"] = user["nama_lengkap"] or user["username"]
                st.rerun()
            else:
                st.error("Username atau Password salah / akun tidak aktif.")

    st.caption("Versi pengembangan • Data tersimpan pada database SQLite lokal")


def dashboard_admin() -> None:
    stats = fetch_laporan_summary()

    st.markdown("## 📊 Dashboard Monitoring")
    st.caption("Ringkasan aktivitas monitoring BINPRES")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Laporan", stats["total"])
    c2.metric("Laporan Bulan Ini", stats["bulan_ini"])
    c3.metric("Cabor Termonitor", stats["cabor_termonitor"])
    c4.metric("Perlu Tindak Lanjut", stats["perlu_tindak"])
    st.markdown("---")


# ------------------------------------------------------------
# MANAJEMEN USER
# ------------------------------------------------------------
def kelola_user() -> None:
    st.subheader("👥 Manajemen User")
    st.dataframe(fetch_all_users(), use_container_width=True, hide_index=True)

    st.markdown("### ➕ Tambah User")
    with st.form("tambah_user"):
        col1, col2 = st.columns(2)
        with col1:
            username = st.text_input("Username baru")
            nama = st.text_input("Nama lengkap")
        with col2:
            password = st.text_input("Password", type="password")
            role = st.selectbox("Hak akses", ["user", "admin"])

        if st.form_submit_button("Simpan User", use_container_width=True, type="primary"):
            if not username.strip() or not password:
                st.error("Username dan password wajib diisi.")
            elif len(password) < 6:
                st.error("Password minimal 6 karakter.")
            else:
                try:
                    with get_conn() as conn:
                        conn.execute(
                            """INSERT INTO users (username, password, role, nama_lengkap, aktif)
                               VALUES (?, ?, ?, ?, 1)""",
                            (username.strip(), hash_password(password), role, nama.strip()),
                        )
                        conn.commit()
                    st.success("User berhasil ditambahkan.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Username sudah digunakan.")

    st.markdown("### ✏️ Edit / Nonaktifkan User")
    with get_conn() as conn:
        users = [dict(r) for r in conn.execute(
            "SELECT id, username, nama_lengkap, role, aktif FROM users ORDER BY username"
        ).fetchall()]

    if not users:
        return

    pilihan = st.selectbox(
        "Pilih user",
        users,
        format_func=lambda x: f"{x['username']} — {x['nama_lengkap'] or '-'}",
    )

    with st.form("edit_user"):
        nama_baru = st.text_input("Nama lengkap", value=pilihan["nama_lengkap"] or "")
        role_baru = st.selectbox(
            "Role", ["user", "admin"],
            index=0 if pilihan["role"] == "user" else 1,
        )
        aktif_baru = st.checkbox("Akun aktif", value=bool(pilihan["aktif"]))
        password_baru = st.text_input(
            "Password baru (kosongkan jika tidak diubah)", type="password"
        )

        if st.form_submit_button("💾 Simpan Perubahan", use_container_width=True):
            with get_conn() as conn:
                if password_baru:
                    conn.execute(
                        """UPDATE users
                           SET nama_lengkap = ?, role = ?, aktif = ?, password = ?
                           WHERE id = ?""",
                        (nama_baru.strip(), role_baru, int(aktif_baru),
                         hash_password(password_baru), pilihan["id"]),
                    )
                else:
                    conn.execute(
                        """UPDATE users
                           SET nama_lengkap = ?, role = ?, aktif = ?
                           WHERE id = ?""",
                        (nama_baru.strip(), role_baru, int(aktif_baru), pilihan["id"]),
                    )
                conn.commit()
            st.success("Data user berhasil diperbarui.")
            st.rerun()


# ------------------------------------------------------------
# MANAJEMEN CABOR
# ------------------------------------------------------------
def kelola_cabor() -> None:
    st.subheader("🏅 Kelola Cabang Olahraga")
    st.dataframe(fetch_all_cabor(), use_container_width=True, hide_index=True)

    with st.form("tambah_cabor"):
        nama_cabor = st.text_input("Nama Cabang Olahraga Baru")
        if st.form_submit_button("➕ Tambah Cabor", use_container_width=True):
            if not nama_cabor.strip():
                st.error("Nama cabor wajib diisi.")
            else:
                try:
                    with get_conn() as conn:
                        conn.execute(
                            "INSERT INTO cabor (nama) VALUES (?)",
                            (nama_cabor.strip().upper(),),
                        )
                        conn.commit()
                    st.success("Cabor berhasil ditambahkan.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Cabor tersebut sudah ada.")

    st.markdown("### ✏️ Edit / Nonaktifkan Cabor")
    with get_conn() as conn:
        rows = [dict(r) for r in conn.execute("SELECT * FROM cabor ORDER BY nama").fetchall()]

    if not rows:
        return

    pilihan = st.selectbox("Pilih cabor", rows, format_func=lambda x: x["nama"])
    with st.form("edit_cabor"):
        nama_baru = st.text_input("Nama cabor", value=pilihan["nama"])
        aktif = st.checkbox("Cabor aktif", value=bool(pilihan["aktif"]))
        if st.form_submit_button("💾 Simpan", use_container_width=True):
            try:
                with get_conn() as conn:
                    conn.execute(
                        "UPDATE cabor SET nama = ?, aktif = ? WHERE id = ?",
                        (nama_baru.strip().upper(), int(aktif), pilihan["id"]),
                    )
                    conn.commit()
                st.success("Cabor berhasil diperbarui.")
                st.rerun()
            except sqlite3.IntegrityError:
                st.error("Nama cabor sudah digunakan.")


# ------------------------------------------------------------
# LAPORAN ADMIN
# ------------------------------------------------------------
def halaman_laporan_admin() -> None:
    st.subheader("📄 Data Laporan Monitoring")

    df = fetch_laporan_list()
    if df.empty:
        st.info("Belum ada laporan masuk.")
        return

    # Filter
    f1, f2, f3 = st.columns(3)
    with f1:
        cabor_filter = st.selectbox(
            "Filter Cabor",
            ["Semua"] + sorted(df["cabor"].dropna().unique().tolist()),
        )
    with f2:
        status_filter = st.selectbox(
            "Filter Status",
            ["Semua"] + sorted(df["status"].fillna(DEFAULT_STATUS).unique().tolist()),
        )
    with f3:
        keyword = st.text_input("🔎 Cari lokasi / petugas")

    filtered = df.copy()
    if cabor_filter != "Semua":
        filtered = filtered[filtered["cabor"] == cabor_filter]
    if status_filter != "Semua":
        filtered = filtered[filtered["status"].fillna(DEFAULT_STATUS) == status_filter]
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

    # Export
    st.markdown("---")
    st.markdown("### 📥 Export Laporan")
    c1, c2 = st.columns(2)

    with c1:
        pilihan_id = st.selectbox("Pilih ID laporan untuk Word", filtered["id"].tolist())
        if st.button("📄 Generate Word Satuan", use_container_width=True):
            row = get_laporan_by_id(int(pilihan_id))
            if row:
                word_file = generate_word_report([dict(row)], is_all=False)
                st.download_button(
                    "⬇️ Download Word",
                    data=word_file,
                    file_name=f"Laporan_{row['cabor']}_{row['tanggal']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

    with c2:
        st.write("**Rekap hasil filter**")
        if st.button("📚 Generate Word Rekap", use_container_width=True):
            ids = tuple(int(x) for x in filtered["id"].tolist())
            placeholders = ",".join("?" * len(ids))
            with get_conn() as conn:
                rows = conn.execute(
                    f"""SELECT * FROM laporan_monitoring
                        WHERE id IN ({placeholders})
                        ORDER BY tanggal ASC, id ASC""",
                    ids,
                ).fetchall()
            word_file = generate_word_report([dict(r) for r in rows], is_all=True)
            st.download_button(
                "⬇️ Download Rekap Word",
                data=word_file,
                file_name=f"Rekap_Monev_{datetime.date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

    # Tindak lanjut
    st.markdown("---")
    st.markdown("### ✏️ Tindak Lanjut Laporan")

    pilihan_edit = st.selectbox(
        "Pilih laporan yang akan diperbarui",
        filtered["id"].tolist(),
        key="edit_laporan_id",
    )
    row = get_laporan_by_id(int(pilihan_edit))

    if row:
        with st.form("update_status_laporan"):
            current_status = row["status"] or DEFAULT_STATUS
            status = st.selectbox(
                "Status tindak lanjut",
                STATUS_OPTIONS,
                index=STATUS_OPTIONS.index(current_status)
                if current_status in STATUS_OPTIONS else 0,
            )
            catatan = st.text_area(
                "Catatan Admin / Rekomendasi",
                value=row["catatan_admin"] or "",
            )
            if st.form_submit_button("💾 Simpan Tindak Lanjut", use_container_width=True):
                with get_conn() as conn:
                    conn.execute(
                        """UPDATE laporan_monitoring
                           SET status = ?, catatan_admin = ?, updated_at = CURRENT_TIMESTAMP
                           WHERE id = ?""",
                        (status, catatan, int(pilihan_edit)),
                    )
                    conn.commit()
                st.success("Tindak lanjut berhasil diperbarui.")
                st.rerun()


# ------------------------------------------------------------
# HALAMAN ADMIN
# ------------------------------------------------------------
def halaman_admin() -> None:
    render_sidebar()
    dashboard_admin()

    tab1, tab2, tab3 = st.tabs([
        "📄 Laporan Monitoring",
        "👥 Manajemen User",
        "🏅 Kelola Cabor",
    ])
    with tab1:
        halaman_laporan_admin()
    with tab2:
        kelola_user()
    with tab3:
        kelola_cabor()


# ------------------------------------------------------------
# HALAMAN PETUGAS / USER
# ------------------------------------------------------------
def halaman_user() -> None:
    render_sidebar()

    st.markdown("## 📝 Form Input Monitoring")
    st.caption(
        f"Petugas: **{st.session_state.get('nama_lengkap', st.session_state['username'])}**"
    )

    cabor_list = get_cabor_list()

    with st.form("form_monitoring", clear_on_submit=True):
        st.markdown("### 📌 Informasi Dasar")
        col1, col2 = st.columns(2)
        with col1:
            tanggal = st.date_input("Tanggal Monitoring", datetime.date.today())
            cabor = st.selectbox("Cabang Olahraga", ["Pilih Cabor..."] + cabor_list)
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
            type="primary",
        )

        if submitted:
            if cabor == "Pilih Cabor...":
                st.error("⚠️ Harap pilih Cabang Olahraga.")
            elif not lokasi.strip():
                st.error("⚠️ Lokasi wajib diisi.")
            else:
                with get_conn() as conn:
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
                            DEFAULT_STATUS, "",
                        ),
                    )
                    conn.commit()
                st.success("✅ Laporan berhasil disimpan. Admin dapat melihat dan menindaklanjuti laporan.")


# ============================================================
# ENTRY POINT
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
