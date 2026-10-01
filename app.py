import streamlit as st
import datetime
import sqlite3

# --- KONFIGURASI DATABASE ---
def init_db():
    conn = sqlite3.connect('monitoring_koni.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS laporan_monitoring (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal DATE NOT NULL,
            cabor TEXT NOT NULL,
            lokasi TEXT NOT NULL,
            
            -- Fisik & Kebugaran
            fisik_parameter TEXT,
            fisik_peaking TEXT,
            fisik_recovery TEXT,
            fisik_cedera TEXT,
            
            -- Taktis & Strategi
            taktis_lawan TEXT,
            taktis_instruksi TEXT,
            taktis_ujicoba TEXT,
            
            -- Mental & Psikologis
            mental_cemas TEXT,
            mental_fokus TEXT,
            mental_rutinitas TEXT,
            mental_psikolog TEXT,
            
            -- Nutrisi & Gaya Hidup
            nutrisi_bb TEXT,
            nutrisi_asupan TEXT,
            nutrisi_hidrasi TEXT,
            nutrisi_tidur TEXT,
            
            -- Medis & Logistik
            medis_rekam TEXT,
            medis_doping TEXT,
            medis_alat TEXT,
            medis_nonteknis TEXT,
            
            foto_path TEXT
        )
    ''')
    conn.commit()
    conn.close()

# --- HALAMAN INPUT MONITORING ---
def halaman_input_monitoring():
    st.header("📝 Form Laporan Monitoring Atlet & Cabor")
    st.caption("Silakan pilih tanggal dan lengkapi indikator evaluasi kesiapan atlet di bawah ini.")

    with st.form("form_monitoring", clear_on_submit=True):
        
        # 1. INFORMASI DASAR (User milih sendiri)
        st.subheader("Informasi Dasar")
        col1, col2 = st.columns(2)
        with col1:
            tanggal = st.date_input("Tanggal Monitoring", datetime.date.today())
            cabor = st.selectbox("Cabang Olahraga", ["Pilih Cabor...", "Panahan", "Taekwondo", "Sepatu Roda", "Judo", "Sepak Takraw", "Catur"])
        with col2:
            lokasi = st.text_input("Lokasi Latihan / Try-out / Sparing")
            foto = st.file_uploader("Upload Dokumentasi", type=['jpg', 'png', 'jpeg'])

        st.markdown("---")
        st.subheader("Indikator Evaluasi")
        st.caption("Isi catatan pada setiap kategori di bawah ini sesuai hasil pantauan di lapangan.")

        # 2. PERFORMA FISIK & KEBUGARAN
        with st.expander("💪 1. Performa Fisik & Kebugaran"):
            fisik_parameter = st.text_area("Capaian parameter fisik (kekuatan, daya tahan, kecepatan, kelincahan) vs benchmark target (Catatan dari tim SC):", height=68)
            fisik_peaking = st.text_area("Apakah atlet mencapai grafik performa puncak (peaking) sesuai timeline?", height=68)
            fisik_recovery = st.text_area("Bagaimana tingkat pemulihan fisik (recovery) harian atlet pasca-latihan intensitas tinggi?", height=68)
            fisik_cedera = st.text_area("Apakah ada keluhan cedera lama yang kambuh / indikasi cedera baru?", height=68)

        # 3. KESIAPAN TAKTIS & STRATEGI
        with st.expander("🎯 2. Kesiapan Taktis & Penguasaan Strategi"):
            taktis_lawan = st.text_area("Bagaimana gambaran dan pemetaan kekuatan calon lawan?", height=68)
            taktis_instruksi = st.text_area("Kemampuan atlet mengikuti instruksi teknis pelatih di bawah kondisi tekanan (pressure):", height=68)
            taktis_ujicoba = st.text_area("Hasil try-out / sparing: Apakah menunjukkan peningkatan efektivitas skema permainan?", height=68)

        # 4. MENTAL & PSIKOLOGIS
        with st.expander("🧠 3. Mental, Psikologis & Kesiapan Mental"):
            mental_cemas = st.text_area("Tingkat kecemasan (anxiety) dan kemampuan mengendalikan stres jelang pertandingan:", height=68)
            mental_fokus = st.text_area("Tingkat fokus, motivasi, dan kepercayaan diri (self-confidence) saat latihan/simulasi:", height=68)
            mental_rutinitas = st.text_area("Apakah atlet memiliki rutinitas mental khusus (mental routine) saat masuk lapangan?", height=68)
            mental_psikolog = st.text_area("Bagaimana koordinasi dengan tim psikolog olahraga terkait beban target?", height=68)

        # 5. NUTRISI & GAYA HIDUP
        with st.expander("🥗 4. Nutrisi, Berat Badan & Gaya Hidup"):
            nutrisi_bb = st.text_area("Progres penyesuaian berat badan (weight management) tanpa mengorbankan kondisi fisik:", height=68)
            nutrisi_asupan = st.text_area("Pemantauan asupan nutrisi dan suplemen harian sesuai fase latihan:", height=68)
            nutrisi_hidrasi = st.text_area("Status hidrasi harian atlet:", height=68)
            nutrisi_tidur = st.text_area("Kualitas dan kecukupan waktu tidur/istirahat atlet setiap hari:", height=68)

        # 6. MEDIS & LOGISTIK
        with st.expander("⚕️ 5. Medis, Bebas Doping & Logistik"):
            medis_rekam = st.text_area("Status rekam medis terkini & kesiapan tim medis/fisioterapis:", height=68)
            medis_doping = st.text_area("Keamanan obat, suplemen, dan terapi (Bebas Doping):", height=68)
            medis_alat = st.text_area("Kesiapan perlengkapan khusus bertanding (sepatu, pakaian, alat tanding pribadi):", height=68)
            medis_nonteknis = st.text_area("Kendala non-teknis (akomodasi, transportasi, administrasi) yang berpotensi mengganggu konsentrasi:", height=68)

        # Tombol Submit
        submitted = st.form_submit_button("💾 Simpan Laporan Monitoring", use_container_width=True)
        
        if submitted:
            if cabor == "Pilih Cabor...":
                st.error("⚠️ Harap pilih Cabang Olahraga terlebih dahulu!")
            elif not lokasi:
                st.error("⚠️ Lokasi wajib diisi!")
            else:
                # Proses simpan ke database
                try:
                    conn = sqlite3.connect('monitoring_koni.db')
                    cursor = conn.cursor()
                    cursor.execute('''
                        INSERT INTO laporan_monitoring (
                            tanggal, cabor, lokasi,
                            fisik_parameter, fisik_peaking, fisik_recovery, fisik_cedera,
                            taktis_lawan, taktis_instruksi, taktis_ujicoba,
                            mental_cemas, mental_fokus, mental_rutinitas, mental_psikolog,
                            nutrisi_bb, nutrisi_asupan, nutrisi_hidrasi, nutrisi_tidur,
                            medis_rekam, medis_doping, medis_alat, medis_nonteknis
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        tanggal, cabor, lokasi,
                        fisik_parameter, fisik_peaking, fisik_recovery, fisik_cedera,
                        taktis_lawan, taktis_instruksi, taktis_ujicoba,
                        mental_cemas, mental_fokus, mental_rutinitas, mental_psikolog,
                        nutrisi_bb, nutrisi_asupan, nutrisi_hidrasi, nutrisi_tidur,
                        medis_rekam, medis_doping, medis_alat, medis_nonteknis
                    ))
                    conn.commit()
                    conn.close()
                    
                    st.success(f"✅ Laporan monitoring cabang **{cabor}** untuk tanggal **{tanggal.strftime('%d %B %Y')}** berhasil disimpan!")
                except Exception as e:
                    st.error(f"Terjadi kesalahan saat menyimpan data: {e}")

# Inisialisasi DB sebelum menjalankan aplikasi
init_db()

# Tampilkan UI
halaman_input_monitoring()
