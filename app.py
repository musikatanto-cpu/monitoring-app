

from __future__ import annotations

import streamlit as st
import datetime
import hashlib
import os
import uuid
import re
from io import BytesIO
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple

import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image, ImageDraw, ImageFont
from supabase import create_client, Client


# ============================================================
# KONFIGURASI
# ============================================================
st.set_page_config(
    page_title="Sistem Monitoring BINPRES",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# KONFIGURASI FOTO & WORD
# ============================================================
WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MAX_PHOTO_MB = 10
MAX_PHOTOS_PER_REPORT = 10
PHOTO_TIMESTAMP_FORMAT = "%d-%m-%Y %H:%M:%S"

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

FIELD_LABELS: Dict[str, List[Tuple[str, str]]] = {
    "1. Performa Fisik & Kebugaran": [
        ("fisik_parameter", "Capaian parameter fisik (vs benchmark target)"),
        ("fisik_peaking", "Grafik performa puncak (peaking)"),
        ("fisik_recovery", "Tingkat pemulihan fisik (recovery)"),
        ("fisik_cedera", "Keluhan cedera lama / indikasi cedera baru"),
    ],
    "2. Kesiapan Taktis & Strategi": [
        ("taktis_lawan", "Pemetaan kekuatan calon lawan"),
        ("taktis_instruksi", "Kemampuan mengikuti instruksi teknis"),
        ("taktis_ujicoba", "Hasil try-out / sparing"),
    ],
    "3. Mental, Psikologis & Kesiapan Mental": [
        ("mental_cemas", "Tingkat kecemasan & pengendalian stres"),
        ("mental_fokus", "Fokus, motivasi, dan self-confidence"),
        ("mental_rutinitas", "Rutinitas mental khusus"),
        ("mental_psikolog", "Koordinasi dengan psikolog olahraga"),
    ],
    "4. Nutrisi, Berat Badan & Gaya Hidup": [
        ("nutrisi_bb", "Progres penyesuaian berat badan"),
        ("nutrisi_asupan", "Asupan nutrisi dan suplemen"),
        ("nutrisi_hidrasi", "Status hidrasi"),
        ("nutrisi_tidur", "Kualitas dan kecukupan tidur"),
    ],
    "5. Medis, Bebas Doping & Logistik": [
        ("medis_rekam", "Status rekam medis & tim medis"),
        ("medis_doping", "Keamanan obat / suplemen (bebas doping)"),
        ("medis_alat", "Kesiapan perlengkapan tanding"),
        ("medis_nonteknis", "Kendala non-teknis"),
    ],
}

CUSTOM_CSS = """
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        color: white;
        padding: 1.2rem 1.5rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 12px rgba(30, 58, 138, 0.25);
    }
    .main-header h1 { margin: 0; font-size: 1.6rem; font-weight: 700; }
    .main-header p { margin: 0.3rem 0 0 0; opacity: 0.9; font-size: 0.95rem; }

    div[data-testid="stMetric"] {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 0.8rem 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    div[data-testid="stMetric"] label { color: #64748b !important; font-size: 0.85rem !important; }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] { color: #1e3a8a !important; font-weight: 700; }

    .section-badge {
        display: inline-block;
        background: #1e3a8a;
        color: white;
        padding: 0.35rem 0.9rem;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.95rem;
        margin-bottom: 0.8rem;
    }

    .badge-belum { background:#fef3c7; color:#92400e; padding:3px 10px; border-radius:12px; font-size:0.8rem; font-weight:600; }
    .badge-sedang { background:#dbeafe; color:#1e40af; padding:3px 10px; border-radius:12px; font-size:0.8rem; font-weight:600; }
    .badge-selesai { background:#d1fae5; color:#065f46; padding:3px 10px; border-radius:12px; font-size:0.8rem; font-weight:600; }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
    }
    section[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
    section[data-testid="stSidebar"] .stButton > button {
        background: #334155; border: 1px solid #475569; color: white;
    }
    section[data-testid="stSidebar"] .stButton > button:hover {
        background: #475569; border-color: #64748b;
    }

    .stButton > button[kind="primary"],
    div[data-testid="stFormSubmitButton"] > button {
        background: linear-gradient(135deg, #dc2626 0%, #ef4444 100%) !important;
        border: none !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
    }

    div[data-testid="stExpander"] {
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        margin-bottom: 0.5rem;
    }

    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
</style>
"""


# ============================================================
# SUPABASE CONFIGURATION & BACKEND
# ============================================================
SUPABASE_BUCKET = "monitoring-binpres"
SUPABASE_URL_KEY = "SUPABASE_URL"
SUPABASE_SERVICE_KEY = "SUPABASE_SECRET_KEY"


def _get_secret(name: str) -> str:
    try:
        value = st.secrets.get(name)
    except Exception:
        value = None
    if value:
        return str(value).strip()
    return os.getenv(name, "").strip()


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def verify_password(password: str, stored: str) -> bool:
    return hash_password(password) == stored or password == stored


@st.cache_resource(show_spinner=False)
def get_supabase() -> Client:
    """Client Supabase server-side. Jangan pernah menampilkan service-role key ke user."""
    url = _get_secret(SUPABASE_URL_KEY)
    key = _get_secret(SUPABASE_SERVICE_KEY) or _get_secret("SUPABASE_SERVICE_ROLE_KEY") or _get_secret("SUPABASE_KEY")
    if not url or not key:
        raise RuntimeError(
            "SUPABASE_URL dan SUPABASE_SECRET_KEY belum diisi di Streamlit Secrets."
        )
    return create_client(url, key)


def _sb_data(response) -> list:
    data = getattr(response, "data", None)
    if data is None:
        return []
    return data if isinstance(data, list) else [data]


def safe_filename(value: str) -> str:
    value = str(value or "").strip()
    value = re.sub(r'[\\/:*?"<>|]+', "-", value)
    value = re.sub(r"\s+", " ", value)
    return value[:180] or "Laporan_Monitoring"


def _now_iso() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def ensure_storage_bucket() -> None:
    """Buat bucket private jika belum ada. Gunakan service-role key di server."""
    sb = get_supabase()
    try:
        buckets = sb.storage.list_buckets() or []
        names = set()
        for bucket in buckets:
            if isinstance(bucket, dict):
                names.add(bucket.get("id"))
                names.add(bucket.get("name"))
            else:
                names.add(getattr(bucket, "id", None))
                names.add(getattr(bucket, "name", None))
        if SUPABASE_BUCKET not in names:
            sb.storage.create_bucket(
                SUPABASE_BUCKET,
                options={
                    "public": False,
                    "allowed_mime_types": ["image/jpeg", "image/png", "image/webp"],
                    "file_size_limit": MAX_PHOTO_MB * 1024 * 1024,
                },
            )
    except Exception as exc:
        st.warning(
            f"Bucket Storage '{SUPABASE_BUCKET}' belum dapat dibuat/diperiksa otomatis. "
            f"Pastikan bucket tersebut tersedia di Supabase. Detail: {exc}"
        )


def init_backend() -> None:
    """Validasi struktur tabel, seed akun utama dan cabor."""
    try:
        sb = get_supabase()
        # Validasi tabel utama.
        for table in ("users", "cabor", "laporan_monitoring", "laporan_foto"):
            sb.table(table).select("id").limit(1).execute()

        # Akun wajib: userkoni / koni123.
        found = _sb_data(
            sb.table("users").select("id").eq("username", "userkoni").limit(1).execute()
        )
        payload = {
            "username": "userkoni",
            "password": hash_password("koni123"),
            "role": "admin",
            "nama_lengkap": "User KONI",
            "aktif": True,
        }
        if found:
            sb.table("users").update(payload).eq("id", int(found[0]["id"])).execute()
        else:
            sb.table("users").insert(payload).execute()

        # Seed cabang olahraga bila kosong.
        cabor_exists = _sb_data(sb.table("cabor").select("id").limit(1).execute())
        if not cabor_exists:
            sb.table("cabor").insert(
                [{"nama": n, "aktif": True} for n in DEFAULT_CABOR]
            ).execute()

        ensure_storage_bucket()
    except Exception as exc:
        st.error("Supabase belum siap digunakan.")
        st.code(str(exc))
        st.info(
            "1) Jalankan supabase_schema.sql di Supabase SQL Editor. "
            "2) Isi SUPABASE_URL dan SUPABASE_SECRET_KEY di Streamlit Secrets. "
            "3) Pastikan bucket monitoring-binpres tersedia."
        )
        st.stop()


def logout() -> None:
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# ============================================================
# HELPER / QUERY
# ============================================================
def get_cabor_list() -> List[str]:
    rows = _sb_data(
        get_supabase().table("cabor").select("nama").eq("aktif", True).order("nama").execute()
    )
    return [r["nama"] for r in rows]


def get_current_user() -> Optional[dict]:
    username = st.session_state.get("username")
    if not username:
        return None
    rows = _sb_data(
        get_supabase().table("users").select("*").eq("username", username).limit(1).execute()
    )
    return rows[0] if rows else None


def fetch_all_users() -> pd.DataFrame:
    rows = _sb_data(
        get_supabase().table("users").select("id,username,nama_lengkap,role,aktif,created_at")
        .order("id", desc=True).execute()
    )
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["id", "username", "nama_lengkap", "role", "status", "created_at"])
    df["status"] = df["aktif"].map({True: "Aktif", False: "Nonaktif"})
    return df[["id", "username", "nama_lengkap", "role", "status", "created_at"]]


def fetch_all_cabor() -> pd.DataFrame:
    rows = _sb_data(get_supabase().table("cabor").select("id,nama,aktif").order("nama").execute())
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["id", "nama", "status"])
    df["status"] = df["aktif"].map({True: "Aktif", False: "Nonaktif"})
    return df[["id", "nama", "status"]]


def _all_laporan_rows() -> List[dict]:
    return _sb_data(
        get_supabase().table("laporan_monitoring").select("*")
        .order("tanggal", desc=True).order("id", desc=True).execute()
    )


def fetch_laporan_summary() -> Dict[str, int]:
    rows = _all_laporan_rows()
    total = len(rows)
    bulan_ini = datetime.date.today().strftime("%Y-%m")
    return {
        "total": total,
        "bulan_ini": sum(str(r.get("tanggal", ""))[:7] == bulan_ini for r in rows),
        "cabor_termonitor": len({r.get("cabor") for r in rows if r.get("cabor")}),
        "perlu_tindak": sum((r.get("status") or DEFAULT_STATUS) == DEFAULT_STATUS for r in rows),
        "selesai": sum((r.get("status") or "") == "Selesai" for r in rows),
    }


def fetch_chart_data():
    rows = _all_laporan_rows()
    if not rows:
        return (
            pd.DataFrame(columns=["cabor", "jumlah"]),
            pd.DataFrame(columns=["bulan", "jumlah"]),
            pd.DataFrame(columns=["status", "jumlah"]),
        )
    df = pd.DataFrame(rows)
    by_cabor = (
        df.dropna(subset=["cabor"]).groupby("cabor").size().reset_index(name="jumlah")
        .sort_values("jumlah", ascending=False).head(15)
    )
    df["bulan"] = df["tanggal"].fillna("").astype(str).str[:7]
    by_month = df[df["bulan"].ne("")].groupby("bulan").size().reset_index(name="jumlah").sort_values("bulan")
    df["status"] = df["status"].fillna(DEFAULT_STATUS)
    by_status = df.groupby("status").size().reset_index(name="jumlah")
    return by_cabor, by_month, by_status


def fetch_laporan_list(
    petugas: Optional[str] = None,
    cabor: Optional[str] = None,
    status: Optional[str] = None,
    tgl_awal: Optional[datetime.date] = None,
    tgl_akhir: Optional[datetime.date] = None,
    keyword: Optional[str] = None,
) -> pd.DataFrame:
    rows = _all_laporan_rows()
    columns = ["id", "tanggal", "cabor", "lokasi", "petugas", "status", "updated_at"]
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=columns)
    if petugas:
        df = df[df["petugas"].fillna("").eq(petugas)]
    if cabor and cabor != "Semua":
        df = df[df["cabor"].fillna("").eq(cabor)]
    if status and status != "Semua":
        df = df[df["status"].fillna(DEFAULT_STATUS).eq(status)]
    if tgl_awal:
        df = df[df["tanggal"].astype(str).ge(str(tgl_awal))]
    if tgl_akhir:
        df = df[df["tanggal"].astype(str).le(str(tgl_akhir))]
    if keyword:
        kw = keyword.lower()
        mask = (
            df["lokasi"].fillna("").str.lower().str.contains(kw, regex=False)
            | df["petugas"].fillna("").str.lower().str.contains(kw, regex=False)
            | df["cabor"].fillna("").str.lower().str.contains(kw, regex=False)
        )
        df = df[mask]
    df["status"] = df["status"].fillna(DEFAULT_STATUS)
    return df[columns].sort_values(["tanggal", "id"], ascending=[False, False]).reset_index(drop=True)


def get_laporan_by_id(laporan_id: int) -> Optional[dict]:
    rows = _sb_data(get_supabase().table("laporan_monitoring").select("*").eq("id", int(laporan_id)).limit(1).execute())
    return rows[0] if rows else None


def get_fotos_by_laporan(laporan_id: int) -> List[dict]:
    return _sb_data(
        get_supabase().table("laporan_foto").select("*").eq("laporan_id", int(laporan_id)).order("id").execute()
    )


def _mime_from_ext(ext: str) -> str:
    return {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}.get(ext.lower(), "image/jpeg")


def get_foto_bytes(foto: dict) -> Optional[bytes]:
    path = foto.get("storage_path") or ""
    if not path:
        return None
    try:
        return get_supabase().storage.from_(SUPABASE_BUCKET).download(path)
    except Exception:
        return None


def _load_font(size: int = 24):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/arialbd.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size=size)
        except Exception:
            continue
    return ImageFont.load_default()


def add_timestamp_watermark(raw: bytes, timestamp_text: str) -> bytes:
    try:
        img = Image.open(BytesIO(raw)).convert("RGB")
        max_side = 2400
        if max(img.size) > max_side:
            img.thumbnail((max_side, max_side), Image.Resampling.LANCZOS)
        draw = ImageDraw.Draw(img, "RGBA")
        font_size = max(18, int(min(img.size) * 0.025))
        font = _load_font(font_size)
        text = f"MONITORING BINPRES | {timestamp_text}"
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        margin = max(12, int(font_size * 0.6))
        x = margin
        y = img.height - th - margin * 2
        draw.rounded_rectangle(
            (x - margin, y - margin, x + tw + margin, y + th + margin),
            radius=10,
            fill=(0, 0, 0, 155),
        )
        draw.text((x, y), text, font=font, fill=(255, 255, 255, 235))
        out = BytesIO()
        img.save(out, format="JPEG", quality=92, optimize=True)
        return out.getvalue()
    except Exception:
        return raw


def save_uploaded_photos(laporan_id: int, uploaded_files: list) -> int:
    if not uploaded_files:
        return 0
    count = 0
    sb = get_supabase()
    for f in uploaded_files[:MAX_PHOTOS_PER_REPORT]:
        ext = Path(f.name).suffix.lower()
        if ext not in {".jpg", ".jpeg", ".png", ".webp"}:
            continue
        raw = f.getbuffer().tobytes()
        if not raw:
            continue
        timestamp_text = datetime.datetime.now().strftime(PHOTO_TIMESTAMP_FORMAT)
        processed = add_timestamp_watermark(raw, timestamp_text)
        stem = safe_filename(Path(f.name).stem).replace(" ", "_")
        storage_path = f"laporan/{laporan_id}/{uuid.uuid4().hex}_{stem}.jpg"
        try:
            sb.storage.from_(SUPABASE_BUCKET).upload(
                storage_path,
                processed,
                {"content-type": "image/jpeg", "cache-control": "3600", "upsert": "false"},
            )
            sb.table("laporan_foto").insert({
                "laporan_id": int(laporan_id),
                "filename": Path(storage_path).name,
                "original_name": f.name,
                "mime_type": "image/jpeg",
                "storage_path": storage_path,
                "timestamp_mark": timestamp_text,
            }).execute()
            count += 1
        except Exception as exc:
            try:
                sb.storage.from_(SUPABASE_BUCKET).remove([storage_path])
            except Exception:
                pass
            st.warning(f"Foto '{f.name}' gagal disimpan: {exc}")
    return count


def delete_laporan(laporan_id: int) -> None:
    sb = get_supabase()
    fotos = get_fotos_by_laporan(laporan_id)
    paths = [f.get("storage_path") for f in fotos if f.get("storage_path")]
    if paths:
        try:
            sb.storage.from_(SUPABASE_BUCKET).remove(paths)
        except Exception:
            pass
    sb.table("laporan_foto").delete().eq("laporan_id", int(laporan_id)).execute()
    sb.table("laporan_monitoring").delete().eq("id", int(laporan_id)).execute()


def status_badge_html(status: str) -> str:
    s = status or DEFAULT_STATUS
    if s == "Selesai":
        cls = "badge-selesai"
    elif s == "Sedang Ditindaklanjuti":
        cls = "badge-sedang"
    else:
        cls = "badge-belum"
    return f'<span class="{cls}">{s}</span>'

# ============================================================
# WORD & EXCEL EXPORT
# ============================================================
def _set_run_font(run, size_pt: float = 10, bold: bool = False) -> None:
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.font.name = "Calibri"


def _add_compact_para(doc, text: str, bold: bool = False, size: float = 10, space_after: float = 2) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(text)
    _set_run_font(run, size_pt=size, bold=bold)


def generate_word_report(
    data_list: List[Dict[str, Any]],
    is_all: bool = False,
    include_photos: bool = True,
) -> bytes:
    """
    Layout hemat kertas:
    - Margin kecil, font 10pt, spasi rapat
    - Lembar isi laporan dulu
    - Lembar foto terpisah (halaman baru) di akhir tiap laporan
    - 1 file Word
    """
    doc = Document()
    section = doc.sections[0]
    # Margin rapat (hemat kertas)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.6)
    section.right_margin = Inches(0.6)

    # Style default body
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(2)
    style.paragraph_format.line_spacing = 1.0

    judul = "REKAPITULASI LAPORAN MONEV BINPRES" if is_all else "LAPORAN MONEV BINPRES"
    title = doc.add_heading(judul, level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.size = Pt(14)
        run.font.name = "Calibri"

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.paragraph_format.space_after = Pt(6)
    r = meta.add_run(
        f"KONI Kabupaten Tangerang  |  Dicetak: {datetime.date.today().strftime('%d/%m/%Y')}"
    )
    _set_run_font(r, size_pt=9)

    def add_info_line(label: str, value: Any) -> None:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.0
        r1 = p.add_run(f"{label}: ")
        _set_run_font(r1, size_pt=10, bold=True)
        r2 = p.add_run(str(value) if value else "-")
        _set_run_font(r2, size_pt=10)

    def add_section_compact(title_text: str, items: List[Tuple[str, Any]]) -> None:
        # Judul section
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(6)
        h.paragraph_format.space_after = Pt(2)
        hr = h.add_run(title_text)
        _set_run_font(hr, size_pt=11, bold=True)

        for question, answer in items:
            # Pertanyaan + jawaban dalam 1 blok rapat
            pq = doc.add_paragraph()
            pq.paragraph_format.space_before = Pt(2)
            pq.paragraph_format.space_after = Pt(0)
            pq.paragraph_format.line_spacing = 1.0
            rq = pq.add_run(f"• {question}")
            _set_run_font(rq, size_pt=9, bold=True)

            pa = doc.add_paragraph()
            pa.paragraph_format.space_before = Pt(0)
            pa.paragraph_format.space_after = Pt(2)
            pa.paragraph_format.line_spacing = 1.0
            pa.paragraph_format.left_indent = Inches(0.15)
            ra = pa.add_run(str(answer).strip() if answer and str(answer).strip() else "—")
            _set_run_font(ra, size_pt=9)

    # Kumpulkan foto per laporan untuk halaman terpisah
    photos_queue: List[Tuple[Dict[str, Any], list]] = []

    for idx, data in enumerate(data_list):
        if is_all and idx > 0:
            doc.add_page_break()

        if is_all:
            h = doc.add_paragraph()
            h.paragraph_format.space_before = Pt(4)
            h.paragraph_format.space_after = Pt(4)
            hr = h.add_run(
                f"Laporan {idx + 1}: {data.get('cabor', '-')} — {data.get('tanggal', '-')}"
            )
            _set_run_font(hr, size_pt=12, bold=True)

        # --- Lembar isi (rapat) ---
        add_info_line("Cabang Olahraga", data.get("cabor"))
        add_info_line("Tanggal", data.get("tanggal"))
        add_info_line("Lokasi", data.get("lokasi"))
        add_info_line("Petugas Monev", data.get("petugas"))
        add_info_line("Status", data.get("status") or DEFAULT_STATUS)
        add_info_line("Perwakilan KONI yang Hadir", data.get("daftar_hadir_koni"))

        # Garis pemisah tipis
        sep = doc.add_paragraph()
        sep.paragraph_format.space_before = Pt(2)
        sep.paragraph_format.space_after = Pt(2)
        sr = sep.add_run("─" * 55)
        _set_run_font(sr, size_pt=8)

        for section_title, fields in FIELD_LABELS.items():
            items = [(label, data.get(key)) for key, label in fields]
            add_section_compact(section_title, items)

        if data.get("catatan_admin"):
            add_section_compact("Catatan Admin / Tindak Lanjut", [
                ("Catatan", data["catatan_admin"]),
            ])

        # Simpan foto untuk halaman terpisah
        if include_photos and data.get("id"):
            fotos = get_fotos_by_laporan(int(data["id"]))
            if fotos:
                photos_queue.append((data, fotos))

    # --- Lembar foto terpisah (di akhir, tetap 1 file) ---
    if include_photos and photos_queue:
        for data, fotos in photos_queue:
            doc.add_page_break()

            h = doc.add_paragraph()
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            h.paragraph_format.space_after = Pt(4)
            hr = h.add_run("DOKUMENTASI FOTO")
            _set_run_font(hr, size_pt=12, bold=True)

            sub = doc.add_paragraph()
            sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            sub.paragraph_format.space_after = Pt(8)
            sr = sub.add_run(
                f"{data.get('cabor', '-')}  |  {data.get('tanggal', '-')}  |  "
                f"{data.get('lokasi', '-')}  |  Petugas: {data.get('petugas', '-')}"
            )
            _set_run_font(sr, size_pt=9)

            for foto in fotos:
                img_bytes = get_foto_bytes(foto)
                if img_bytes:
                    try:
                        # Lebar sedang agar hemat ruang, bisa 2 foto per halaman
                        doc.add_picture(BytesIO(img_bytes), width=Inches(5.2))
                        cap = doc.add_paragraph()
                        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        cap.paragraph_format.space_before = Pt(2)
                        cap.paragraph_format.space_after = Pt(8)
                        cr = cap.add_run(f'{foto["original_name"] or foto["filename"]} | Time-mark: {foto["timestamp_mark"] or "-"}')
                        _set_run_font(cr, size_pt=8)
                    except Exception:
                        _add_compact_para(
                            doc, f"[Gagal memuat: {foto['original_name']}]", size=9
                        )
                else:
                    _add_compact_para(
                        doc, f"[Foto tidak tersedia: {foto['original_name']}]", size=9
                    )

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


def generate_excel_report(df: pd.DataFrame) -> bytes:
    if df.empty:
        out = BytesIO()
        pd.DataFrame({"info": ["Tidak ada data"]}).to_excel(out, index=False)
        return out.getvalue()
    rows = []
    for laporan_id in [int(x) for x in df["id"].tolist()]:
        row = get_laporan_by_id(laporan_id)
        if row:
            rows.append(row)
    detail = pd.DataFrame(rows)
    rename_map = {
        "id": "ID", "tanggal": "Tanggal", "cabor": "Cabang Olahraga", "lokasi": "Lokasi", "petugas": "Petugas",
        "status": "Status", "catatan_admin": "Catatan Admin", "daftar_hadir_koni": "Daftar Hadir Perwakilan KONI",
        "fisik_parameter": "Fisik - Parameter", "fisik_peaking": "Fisik - Peaking", "fisik_recovery": "Fisik - Recovery", "fisik_cedera": "Fisik - Cedera",
        "taktis_lawan": "Taktis - Lawan", "taktis_instruksi": "Taktis - Instruksi", "taktis_ujicoba": "Taktis - Ujicoba",
        "mental_cemas": "Mental - Cemas", "mental_fokus": "Mental - Fokus", "mental_rutinitas": "Mental - Rutinitas", "mental_psikolog": "Mental - Psikolog",
        "nutrisi_bb": "Nutrisi - BB", "nutrisi_asupan": "Nutrisi - Asupan", "nutrisi_hidrasi": "Nutrisi - Hidrasi", "nutrisi_tidur": "Nutrisi - Tidur",
        "medis_rekam": "Medis - Rekam", "medis_doping": "Medis - Doping", "medis_alat": "Medis - Alat", "medis_nonteknis": "Medis - Nonteknis",
        "updated_at": "Diperbarui", "created_at": "Dibuat",
    }
    if not detail.empty:
        detail = detail.rename(columns={k: v for k, v in rename_map.items() if k in detail.columns})
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        detail.to_excel(writer, sheet_name="Laporan Monitoring", index=False)
    return out.getvalue()

# ============================================================
# UI COMPONENTS
# ============================================================
def render_sidebar() -> None:
    user = get_current_user()
    with st.sidebar:
        st.markdown("## 🏆 BINPRES")
        st.caption("Monitoring & Evaluasi Cabor")
        st.markdown("---")
        nama = user.get("nama_lengkap") if user and user.get("nama_lengkap") else st.session_state.get("username", "")
        st.write(f"👤 **{nama}**")
        st.caption(f"Role: {st.session_state.get('role', '').upper()}")
        st.markdown("---")
        with st.expander("🔑 Ganti Password"):
            with st.form("form_ganti_password"):
                pw_lama = st.text_input("Password lama", type="password")
                pw_baru = st.text_input("Password baru", type="password")
                pw_konfirm = st.text_input("Konfirmasi password baru", type="password")
                if st.form_submit_button("Simpan Password", use_container_width=True):
                    if not pw_lama or not pw_baru:
                        st.error("Semua field wajib diisi.")
                    elif len(pw_baru) < 6:
                        st.error("Password baru minimal 6 karakter.")
                    elif pw_baru != pw_konfirm:
                        st.error("Konfirmasi password tidak cocok.")
                    else:
                        u = get_current_user()
                        if u and verify_password(pw_lama, u["password"]):
                            get_supabase().table("users").update({"password": hash_password(pw_baru)}).eq("id", int(u["id"])).execute()
                            st.success("Password berhasil diubah.")
                        else:
                            st.error("Password lama salah.")
        st.markdown("---")
        if st.button("🚪 Logout", use_container_width=True):
            logout()

def halaman_login() -> None:
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    st.markdown(
        """
        <div style="text-align:center;padding:40px 0 10px 0">
            <div style="font-size:56px">🏆</div>
            <h1 style="margin-bottom:4px;color:#1e3a8a">Sistem Monitoring BINPRES</h1>
            <p style="color:#64748b;font-size:1.05rem">KONI Kabupaten Tangerang</p>
        </div>
        """, unsafe_allow_html=True,
    )
    _, col_c, _ = st.columns([1, 1.4, 1])
    with col_c:
        with st.form("login_form"):
            st.markdown("#### Masuk ke Sistem")
            username = st.text_input("👤 Username", placeholder="Masukkan username")
            password = st.text_input("🔒 Password", type="password", placeholder="Masukkan password")
            submit = st.form_submit_button("Masuk ke Sistem", use_container_width=True, type="primary")
            if submit:
                rows = _sb_data(get_supabase().table("users").select("*").eq("username", username.strip()).eq("aktif", True).limit(1).execute())
                user = rows[0] if rows else None
                if user and verify_password(password, user["password"]):
                    st.session_state["logged_in"] = True
                    st.session_state["username"] = user["username"]
                    st.session_state["role"] = user["role"]
                    st.session_state["nama_lengkap"] = user.get("nama_lengkap") or user["username"]
                    st.rerun()
                else:
                    st.error("Username atau Password salah / akun tidak aktif.")
        st.caption("Login utama: userkoni / koni123 • Database & foto tersimpan di Supabase")

def render_detail_laporan(row: dict) -> None:
    st.markdown(
        f"**Cabor:** {row['cabor']} &nbsp;|&nbsp; "
        f"**Tanggal:** {row['tanggal']} &nbsp;|&nbsp; "
        f"**Lokasi:** {row['lokasi']} &nbsp;|&nbsp; "
        f"**Petugas:** {row['petugas']}"
    )
    st.markdown(
        f"**Status:** {status_badge_html(row['status'] or DEFAULT_STATUS)}",
        unsafe_allow_html=True,
    )
    if row["catatan_admin"]:
        st.info(f"**Catatan Admin:** {row['catatan_admin']}")

    if row["daftar_hadir_koni"]:
        st.markdown("#### 👥 Perwakilan KONI Kabupaten Tangerang yang Hadir")
        st.write(row["daftar_hadir_koni"])

    for section_title, fields in FIELD_LABELS.items():
        with st.expander(section_title, expanded=False):
            for key, label in fields:
                val = row[key] if key in row.keys() else None
                st.markdown(f"**{label}**")
                st.write(val if val else "—")

    fotos = get_fotos_by_laporan(int(row["id"]))
    if fotos:
        st.markdown("#### 🖼️ Dokumentasi Foto")
        cols = st.columns(min(4, len(fotos)))
        for i, foto in enumerate(fotos):
            img_bytes = get_foto_bytes(foto)
            if img_bytes:
                with cols[i % len(cols)]:
                    st.image(
                        img_bytes,
                        caption=foto["original_name"],
                        use_container_width=True,
                    )


# ------------------------------------------------------------
# DASHBOARD ADMIN
# ------------------------------------------------------------
def dashboard_admin() -> None:
    stats = fetch_laporan_summary()

    st.markdown(
        """
        <div class="main-header">
            <h1>📊 Dashboard Monitoring BINPRES</h1>
            <p>Ringkasan aktivitas monitoring & evaluasi cabang olahraga</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Laporan", stats["total"])
    c2.metric("Bulan Ini", stats["bulan_ini"])
    c3.metric("Cabor Termonitor", stats["cabor_termonitor"])
    c4.metric("Perlu Tindak Lanjut", stats["perlu_tindak"])
    c5.metric("Selesai", stats["selesai"])

    st.markdown("---")

    by_cabor, by_month, by_status = fetch_chart_data()

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("##### 🏅 Top Cabor (jumlah laporan)")
        if not by_cabor.empty:
            st.bar_chart(by_cabor.set_index("cabor")["jumlah"], height=280)
        else:
            st.caption("Belum ada data.")

    with col_b:
        st.markdown("##### 📅 Tren Laporan per Bulan")
        if not by_month.empty:
            st.line_chart(by_month.set_index("bulan")["jumlah"], height=280)
        else:
            st.caption("Belum ada data.")

    st.markdown("##### 📌 Distribusi Status")
    if not by_status.empty:
        st.bar_chart(by_status.set_index("status")["jumlah"], height=200)
    else:
        st.caption("Belum ada data.")

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
                    get_supabase().table("users").insert({
                        "username": username.strip(), "password": hash_password(password),
                        "role": role, "nama_lengkap": nama.strip(), "aktif": True,
                    }).execute()
                    st.success("User berhasil ditambahkan.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"User gagal ditambahkan. Pastikan username belum digunakan. Detail: {exc}")
    st.markdown("### ✏️ Edit / Nonaktifkan User")
    users = _sb_data(get_supabase().table("users").select("id,username,nama_lengkap,role,aktif").order("username").execute())
    if not users:
        return
    pilihan = st.selectbox("Pilih user", users, format_func=lambda x: f"{x['username']} — {x.get('nama_lengkap') or '-'}")
    with st.form("edit_user"):
        nama_baru = st.text_input("Nama lengkap", value=pilihan.get("nama_lengkap") or "")
        role_baru = st.selectbox("Role", ["user", "admin"], index=0 if pilihan.get("role") == "user" else 1)
        aktif_baru = st.checkbox("Akun aktif", value=bool(pilihan.get("aktif")))
        password_baru = st.text_input("Password baru (kosongkan jika tidak diubah)", type="password")
        if st.form_submit_button("💾 Simpan Perubahan", use_container_width=True):
            if password_baru and len(password_baru) < 6:
                st.error("Password minimal 6 karakter.")
                return
            payload = {"nama_lengkap": nama_baru.strip(), "role": role_baru, "aktif": bool(aktif_baru)}
            if password_baru:
                payload["password"] = hash_password(password_baru)
            try:
                get_supabase().table("users").update(payload).eq("id", int(pilihan["id"])).execute()
                st.success("Data user berhasil diperbarui.")
                st.rerun()
            except Exception as exc:
                st.error(f"Gagal memperbarui user: {exc}")

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
                    get_supabase().table("cabor").insert({"nama": nama_cabor.strip().upper(), "aktif": True}).execute()
                    st.success("Cabor berhasil ditambahkan.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Cabor tersebut mungkin sudah ada. Detail: {exc}")
    st.markdown("### ✏️ Edit / Nonaktifkan Cabor")
    rows = _sb_data(get_supabase().table("cabor").select("*").order("nama").execute())
    if not rows:
        return
    pilihan = st.selectbox("Pilih cabor", rows, format_func=lambda x: x["nama"])
    with st.form("edit_cabor"):
        nama_baru = st.text_input("Nama cabor", value=pilihan["nama"])
        aktif = st.checkbox("Cabor aktif", value=bool(pilihan["aktif"]))
        if st.form_submit_button("💾 Simpan", use_container_width=True):
            try:
                get_supabase().table("cabor").update({"nama": nama_baru.strip().upper(), "aktif": bool(aktif)}).eq("id", int(pilihan["id"])).execute()
                st.success("Cabor berhasil diperbarui.")
                st.rerun()
            except Exception as exc:
                st.error(f"Gagal memperbarui cabor: {exc}")

def halaman_laporan_admin() -> None:
    st.subheader("📄 Data Laporan Monitoring")

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        cabor_opts = ["Semua"] + get_cabor_list()
        cabor_filter = st.selectbox("Filter Cabor", cabor_opts)
    with f2:
        status_filter = st.selectbox("Filter Status", ["Semua"] + STATUS_OPTIONS)
    with f3:
        tgl_awal = st.date_input("Dari tanggal", value=None)
    with f4:
        tgl_akhir = st.date_input("Sampai tanggal", value=None)

    keyword = st.text_input(
        "🔎 Cari lokasi / petugas / cabor", placeholder="Ketik kata kunci..."
    )

    df = fetch_laporan_list(
        cabor=cabor_filter,
        status=status_filter,
        tgl_awal=tgl_awal if tgl_awal else None,
        tgl_akhir=tgl_akhir if tgl_akhir else None,
        keyword=keyword.strip() or None,
    )

    st.caption(f"Menampilkan **{len(df)}** laporan.")
    if df.empty:
        st.info("Tidak ada laporan sesuai filter.")
        return

    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 🔍 Detail, Export & Tindak Lanjut")

    pilihan_id = st.selectbox(
        "Pilih laporan",
        df["id"].tolist(),
        format_func=lambda x: (
            f"ID {x} — {df.loc[df['id']==x, 'cabor'].values[0]} "
            f"({df.loc[df['id']==x, 'tanggal'].values[0]}) — "
            f"{df.loc[df['id']==x, 'status'].values[0]}"
        ),
        key="admin_pilih_laporan",
    )

    row = get_laporan_by_id(int(pilihan_id))
    if not row:
        st.error("Laporan tidak ditemukan.")
        return

    with st.expander("📋 Lihat Detail Lengkap", expanded=False):
        render_detail_laporan(row)

    col_w, col_x, col_d = st.columns(3)
    with col_w:
        word_file = generate_word_report([dict(row)], is_all=False)
        st.download_button(
            "⬇️ Word Satuan",
            data=word_file,
            file_name=safe_filename(f"Laporan_{row['cabor']}_{row['tanggal']}.docx"),
            mime=WORD_MIME,
            use_container_width=True,
        )
    with col_x:
        excel_file = generate_excel_report(df)
        st.download_button(
            "⬇️ Excel (hasil filter)",
            data=excel_file,
            file_name=f"Rekap_Monev_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_d:
        if st.button("🗑️ Hapus Laporan", use_container_width=True):
            st.session_state["confirm_delete_id"] = int(pilihan_id)

    if st.session_state.get("confirm_delete_id") == int(pilihan_id):
        st.warning(
            f"Yakin hapus laporan ID {pilihan_id}? Tindakan ini tidak dapat dibatalkan."
        )
        c_yes, c_no = st.columns(2)
        with c_yes:
            if st.button("✅ Ya, Hapus", type="primary", use_container_width=True):
                delete_laporan(int(pilihan_id))
                st.session_state.pop("confirm_delete_id", None)
                st.success("Laporan berhasil dihapus.")
                st.rerun()
        with c_no:
            if st.button("❌ Batal", use_container_width=True):
                st.session_state.pop("confirm_delete_id", None)
                st.rerun()

    st.markdown("#### 📚 Export Rekap Word (semua hasil filter)")
    if st.button("Generate Word Rekap", use_container_width=True):
        rows = []
        for laporan_id in sorted(int(x) for x in df["id"].tolist()):
            row_rekap = get_laporan_by_id(laporan_id)
            if row_rekap:
                rows.append(row_rekap)
        word_rekap = generate_word_report([dict(r) for r in rows], is_all=True)
        st.download_button(
            "⬇️ Download Rekap Word",
            data=word_rekap,
            file_name=f"Rekap_Monev_{datetime.date.today()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    st.markdown("---")
    st.markdown("### ✏️ Tindak Lanjut Laporan")
    with st.form("update_status_laporan"):
        current = row["status"] or DEFAULT_STATUS
        status = st.selectbox(
            "Status tindak lanjut",
            STATUS_OPTIONS,
            index=STATUS_OPTIONS.index(current) if current in STATUS_OPTIONS else 0,
        )
        catatan = st.text_area(
            "Catatan Admin / Rekomendasi",
            value=row["catatan_admin"] or "",
            height=100,
        )
        if st.form_submit_button("💾 Simpan Tindak Lanjut", use_container_width=True):
            get_supabase().table("laporan_monitoring").update({
                "status": status,
                "catatan_admin": catatan,
                "updated_at": _now_iso(),
            }).eq("id", int(pilihan_id)).execute()
            st.success("Tindak lanjut berhasil diperbarui.")
            st.rerun()


# ------------------------------------------------------------
# HALAMAN ADMIN
# ------------------------------------------------------------
def halaman_admin() -> None:
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
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
# HALAMAN PETUGAS
# ------------------------------------------------------------
def form_input_monitoring() -> None:
    st.markdown(
        """
        <div class="main-header">
            <h1>📝 Form Input Monitoring</h1>
            <p>Isi laporan monitoring cabang olahraga secara lengkap</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(
        f"Petugas: **{st.session_state.get('nama_lengkap', st.session_state['username'])}**"
    )

    cabor_list = get_cabor_list()

    # ----------------------------------------------------------
    # UPLOAD FOTO DI LUAR FORM
    # (file_uploader di dalam form sering tidak muncul di Cloud)
    # ----------------------------------------------------------
    st.markdown(
        '<span class="section-badge">📷 Dokumentasi Foto</span>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Upload foto kegiatan **sebelum** mengisi form di bawah. Bisa lebih dari satu foto."
    )
    uploaded_files = st.file_uploader(
        "Pilih foto (JPG / PNG / WEBP)",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        key="uploader_monitoring",
        help="Foto akan ikut tersimpan di laporan dan muncul di file Word saat diunduh.",
    )
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} foto siap diunggah bersama laporan.")
        preview_cols = st.columns(min(4, len(uploaded_files)))
        for i, f in enumerate(uploaded_files):
            with preview_cols[i % len(preview_cols)]:
                st.image(f, caption=f.name, use_container_width=True)
    else:
        st.info(
            "Belum ada foto dipilih. Foto monitoring wajib minimal 1 foto."
        )

    st.markdown("---")

    with st.form("form_monitoring", clear_on_submit=True):
        st.markdown(
            '<span class="section-badge">📌 Informasi Dasar</span>',
            unsafe_allow_html=True,
        )
        col1, col2 = st.columns(2)
        with col1:
            tanggal = st.date_input("Tanggal Monitoring", datetime.date.today())
            cabor = st.selectbox("Cabang Olahraga", ["Pilih Cabor..."] + cabor_list)
        with col2:
            lokasi = st.text_input(
                "Lokasi Latihan / Try-out", placeholder="Contoh: GOR Cikokol"
            )

        st.markdown("---")

        with st.expander("💪 1. Performa Fisik & Kebugaran", expanded=True):
            fisik_1 = st.text_area(
                "Capaian parameter fisik (vs benchmark target):", height=70
            )
            fisik_2 = st.text_area(
                "Apakah atlet mencapai grafik performa puncak (peaking)?", height=70
            )
            fisik_3 = st.text_area("Tingkat pemulihan fisik (recovery):", height=70)
            fisik_4 = st.text_area(
                "Keluhan cedera lama / indikasi cedera baru:", height=70
            )

        with st.expander("🎯 2. Kesiapan Taktis & Strategi"):
            taktis_1 = st.text_area("Pemetaan kekuatan calon lawan:", height=70)
            taktis_2 = st.text_area(
                "Kemampuan mengikuti instruksi teknis di bawah tekanan:", height=70
            )
            taktis_3 = st.text_area("Hasil try-out / sparing:", height=70)

        with st.expander("🧠 3. Mental, Psikologis & Kesiapan Mental"):
            mental_1 = st.text_area(
                "Tingkat kecemasan & kemampuan mengendalikan stres:", height=70
            )
            mental_2 = st.text_area(
                "Fokus, motivasi, dan self-confidence:", height=70
            )
            mental_3 = st.text_area(
                "Rutinitas mental khusus saat bertanding:", height=70
            )
            mental_4 = st.text_area(
                "Koordinasi dengan tim psikolog olahraga:", height=70
            )

        with st.expander("🥗 4. Nutrisi, Berat Badan & Gaya Hidup"):
            nutrisi_1 = st.text_area("Progres penyesuaian berat badan:", height=70)
            nutrisi_2 = st.text_area(
                "Pemantauan asupan nutrisi dan suplemen:", height=70
            )
            nutrisi_3 = st.text_area("Status hidrasi harian atlet:", height=70)
            nutrisi_4 = st.text_area(
                "Kualitas dan kecukupan waktu tidur:", height=70
            )

        with st.expander("⚕️ 5. Medis, Bebas Doping & Logistik"):
            medis_1 = st.text_area(
                "Status rekam medis terkini & kesiapan fisioterapis:", height=70
            )
            medis_2 = st.text_area(
                "Keamanan obat, suplemen (Bebas Doping):", height=70
            )
            medis_3 = st.text_area(
                "Kesiapan perlengkapan khusus bertanding:", height=70
            )
            medis_4 = st.text_area(
                "Kendala non-teknis (akomodasi, transportasi):", height=70
            )

        st.markdown("---")
        st.markdown(
            '<span class="section-badge">👥 Daftar Hadir Perwakilan KONI Kabupaten Tangerang</span>',
            unsafe_allow_html=True,
        )
        daftar_hadir_koni = st.text_area(
            "Nama yang hadir mewakili KONI Kabupaten Tangerang",
            placeholder="Contoh:\n1. Nama — Jabatan\n2. Nama — Jabatan",
            height=110,
            help="Wajib diisi. Satu orang per baris.",
        )

        n_ready = len(uploaded_files) if uploaded_files else 0
        if n_ready:
            st.caption(
                f"📷 {n_ready} foto siap disimpan. Setiap foto akan diberi time-mark server "
                f"({datetime.datetime.now().strftime(PHOTO_TIMESTAMP_FORMAT)})."
            )
        else:
            st.error("📷 Foto monitoring wajib diunggah. Minimal 1 foto.")

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
            elif not uploaded_files:
                st.error("⚠️ Foto monitoring wajib diunggah minimal 1 foto.")
            elif len(uploaded_files) > MAX_PHOTOS_PER_REPORT:
                st.error(f"⚠️ Maksimal {MAX_PHOTOS_PER_REPORT} foto per laporan.")
            elif any(len(f.getbuffer()) > MAX_PHOTO_MB * 1024 * 1024 for f in uploaded_files):
                st.error(f"⚠️ Ukuran setiap foto maksimal {MAX_PHOTO_MB} MB.")
            elif not daftar_hadir_koni.strip():
                st.error("⚠️ Daftar hadir perwakilan KONI wajib diisi.")
            else:
                payload_laporan = {
                    "tanggal": tanggal.isoformat(),
                    "cabor": cabor,
                    "lokasi": lokasi.strip(),
                    "petugas": st.session_state["username"],
                    "fisik_parameter": fisik_1,
                    "fisik_peaking": fisik_2,
                    "fisik_recovery": fisik_3,
                    "fisik_cedera": fisik_4,
                    "taktis_lawan": taktis_1,
                    "taktis_instruksi": taktis_2,
                    "taktis_ujicoba": taktis_3,
                    "mental_cemas": mental_1,
                    "mental_fokus": mental_2,
                    "mental_rutinitas": mental_3,
                    "mental_psikolog": mental_4,
                    "nutrisi_bb": nutrisi_1,
                    "nutrisi_asupan": nutrisi_2,
                    "nutrisi_hidrasi": nutrisi_3,
                    "nutrisi_tidur": nutrisi_4,
                    "medis_rekam": medis_1,
                    "medis_doping": medis_2,
                    "medis_alat": medis_3,
                    "medis_nonteknis": medis_4,
                    "status": DEFAULT_STATUS,
                    "catatan_admin": "",
                    "daftar_hadir_koni": daftar_hadir_koni.strip(),
                }
                result_insert = get_supabase().table("laporan_monitoring").insert(payload_laporan).execute()
                inserted_rows = _sb_data(result_insert)
                if not inserted_rows:
                    st.error("Laporan gagal disimpan ke Supabase.")
                    st.stop()
                laporan_id = int(inserted_rows[0]["id"])

                files_to_save = (
                    st.session_state.get("uploader_monitoring")
                    or uploaded_files
                    or []
                )
                n_foto = save_uploaded_photos(laporan_id, files_to_save)

                # Simpan ID laporan terakhir agar bisa langsung diunduh
                st.session_state["last_saved_laporan_id"] = int(laporan_id)

                msg = f"✅ Laporan berhasil disimpan (ID: {laporan_id})."
                if n_foto:
                    msg += f" **{n_foto} foto** ikut tersimpan dan akan muncul di file Word."
                else:
                    msg += " (tanpa foto)"
                st.success(msg)

                st.info(
                    "Admin dapat melihat dan menindaklanjuti laporan Anda. "
                    "Cek tab **Laporan Saya** untuk melihat status dan mengunduh Word."
                )

    # --- Tombol Download Word langsung setelah simpan (di luar form) ---
    last_id = st.session_state.get("last_saved_laporan_id")
    if last_id:
        row_last = get_laporan_by_id(int(last_id))
        if row_last:
            st.markdown("---")
            st.markdown("### ⬇️ Download Laporan yang Baru Disimpan")
            word_file = generate_word_report([dict(row_last)], is_all=False, include_photos=True)
            st.download_button(
                label="⬇️ Download Word Laporan Ini (termasuk foto)",
                data=word_file,
                file_name=safe_filename(f"Laporan_{row_last['cabor']}_{row_last['tanggal']}.docx"),
                mime=WORD_MIME,
                use_container_width=True,
                type="primary",
                key=f"dl_after_save_{last_id}",
            )
            st.caption("File Word berisi isi lengkap laporan + dokumentasi foto (jika ada).")


def halaman_laporan_saya() -> None:
    st.markdown(
        """
        <div class="main-header">
            <h1>📂 Laporan Saya</h1>
            <p>Daftar laporan yang pernah Anda input — unduh Word di setiap laporan</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    username = st.session_state["username"]

    f1, f2 = st.columns(2)
    with f1:
        status_f = st.selectbox(
            "Filter Status", ["Semua"] + STATUS_OPTIONS, key="user_status_f"
        )
    with f2:
        keyword = st.text_input("🔎 Cari cabor / lokasi", key="user_kw")

    df = fetch_laporan_list(
        petugas=username,
        status=status_f if status_f != "Semua" else None,
        keyword=keyword.strip() or None,
    )

    if df.empty:
        st.info(
            "Belum ada laporan. Silakan input melalui tab **Form Input Monitoring**."
        )
        return

    st.caption(f"Total **{len(df)}** laporan Anda.")
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 📥 Detail & Download Word per Laporan")

    pilihan_id = st.selectbox(
        "Pilih laporan untuk diunduh / dilihat detail",
        df["id"].tolist(),
        format_func=lambda x: (
            f"ID {x} — {df.loc[df['id']==x, 'cabor'].values[0]} "
            f"({df.loc[df['id']==x, 'tanggal'].values[0]}) — "
            f"{df.loc[df['id']==x, 'status'].values[0]}"
        ),
        key="user_pilih_laporan",
    )

    row = get_laporan_by_id(int(pilihan_id))
    if not row:
        st.error("Laporan tidak ditemukan.")
        return

    with st.expander("📋 Lihat Detail Lengkap", expanded=True):
        render_detail_laporan(row)

    # --- Download Word selalu tersedia di setiap laporan ---
    st.markdown("### ⬇️ Download File Word")
    word_file = generate_word_report([dict(row)], is_all=False, include_photos=True)
    st.download_button(
        label="⬇️ Download Laporan Word (termasuk foto)",
        data=word_file,
        file_name=safe_filename(f"Laporan_{row['cabor']}_{row['tanggal']}.docx"),
        mime=WORD_MIME,
        use_container_width=True,
        type="primary",
        key=f"user_dl_word_{row['id']}",
    )
    st.caption("File berisi data lengkap + dokumentasi foto (jika ada).")



def halaman_user() -> None:
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    render_sidebar()

    tab1, tab2 = st.tabs([
        "📝 Form Input Monitoring",
        "📂 Laporan Saya",
    ])
    with tab1:
        form_input_monitoring()
    with tab2:
        halaman_laporan_saya()


# ============================================================
# ENTRY POINT
# ============================================================
init_backend()

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    halaman_login()
else:
    if st.session_state.get("role") == "admin":
        halaman_admin()
    else:
        halaman_user()
