from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

try:
  TZ_WIB = ZoneInfo("Asia/Jakarta")
except Exception:
  TZ_WIB = timezone(timedelta(hours=7))

import hmac
import html
import re
import time
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Tracking Jadwal PR & Gantt Chart",
    page_icon="📋",
    layout="wide",
)

# --- KONFIGURASI KEAMANAN ---
MAX_PIN_ATTEMPTS = 3
LOCKOUT_SECONDS = 60
SESSION_TIMEOUT_SECONDS = 5 * 60  # 5 menit
MAX_EXCEL_SIZE_MB = 10
MAX_EXCEL_SIZE_BYTES = MAX_EXCEL_SIZE_MB * 1024 * 1024
MAX_EXCEL_SHEETS = 15
MAX_EXCEL_ROWS = 10000

# --- STYLING RINGAN UNTUK MEMPERCANTIK KOMPONEN BAWAAN STREAMLIT ---
st.markdown(
    """
    <style>
      .stApp {
        background-color: #f8fafc;
      }
      div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-top: 4px solid #2563eb;
        border-radius: 10px;
        padding: 0.65rem 0.75rem;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
      }
      div[data-testid="stMetricLabel"] p {
        font-size: 0.82rem !important;
        white-space: normal !important;
        overflow-wrap: break-word !important;
        line-height: 1.2 !important;
      }
      div[data-testid="stMetricValue"] {
        font-size: 1.15rem !important;
        line-height: 1.25 !important;
      }
      div[data-testid="stMetricValue"] > div {
        font-size: 1.15rem !important;
        white-space: normal !important;
        overflow: visible !important;
        text-overflow: unset !important;
        word-break: break-word !important;
      }
      div[data-testid="stMetricDelta"] {
        font-size: 0.78rem !important;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- BLOK PENGECEKAN PIN DENGAN BATASAN 3x PERCOBAAN & TIMEOUT 5 MENIT ---
if "pin_valid" not in st.session_state:
  st.session_state.pin_valid = False
if "login_at" not in st.session_state:
  st.session_state.login_at = 0.0
if "failed_pin_attempts" not in st.session_state:
  st.session_state.failed_pin_attempts = 0
if "pin_lock_until" not in st.session_state:
  st.session_state.pin_lock_until = 0.0
if "session_expired_msg" not in st.session_state:
  st.session_state.session_expired_msg = ""
if "uploader_key" not in st.session_state:
  st.session_state.uploader_key = 0


def _clear_uploaded_data():
  """Menghapus file unggahan & filter PR dari state sesi aktif."""
  st.session_state.uploader_key += 1
  if "multiselect_nomor_pr" in st.session_state:
    del st.session_state["multiselect_nomor_pr"]


def _logout_session(expired=False):
  """Mengunci ulang sesi dan membersihkan data unggahan dari memori sesi."""
  _clear_uploaded_data()
  st.session_state.pin_valid = False
  st.session_state.login_at = 0.0
  if expired:
    st.session_state.session_expired_msg = (
        "Sesi otomatis berakhir setelah 5 menit. Silakan masukkan PIN kembali."
    )


now_ts = time.time()

# Cek apakah sesi aktif sudah melewati batas waktu 5 menit
if st.session_state.pin_valid and st.session_state.login_at > 0:
  if (now_ts - st.session_state.login_at) >= SESSION_TIMEOUT_SECONDS:
    _logout_session(expired=True)
    st.rerun()

if not st.session_state.pin_valid:
  _, col_mid, _ = st.columns([1, 1.3, 1])
  with col_mid:
    with st.container(border=True):
      st.subheader("🔒 Verifikasi Akses")

      if st.session_state.session_expired_msg:
        st.warning(st.session_state.session_expired_msg)

      if now_ts < st.session_state.pin_lock_until:
        sisa_detik = max(1, int(st.session_state.pin_lock_until - now_ts))
        st.error(
            f"Terlalu banyak percobaan PIN gagal (maksimal {MAX_PIN_ATTEMPTS}x). "
            f"Silakan tunggu **{sisa_detik} detik** sebelum mencoba lagi."
        )
        st.stop()

      st.caption(
          f"Masukkan PIN untuk membuka aplikasi Tracking Jadwal PR "
          f"(maksimal {MAX_PIN_ATTEMPTS}x percobaan • durasi sesi 5 menit)."
      )
      input_pin = st.text_input(
          "PIN Akses", type="password", placeholder="Masukkan PIN..."
      )
      if st.button("Masuk", type="primary", use_container_width=True):
        benar_pin = str(st.secrets.get("APP_PIN", ""))
        if not benar_pin:
          st.error("Konfigurasi PIN aplikasi belum diatur di server.")
        elif input_pin and hmac.compare_digest(input_pin, benar_pin):
          st.session_state.pin_valid = True
          st.session_state.login_at = time.time()
          st.session_state.failed_pin_attempts = 0
          st.session_state.pin_lock_until = 0.0
          st.session_state.session_expired_msg = ""
          st.rerun()
        else:
          st.session_state.failed_pin_attempts += 1
          if st.session_state.failed_pin_attempts >= MAX_PIN_ATTEMPTS:
            st.session_state.pin_lock_until = time.time() + LOCKOUT_SECONDS
            st.session_state.failed_pin_attempts = 0
            st.rerun()
          else:
            sisa_percobaan = (
                MAX_PIN_ATTEMPTS - st.session_state.failed_pin_attempts
            )
            st.error(
                f"PIN salah. Sisa kesempatan percobaan: **{sisa_percobaan} kali**."
            )
  st.stop()

# --- HEADER UTAMA ---
with st.container(border=True):
  st.markdown("## 📋 Tracking Timeline Pengadaan Berdasarkan Nomor PR")
  st.caption(
      "Gantt Chart Jadwal Proses Pengadaan (Plan & Realisasi) • Fungsi Asset Management"
  )


def _sanitize_text(val, max_len=250):
  """Membersihkan teks dinamis dari Excel agar aman dari injeksi HTML/Markdown."""
  if pd.isna(val):
    return ""
  text = str(val).replace("\n", " ").replace("\r", " ")
  text = re.sub(r"[\x00-\x1f\x7f]", "", text).strip()
  text = html.escape(text, quote=True)
  if len(text) > max_len:
    text = text[:max_len].rstrip() + "..."
  return text


def _parse_tanggal_fleksibel(series):
  """Parsing tanggal standar dari Excel maupun teks tanggal bahasa Inggris."""

  def _norm_str(val):
    if pd.isna(val):
      return val
    if isinstance(val, str):
      s = val.strip()
      return pd.NaT if (s == "" or s.lower() == "none") else s
    return val

  return pd.to_datetime(series.apply(_norm_str), errors="coerce")


def _format_bulan_tahun_en(period_obj):
  """Mengubah pd.Period('M') menjadi format 'January 2026'."""
  return period_obj.strftime("%B %Y") if pd.notna(period_obj) else ""


def _clean_id_str(val):
  """Mengubah angka/teks ID (Nomor PR / Nomor PO) menjadi string bersih tanpa '.0'."""
  if pd.isna(val) or str(val).strip() == "":
    return ""
  num_val = pd.to_numeric(val, errors="coerce")
  if pd.notna(num_val):
    return str(int(round(num_val)))
  s = _sanitize_text(val, max_len=64)
  return s[:-2] if s.endswith(".0") else s


def _format_usd(val):
  """Format nilai numerik ke format mata uang."""
  num_val = pd.to_numeric(val, errors="coerce")
  return f"USD {num_val:,.2f}" if pd.notna(num_val) else ""


def _is_none_text(val):
  """Mendeteksi apakah nilai mentah pada kolom Realisasi berisi teks 'none'."""
  if pd.isna(val):
    return False
  return str(val).strip().lower() in {"none", "nan", "nat"}


def _validate_excel_upload(uploaded_file):
  """Memvalidasi ukuran, nama ekstensi, dan signature file Excel."""
  file_size = getattr(uploaded_file, "size", 0)
  if file_size <= 0:
    raise ValueError("File Excel kosong (0 byte).")
  if file_size > MAX_EXCEL_SIZE_BYTES:
    size_mb = file_size / (1024 * 1024)
    raise ValueError(
        f"Ukuran file ({size_mb:.2f} MB) melebihi batas maksimal {MAX_EXCEL_SIZE_MB} MB."
    )

  file_name = str(getattr(uploaded_file, "name", "")).lower()
  if not file_name.endswith((".xlsx", ".xls")):
    raise ValueError("Format file tidak didukung. Gunakan file .xlsx atau .xls.")

  # Cek magic bytes file Excel (.xlsx = ZIP PK\x03\x04, .xls = OLE2 \xD0\xCF\x11\xE0)
  uploaded_file.seek(0)
  header_bytes = uploaded_file.read(4)
  uploaded_file.seek(0)
  is_xlsx_sig = header_bytes == b"PK\x03\x04"
  is_xls_sig = header_bytes == b"\xd0\xcf\x11\xe0"
  if not (is_xlsx_sig or is_xls_sig):
    raise ValueError(
        "Isi file tidak valid sebagai dokumen Excel (.xlsx/.xls)."
    )


def load_schedule_data(uploaded_file):
  _validate_excel_upload(uploaded_file)
  uploaded_file.seek(0)
  xls = pd.ExcelFile(uploaded_file)

  if len(xls.sheet_names) > MAX_EXCEL_SHEETS:
    raise ValueError(
        f"Jumlah sheet ({len(xls.sheet_names)}) melebihi batas maksimal {MAX_EXCEL_SHEETS} sheet."
    )

  # 1. Baca sheet Jadwal dengan keep_default_na=False agar teks "None" tidak otomatis berubah jadi NaN
  jadwal_sheet = next(
      (s for s in xls.sheet_names if s.strip().lower() == "jadwal"),
      xls.sheet_names[0],
  )
  df = pd.read_excel(
      xls,
      sheet_name=jadwal_sheet,
      keep_default_na=False,
      nrows=MAX_EXCEL_ROWS + 1,
  )
  if len(df) > MAX_EXCEL_ROWS:
    raise ValueError(
        f"Jumlah baris melebihi batas maksimal {MAX_EXCEL_ROWS:,} baris."
    )

  # 2. Baca mapping Nomor PR -> Header, Nomor PO, Nilai PR, Nilai PO dari sheet "Nomor PR"
  pr_info_map = {}
  pr_sheet = next(
      (s for s in xls.sheet_names if s.strip().lower() == "nomor pr"), None
  )
  if pr_sheet is not None:
    df_pr_sheet = pd.read_excel(
        xls, sheet_name=pr_sheet, nrows=MAX_EXCEL_ROWS
    )
    df_pr_sheet.columns = df_pr_sheet.columns.astype(str).str.strip()
    header_col = next(
        (c for c in df_pr_sheet.columns if c.strip().lower() == "header"),
        None,
    )
    if "Nomor PR" in df_pr_sheet.columns:
      for _, r_pr in df_pr_sheet.dropna(subset=["Nomor PR"]).iterrows():
        k = _clean_id_str(r_pr["Nomor PR"])
        if k:
          raw_header = r_pr.get(header_col) if header_col else ""
          clean_header = _sanitize_text(raw_header, max_len=200)
          n_pr_num = pd.to_numeric(r_pr.get("Nilai PR"), errors="coerce")
          n_po_num = pd.to_numeric(r_pr.get("Nilai PO"), errors="coerce")
          pr_info_map[k] = {
              "header": clean_header,
              "no_po": _clean_id_str(r_pr.get("Nomor PO")),
              "nilai_pr": _format_usd(n_pr_num),
              "nilai_po": _format_usd(n_po_num),
              "nilai_pr_num": float(n_pr_num) if pd.notna(n_pr_num) else 0.0,
              "nilai_po_num": float(n_po_num) if pd.notna(n_po_num) else 0.0,
          }

  # 3. Normalisasi nama & isi kolom (abaikan kolom catatan/Unnamed)
  df.columns = df.columns.astype(str).str.strip()
  valid_cols = [
      c
      for c in df.columns
      if not c.lower().startswith("unnamed")
      and "warna kuning" not in c.lower()
  ]
  df = df[valid_cols].copy()

  col_rules = [
      (("realisasi", "actual", "aktual"), "Realisasi_Input"),
      (("pr", "nomor"), "Nomor PR"),
      (("deskripsi", "tahapan"), "Deskripsi"),
      (("hari", "duration"), "Hari"),
      (("plan", "tanggal", "mulai"), "Plan_Input"),
  ]
  rename_map = {
      col: target
      for col in df.columns
      for keys, target in col_rules
      if any(k in col.lower() for k in keys)
  }
  df = df.rename(columns=rename_map)
  if "Realisasi_Input" not in df.columns:
    df["Realisasi_Input"] = ""

  if "Nomor PR" not in df.columns or "Deskripsi" not in df.columns:
    raise ValueError(
        "Kolom wajib ('Nomor PR' atau 'Deskripsi') tidak ditemukan pada sheet Jadwal."
    )

  df["Nomor PR"] = df["Nomor PR"].apply(_clean_id_str)
  df = df[df["Nomor PR"] != ""].copy()
  df["Deskripsi"] = df["Deskripsi"].apply(
      lambda v: _sanitize_text(v, max_len=180)
  )
  df["Hari"] = pd.to_numeric(df["Hari"], errors="coerce").fillna(1).astype(int)
  df["Plan_Input"] = _parse_tanggal_fleksibel(df["Plan_Input"])

  # Tandai sel Realisasi yang secara eksplisit berisi teks "None" / "none"
  df["Realisasi_Is_None"] = df["Realisasi_Input"].apply(_is_none_text)
  df["Realisasi_Input"] = _parse_tanggal_fleksibel(df["Realisasi_Input"])

  # 4. Hitung tanggal kumulatif Plan & evaluasi posisi realisasi per Nomor PR
  processed_rows = []
  for pr_num, group in df.groupby("Nomor PR", sort=False):
    group = group.reset_index(drop=True)
    pra_tender_mask = group["Deskripsi"].str.lower().str.contains("pra tender")
    plan_a_pra_tender = (
        group.loc[pra_tender_mask, "Plan_Input"].dropna().iloc[0]
        if pra_tender_mask.any()
        and not group.loc[pra_tender_mask, "Plan_Input"].dropna().empty
        else group.loc[0, "Plan_Input"]
    )
    plan_start = (
        plan_a_pra_tender
        if pd.notna(plan_a_pra_tender)
        else pd.to_datetime("2026-11-20")
    )
    temp_rows = []

    for idx, row in group.iterrows():
      plan_end = plan_start + pd.Timedelta(days=int(row["Hari"]))
      desc_lower = str(row["Deskripsi"]).lower()

      if "u_delivery" in desc_lower or "delivery time" in desc_lower:
        milestone_plan = "📦 Delivery Time PO"
      elif "v_goods receipt" in desc_lower or "goods receipt" in desc_lower:
        milestone_plan = "📥 Goods Receipt (GR)"
      elif idx == 0:
        milestone_plan = "🏁 Mulai (Pra Tender)"
      else:
        milestone_plan = "⏳ Berjalan"

      real_tgl = row["Realisasi_Input"]
      if pd.notna(real_tgl):
        if real_tgl < plan_start:
          posisi_real, status_bar = "Di Luar Bar (Lebih Awal / Kiri)", "hijau"
        elif real_tgl > plan_end:
          posisi_real, status_bar = "Di Luar Bar (Lewat / Kanan)", "merah"
        else:
          posisi_real, status_bar = "Di Dalam Bar Plan", "hijau"
      else:
        posisi_real, status_bar = "-", None

      new_row = row.copy()
      new_row["Plan_A_Pra_Tender"] = plan_a_pra_tender
      new_row["Plan"] = plan_start
      new_row["Selesai"] = plan_end
      new_row["Realisasi_Tgl"] = real_tgl
      new_row["Posisi_Realisasi"] = posisi_real
      new_row["Status_Posisi_Bar"] = status_bar
      new_row["Milestone"] = milestone_plan
      temp_rows.append(new_row)
      plan_start = plan_end

    df_group_calc = pd.DataFrame(temp_rows)

    # Acuan khusus baris T_Penerbitan PO / No PO
    po_mask = (
        df_group_calc["Deskripsi"]
        .str.lower()
        .str.contains(
            r"t_penerbitan\s*po|penerbitan\s*po|no\s*po", regex=True
        )
    )
    if po_mask.any():
      po_row = df_group_calc.loc[po_mask].iloc[-1]
      tgl_plan_po = (
          po_row["Plan_Input"]
          if pd.notna(po_row["Plan_Input"])
          else po_row["Plan"]
      )
      tgl_akhir_po = po_row["Selesai"]
      tgl_real_po = po_row["Realisasi_Tgl"]
      tgl_target_po = tgl_real_po if pd.notna(tgl_real_po) else tgl_akhir_po
    else:
      tgl_plan_po = tgl_akhir_po = tgl_real_po = tgl_target_po = pd.NaT

    durasi_texts = []
    for _, r_m in df_group_calc.iterrows():
      r_tgl = r_m["Realisasi_Tgl"]
      if pd.isna(r_tgl) or pd.isna(tgl_plan_po):
        durasi_texts.append("-")
      elif r_tgl <= tgl_plan_po:
        sisa_hr = int((tgl_plan_po - r_tgl).days)
        durasi_texts.append(
            f"Durasi ke Rencana Penerbitan PO adalah {sisa_hr} hari."
        )
      else:
        lewat_hr = int((r_tgl - tgl_plan_po).days)
        durasi_texts.append(
            f"Durasi dari Penerbitan PO sudah lewat {lewat_hr} hari."
        )

    info_pr = pr_info_map.get(str(pr_num).strip(), {})
    df_group_calc["Header_PR"] = info_pr.get("header", "")
    df_group_calc["Tgl_Plan_PO"] = tgl_plan_po
    df_group_calc["Tgl_Akhir_PO"] = tgl_akhir_po
    df_group_calc["Tgl_Real_PO"] = tgl_real_po
    df_group_calc["Tgl_Target_PO"] = tgl_target_po
    df_group_calc["Durasi_Hover_Text"] = durasi_texts
    df_group_calc["Nomor_PO_Terbit"] = info_pr.get("no_po", "")
    df_group_calc["Nilai_PR_USD"] = info_pr.get("nilai_pr", "")
    df_group_calc["Nilai_PO_USD"] = info_pr.get("nilai_po", "")
    df_group_calc["Nilai_PR_Num"] = info_pr.get("nilai_pr_num", 0.0)
    df_group_calc["Nilai_PO_Num"] = info_pr.get("nilai_po_num", 0.0)
    processed_rows.append(df_group_calc)

  return (
      pd.concat(processed_rows, ignore_index=True)
      if processed_rows
      else pd.DataFrame()
  )


def hitung_progres_pr(df_pr_single):
  """Menghitung tahapan tercapai & persentase progres satu PR secara konsisten."""
  grp_reset = df_pr_single.reset_index(drop=True)
  total_tahapan = len(grp_reset)
  if total_tahapan == 0:
    return 0, 0, 0.0

  last_valid_idx = grp_reset["Realisasi_Tgl"].last_valid_index()
  if last_valid_idx is None:
    return 0, total_tahapan, 0.0

  tahapan_tercapai = int(last_valid_idx) + 1
  persen_progres = (tahapan_tercapai / total_tahapan) * 100.0
  return tahapan_tercapai, total_tahapan, persen_progres


# --- SIDEBAR: UPLOAD & FILTER PR ---
with st.sidebar:
  st.header("📂 Data Source")
  waktu_kadaluarsa = datetime.fromtimestamp(
      st.session_state.login_at + SESSION_TIMEOUT_SECONDS, tz=TZ_WIB
  ).strftime("%H:%M:%S WIB")
  st.caption(f"⏱️ Sesi kadaluarsa pukul: **{waktu_kadaluarsa}**")

  uploaded_file = st.file_uploader(
      f"Upload File Excel - Jadwal (Maks. {MAX_EXCEL_SIZE_MB} MB)",
      type=["xlsx", "xls"],
      key=f"excel_uploader_{st.session_state.uploader_key}",
  )
  if uploaded_file is not None:
    if st.button("🧹 Bersihkan Data Unggahan", use_container_width=True):
      _clear_uploaded_data()
      st.rerun()

  st.divider()
  if st.button("🚪 Kunci / Logout PIN", use_container_width=True):
    _logout_session(expired=False)
    st.rerun()

if uploaded_file is not None:
  try:
    df_all = load_schedule_data(uploaded_file)
    daftar_pr = sorted(df_all["Nomor PR"].unique())
    selected_pr = st.sidebar.multiselect(
        "Pilih Nomor PR:",
        options=daftar_pr,
        default=daftar_pr[:1] if daftar_pr else [],
        key="multiselect_nomor_pr",
    )

    if selected_pr:
      df_filtered = df_all[df_all["Nomor PR"].isin(selected_pr)].copy()

      tab_aktivitas, tab_tabel, tab_gr_kurva = st.tabs([
          "📈 Jadwal Aktivitas",
          "📄 Resume Tabel Jadwal",
          "📊 Plan vs Realisasi Goods Receipt",
      ])

      # --- TAB 1: JADWAL AKTIVITAS (GANTT CHART & KARTU RINGKASAN) ---
      with tab_aktivitas:
        with st.container(border=True):
          fig = px.timeline(
              df_filtered,
              x_start="Plan",
              x_end="Selesai",
              y="Deskripsi",
              color="Milestone",
              color_discrete_map={
                  "🏁 Mulai (Pra Tender)": "#059669",
                  "⏳ Berjalan": "#d97706",
                  "📦 Delivery Time PO": "#e11d48",
                  "📥 Goods Receipt (GR)": "#2563eb",
              },
              title="<b>Gantt Chart Tracking Milestone PR (Plan & Realisasi)</b>",
              hover_data={
                  "Nomor PR": True,
                  "Hari": True,
                  "Plan": "|%d %b %Y",
                  "Selesai": "|%d %b %Y",
                  "Milestone": False,
              },
              labels={
                  "Deskripsi": "Tahapan Pekerjaan",
                  "Plan": "Tanggal Mulai",
                  "Selesai": "Tanggal Selesai",
              },
          )

          df_markers = df_filtered.dropna(subset=["Realisasi_Tgl"]).copy()
          if not df_markers.empty:
            df_markers["Tgl_Str"] = df_markers["Realisasi_Tgl"].dt.strftime(
                "%d %b %Y"
            )
            df_markers["PO_Str"] = df_markers["Tgl_Plan_PO"].apply(
                lambda d: (
                    d.strftime("%d %b %Y")
                    if pd.notna(d)
                    else "Belum ada Plan PO"
                )
            )

            jml_hijau = int((df_markers["Status_Posisi_Bar"] == "hijau").sum())
            jml_merah = int((df_markers["Status_Posisi_Bar"] == "merah").sum())
            total_hm = jml_hijau + jml_merah

            pct_hijau = (jml_hijau / total_hm * 100.0) if total_hm > 0 else 0.0
            pct_merah = (jml_merah / total_hm * 100.0) if total_hm > 0 else 0.0

            marker_configs = [
                (
                    df_markers[df_markers["Status_Posisi_Bar"] == "hijau"],
                    f"Realisasi (Lebih Awal atau Tepat) — {pct_hijau:.2f}%",
                    "#16a34a",
                    "Titik Realisasi (On Track / Lebih Awal)",
                ),
                (
                    df_markers[df_markers["Status_Posisi_Bar"] == "merah"],
                    f"Realisasi (Lebih Lambat) — {pct_merah:.2f}%",
                    "#dc2626",
                    "Titik Realisasi (Lewat Bar Plan)",
                ),
            ]
            for df_sub, name_lbl, clr, hover_title in marker_configs:
              if not df_sub.empty:
                fig.add_trace(
                    go.Scatter(
                        x=df_sub["Realisasi_Tgl"],
                        y=df_sub["Deskripsi"],
                        mode="markers",
                        name=name_lbl,
                        marker=dict(
                            color=clr,
                            size=9,
                            line=dict(color="#0f172a", width=1.1),
                        ),
                        customdata=list(
                            zip(
                                df_sub["Nomor PR"],
                                df_sub["Tgl_Str"],
                                df_sub["Durasi_Hover_Text"],
                                df_sub["PO_Str"],
                            )
                        ),
                        hovertemplate=(
                            f"<b>{hover_title}</b><br>"
                            "Nomor PR: %{customdata[0]}<br>"
                            "Tahapan: %{y}<br>"
                            "Tanggal: <b>%{customdata[1]}</b><br>"
                            "<b>%{customdata[2]}</b> (Plan PO:"
                            " %{customdata[3]})<extra></extra>"
                        ),
                    )
                )

          urutan_deskripsi = (
              df_filtered[df_filtered["Nomor PR"] == selected_pr[0]][
                  "Deskripsi"
              ]
              .drop_duplicates()
              .tolist()
          )
          fig.update_yaxes(
              categoryorder="array",
              categoryarray=urutan_deskripsi,
              autorange="reversed",
          )
          fig.update_layout(
              paper_bgcolor="#ffffff",
              plot_bgcolor="#f8fafc",
              font=dict(size=11, color="#0f172a"),
              margin=dict(t=60, b=45, l=15, r=15),
              xaxis=dict(
                  showgrid=True, gridcolor="#e2e8f0", tickformat="%d %b %Y"
              ),
              yaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
              legend=dict(
                  bgcolor="#ffffff",
                  bordercolor="#e2e8f0",
                  borderwidth=1,
                  title="Status Milestone & Realisasi",
              ),
          )
          st.plotly_chart(
              fig, use_container_width=True, key="gantt_jadwal_aktivitas"
          )

        # --- RINGKASAN INFO TANGGAL & PROGRES PR ---
        st.markdown("### 📌 Informasi Ringkasan Tanggal Proses PR")

        pr_progres_map = {}
        pr_terbit_po_count = 0
        total_pr_created = 0.0
        total_po_created = 0.0

        for pr_key, grp_all in df_all.groupby("Nomor PR", sort=False):
          grp_reset = grp_all.reset_index(drop=True)
          t_capai, t_total, pct_single = hitung_progres_pr(grp_reset)
          pr_progres_map[pr_key] = (t_capai, t_total, pct_single)

          first_row = grp_reset.iloc[0]
          total_pr_created += float(first_row.get("Nilai_PR_Num", 0.0) or 0.0)
          total_po_created += float(first_row.get("Nilai_PO_Num", 0.0) or 0.0)

          if pd.notna(first_row.get("Tgl_Real_PO")) and bool(
              str(first_row.get("Nomor_PO_Terbit", "")).strip()
          ):
            pr_terbit_po_count += 1

        total_pr_all = len(pr_progres_map)
        overall_progres_pct = (
            sum(item[2] for item in pr_progres_map.values()) / total_pr_all
            if total_pr_all > 0
            else 0.0
        )

        m1, m2, m3, m4 = st.columns([0.85, 0.85, 1.15, 1.15])
        m1.metric(
            "📊 Overall Progress PR",
            f"{overall_progres_pct:.2f}%",
            f"Total {total_pr_all} PR",
        )
        m2.metric(
            "✅ PR Terbit PO",
            f"{pr_terbit_po_count} PR",
            f"dari {total_pr_all} PR",
        )
        m3.metric("💵 Total PR Created", f"USD {total_pr_created:,.2f}")
        m4.metric("🧾 Total PO Created", f"USD {total_po_created:,.2f}")

        st.write("")
        hari_ini = pd.Timestamp.now().normalize()

        for pr in selected_pr:
          df_pr_group = df_filtered[df_filtered["Nomor PR"] == pr].reset_index(
              drop=True
          )
          if df_pr_group.empty:
            continue

          first_pr_row = df_pr_group.iloc[0]
          tgl_mulai_dt = first_pr_row["Plan"]
          tgl_mulai_total = tgl_mulai_dt.strftime("%d %B %Y")
          tgl_selesai_total = df_pr_group.iloc[-1]["Selesai"].strftime(
              "%d %B %Y"
          )

          pr_safe = _sanitize_text(pr, max_len=64)
          header_pr_str = _sanitize_text(
              first_pr_row.get("Header_PR", ""), max_len=200
          )
          tgl_plan_po = first_pr_row["Tgl_Plan_PO"]
          tgl_po_plan = first_pr_row["Tgl_Akhir_PO"]
          tgl_po_real = first_pr_row["Tgl_Real_PO"]
          no_po_terbit = _sanitize_text(
              first_pr_row.get("Nomor_PO_Terbit", ""), max_len=64
          )
          nilai_pr_str = str(first_pr_row.get("Nilai_PR_USD", "")).strip()
          nilai_po_str = str(first_pr_row.get("Nilai_PO_USD", "")).strip()

          total_hari = (
              (tgl_po_plan - tgl_mulai_dt).days
              if pd.notna(tgl_po_plan)
              else (df_pr_group.iloc[-1]["Selesai"] - tgl_mulai_dt).days
          )

          if pd.notna(tgl_po_real):
            tgl_po_str = (
                f"**Real Terbit PO: {tgl_po_real.strftime('%d %B %Y')}** (Plan:"
                f" {tgl_po_plan.strftime('%d %B %Y')})"
                if pd.notna(tgl_po_plan)
                else f"**Real Terbit PO: {tgl_po_real.strftime('%d %B %Y')}**"
            )
          elif pd.notna(tgl_po_plan):
            tgl_po_str = (
                f"Plan Terbit PO: {tgl_po_plan.strftime('%d %B %Y')} (Belum"
                " Realisasi)"
            )
          else:
            tgl_po_str = "Tahapan T_Penerbitan PO / No PO tidak ditemukan"

          delivery_mask = (
              df_pr_group["Deskripsi"]
              .str.lower()
              .str.contains(
                  r"u_delivery|u_delivered|21_delivered|delivery\s*time",
                  regex=True,
              )
          )
          if delivery_mask.any():
            tgl_real_delivery = df_pr_group.loc[
                delivery_mask, "Realisasi_Tgl"
            ].iloc[-1]
            tgl_real_delivery_str = (
                tgl_real_delivery.strftime("%d %B %Y")
                if pd.notna(tgl_real_delivery)
                else "Masih On Progres Proses Pengadaan"
            )
          else:
            tgl_real_delivery_str = "Baris U_Delivery Time Sesuai PO tidak ada"

          gr_mask = (
              df_pr_group["Deskripsi"]
              .str.lower()
              .str.contains(r"v_goods\s*receipt|goods\s*receipt", regex=True)
          )
          if gr_mask.any():
            tgl_real_gr = df_pr_group.loc[gr_mask, "Realisasi_Tgl"].iloc[-1]
            tgl_real_gr_str = (
                tgl_real_gr.strftime("%d %B %Y")
                if pd.notna(tgl_real_gr)
                else "Belum Goods Receipt"
            )
          else:
            tgl_real_gr_str = "Baris V_Goods Receipt tidak ada"

          tahapan_tercapai_seri, total_tahapan_pr, persen_progres_pr = (
              pr_progres_map.get(pr, hitung_progres_pr(df_pr_group))
          )
          last_valid_idx = df_pr_group["Realisasi_Tgl"].last_valid_index()

          if tahapan_tercapai_seri > 0 and total_tahapan_pr > 0:
            status_progres_line = (
                "**Status Progres dari Pra Tender sampai Goods Receipt :"
                f" {persen_progres_pr:.2f}%** ({tahapan_tercapai_seri} dari"
                f" {total_tahapan_pr} tahapan)"
            )
          else:
            status_progres_line = (
                "**Status Progres 0.00%** (Belum ada tanggal realisasi diinput)"
            )

          po_sudah_terbit = pd.notna(tgl_po_real) or (
              persen_progres_pr >= 90.0 and bool(no_po_terbit)
          )

          with st.container(border=True):
            col_head, col_prog = st.columns([3, 1.2])
            with col_head:
              judul_pr = f"#### 📌 Nomor PR: {pr_safe}"
              if header_pr_str:
                judul_pr += f" — {header_pr_str}"
              st.markdown(judul_pr)
            with col_prog:
              st.progress(
                  min(max(persen_progres_pr / 100.0, 0.0), 1.0),
                  text=f"Progres: {persen_progres_pr:.1f}%",
              )

            baris_info = []
            if nilai_pr_str:
              baris_info.append(f"**Nilai PR (USD): {nilai_pr_str}**")

            baris_info.extend([
                status_progres_line,
                f"• Tanggal Mulai Proses (Pra Tender): **{tgl_mulai_total}**",
                f"• Tanggal Target Plan Goods Receipt: **{tgl_selesai_total}**",
                (
                    "• Planned Durasi Proses (Pra Tender - Terbit PO):"
                    f" **{total_hari}** hari"
                ),
            ])

            if (
                po_sudah_terbit
                and pd.notna(tgl_po_real)
                and pd.notna(tgl_mulai_dt)
            ):
              total_hari_real = int((tgl_po_real - tgl_mulai_dt).days)
              baris_info.append(
                  "• Realisasi Durasi Proses (Pra Tender - Terbit PO):"
                  f" **{total_hari_real}** hari"
              )

            if persen_progres_pr >= 90.0 and no_po_terbit:
              baris_info.append(f"**Nomor PO: {no_po_terbit}**")
              nilai_po_tampil = (
                  nilai_po_str if nilai_po_str else "Nilai PO belum diinput"
              )
              baris_info.append(f"**Nilai PO (USD): {nilai_po_tampil}**")

            baris_info.extend([
                f"• {tgl_po_str}",
                f"• **Delivery Time Sesuai PO: {tgl_real_delivery_str}**",
                f"• **Goods Receipt (V_Goods Receipt): {tgl_real_gr_str}**",
            ])

            if po_sudah_terbit:
              st.success("  \n".join(baris_info))
            else:
              st.info("  \n".join(baris_info))

            # Evaluasi warning
            is_none_mask = (
                df_pr_group["Realisasi_Is_None"]
                if "Realisasi_Is_None" in df_pr_group.columns
                else pd.Series([False] * len(df_pr_group))
            )
            sudah_terlewati_mask = (
                df_pr_group.index < int(last_valid_idx)
                if last_valid_idx is not None
                else pd.Series([False] * len(df_pr_group))
            )
            df_pr_warn = df_pr_group[
                ~(is_none_mask | sudah_terlewati_mask)
            ].copy()

            plan_a_dt = first_pr_row["Plan_A_Pra_Tender"]
            plan_a_norm = (
                plan_a_dt.normalize() if pd.notna(plan_a_dt) else pd.NaT
            )
            selisih_hari_a = (
                int((hari_ini - plan_a_norm).days)
                if pd.notna(plan_a_norm)
                else None
            )

            belum_real_sampai_hari_ini = df_pr_warn[
                df_pr_warn["Realisasi_Tgl"].isna()
                & df_pr_warn["Plan"].notna()
                & (df_pr_warn["Plan"].dt.normalize() <= hari_ini)
            ]

            tahapan_belum_muncul = []
            for _, r_w in belum_real_sampai_hari_ini.iterrows():
              desc_bersih = " ".join(str(r_w.get("Deskripsi", "")).split())
              plan_tgl_str = (
                  r_w["Plan"].strftime("%d %b %Y")
                  if pd.notna(r_w["Plan"])
                  else "-"
              )
              if desc_bersih:
                tahapan_belum_muncul.append(
                    f"{desc_bersih} (Plan: {plan_tgl_str})"
                )

            warning_lines = []
            tampilkan_warning = False

            if (
                pd.notna(plan_a_norm)
                and selisih_hari_a is not None
                and selisih_hari_a >= 0
            ):
              tampilkan_warning = True
              warning_lines.append(
                  "• Selisih tanggal hari ini"
                  f" ({hari_ini.strftime('%d %b %Y')}) terhadap Plan A_Pra Tender:"
                  f" **{selisih_hari_a} hari**"
              )
              if tahapan_belum_muncul:
                warning_lines.append(
                    "• Tahapan yang belum terealisasi sampai hari ini"
                    " (dibandingkan tanggal Plan): **"
                    + ", ".join(tahapan_belum_muncul)
                    + "**"
                )

            if pd.notna(tgl_po_real) and pd.notna(tgl_plan_po):
              selisih_po = int((tgl_po_real - tgl_plan_po).days)
              if selisih_po > 0:
                tampilkan_warning = True
                warning_lines.append(
                    "• Selisih Realisasi PO terhadap Plan PO:"
                    f" **+{selisih_po} hari** (lewat dari Plan PO)"
                )

            if tampilkan_warning:
              st.error("  \n".join(warning_lines))

      # --- TAB 2: RESUME TABEL JADWAL ---
      with tab_tabel:
        with st.container(border=True):
          st.subheader("📄 Detail Jadwal Tahapan untuk PR")
          df_display = df_filtered[[
              "Nomor PR",
              "Deskripsi",
              "Hari",
              "Plan",
              "Selesai",
              "Realisasi_Tgl",
              "Posisi_Realisasi",
              "Durasi_Hover_Text",
              "Milestone",
          ]].reset_index(drop=True)
          df_display.index = range(1, len(df_display) + 1)

          for col_tgl in ["Plan", "Selesai", "Realisasi_Tgl"]:
            df_display[col_tgl] = (
                pd.to_datetime(df_display[col_tgl], errors="coerce")
                .dt.strftime("%d %b %Y")
                .fillna("-")
            )

          st.dataframe(
              df_display, use_container_width=True, key="tabel_detail_pr"
          )

      # --- TAB 3: GRAFIK AKUMULASI PLAN VS REALISASI V_GOODS RECEIPT ---
      with tab_gr_kurva:
        with st.container(border=True):
          st.subheader("📈 Plan GR berdasarkan Nilai PR vs Realisasi Nilai PO")

          gr_all_mask = (
              df_all["Deskripsi"]
              .str.lower()
              .str.contains(r"v_goods\s*receipt|goods\s*receipt", regex=True)
          )
          df_gr_all = (
              df_all.loc[gr_all_mask]
              .drop_duplicates(subset=["Nomor PR"], keep="last")
              .copy()
          )

          df_pr_unique_all = df_all.drop_duplicates(
              subset=["Nomor PR"], keep="first"
          )
          total_nilai_pr_all = float(df_pr_unique_all["Nilai_PR_Num"].sum())

          if df_gr_all.empty or total_nilai_pr_all <= 0:
            st.info(
                "Data V_Goods Receipt atau Nilai PR belum tersedia untuk"
                " menampilkan grafik akumulasi."
            )
          else:
            df_gr_all["Tgl_Plan_GR"] = df_gr_all["Selesai"].fillna(
                df_gr_all["Plan"]
            )
            df_gr_plan_valid = df_gr_all.dropna(subset=["Tgl_Plan_GR"]).copy()
            df_gr_real_valid = df_gr_all.dropna(subset=["Realisasi_Tgl"]).copy()

            df_gr_plan_valid["Bulan_Period"] = df_gr_plan_valid[
                "Tgl_Plan_GR"
            ].dt.to_period("M")
            df_gr_real_valid["Bulan_Period"] = df_gr_real_valid[
                "Realisasi_Tgl"
            ].dt.to_period("M")

            all_periods = []
            if not df_gr_plan_valid.empty:
              all_periods.extend(df_gr_plan_valid["Bulan_Period"].tolist())
            if not df_gr_real_valid.empty:
              all_periods.extend(df_gr_real_valid["Bulan_Period"].tolist())

            if not all_periods:
              st.info(
                  "Tanggal Plan maupun Realisasi V_Goods Receipt belum terisi."
              )
            else:
              min_period = min(all_periods)
              max_period = max(all_periods)
              period_range = pd.period_range(
                  start=min_period, end=max_period, freq="M"
              )

              plan_monthly = (
                  df_gr_plan_valid.groupby("Bulan_Period")["Nilai_PR_Num"]
                  .sum()
                  .reindex(period_range, fill_value=0.0)
              )
              real_monthly = (
                  df_gr_real_valid.groupby("Bulan_Period")["Nilai_PO_Num"]
                  .sum()
                  .reindex(period_range, fill_value=0.0)
              )

              plan_cum = plan_monthly.cumsum()
              real_cum = real_monthly.cumsum()

              max_real_period = (
                  df_gr_real_valid["Bulan_Period"].max()
                  if not df_gr_real_valid.empty
                  else None
              )

              bulan_labels = [_format_bulan_tahun_en(p) for p in period_range]
              plan_cum_vals = [float(plan_cum.loc[p]) for p in period_range]
              real_cum_vals = [
                  float(real_cum.loc[p])
                  if (max_real_period is not None and p <= max_real_period)
                  else None
                  for p in period_range
              ]

              plan_pct_vals = [
                  (val / total_nilai_pr_all) * 100 for val in plan_cum_vals
              ]
              real_pct_vals = [
                  (val / total_nilai_pr_all) * 100 if val is not None else None
                  for val in real_cum_vals
              ]

              fig_gr = go.Figure()
              fig_gr.add_trace(
                  go.Scatter(
                      x=bulan_labels,
                      y=plan_cum_vals,
                      mode="lines+markers+text",
                      name="Plan V_Goods Receipt (Akumulasi Nilai PR)",
                      line=dict(color="#2563eb", width=3),
                      marker=dict(size=8),
                      text=[f"{pct:.1f}%" for pct in plan_pct_vals],
                      textposition="top center",
                      customdata=list(
                          zip(
                              [
                                  float(plan_monthly.loc[p])
                                  for p in period_range
                              ],
                              plan_pct_vals,
                          )
                      ),
                      hovertemplate=(
                          "<b>Plan V_Goods Receipt</b><br>"
                          "Bulan: %{x}<br>"
                          "Nilai PR Bulan Ini: <b>USD %{customdata[0]:,.2f}</b><br>"
                          "Akumulasi Nilai PR: <b>USD %{y:,.2f}</b><br>"
                          "Persentase vs Total PR: <b>%{customdata[1]:.2f}%</b>"
                          "<extra></extra>"
                      ),
                  )
              )

              fig_gr.add_trace(
                  go.Scatter(
                      x=bulan_labels,
                      y=real_cum_vals,
                      mode="lines+markers+text",
                      name="Realisasi V_Goods Receipt (Akumulasi Nilai PO)",
                      line=dict(color="#16a34a", width=3),
                      marker=dict(size=8),
                      connectgaps=False,
                      text=[
                          f"{pct:.1f}%" if pct is not None else ""
                          for pct in real_pct_vals
                      ],
                      textposition="bottom center",
                      customdata=list(
                          zip(
                              [
                                  float(real_monthly.loc[p])
                                  for p in period_range
                              ],
                              [
                                  pct if pct is not None else 0.0
                                  for pct in real_pct_vals
                              ],
                          )
                      ),
                      hovertemplate=(
                          "<b>Realisasi V_Goods Receipt</b><br>"
                          "Bulan: %{x}<br>"
                          "Nilai PO Bulan Ini: <b>USD %{customdata[0]:,.2f}</b><br>"
                          "Akumulasi Nilai PO: <b>USD %{y:,.2f}</b><br>"
                          "Persentase vs Total PR: <b>%{customdata[1]:.2f}%</b>"
                          "<extra></extra>"
                      ),
                  )
              )

              fig_gr.update_layout(
                  xaxis_title=" ",
                  yaxis_title="Akumulasi Nilai (USD)",
                  yaxis=dict(
                      tickprefix="USD ",
                      tickformat=",.0f",
                      showgrid=True,
                      gridcolor="#e2e8f0",
                  ),
                  xaxis=dict(showgrid=True, gridcolor="#e2e8f0"),
                  plot_bgcolor="white",
                  hovermode="x unified",
                  legend=dict(
                      orientation="h",
                      yanchor="bottom",
                      y=1.02,
                      xanchor="right",
                      x=1,
                  ),
                  margin=dict(t=80, b=50),
              )
              st.plotly_chart(
                  fig_gr, use_container_width=True, key="chart_akumulasi_gr"
              )

              poin1_lines = ["**Daftar Plan Goods Receipt**"]
              for p in period_range:
                lbl_bln = _format_bulan_tahun_en(p)
                items_pr = df_gr_plan_valid[
                    df_gr_plan_valid["Bulan_Period"] == p
                ]
                if not items_pr.empty:
                  poin1_lines.append("")
                  poin1_lines.append(f"**{lbl_bln}**")
                  for idx_item, (_, r_item) in enumerate(
                      items_pr.iterrows(), start=1
                  ):
                    no_pr_item = str(r_item.get("Nomor PR", "")).strip()
                    hdr_item = str(r_item.get("Header_PR", "")).strip()
                    judul_pengadaan = (
                        hdr_item if hdr_item else "(Judul tidak tersedia)"
                    )
                    poin1_lines.append(
                        f"{idx_item}. Nomor PR {no_pr_item} pengadaan"
                        f" {judul_pengadaan}"
                    )

              st.info("  \n".join(poin1_lines))

              poin2_lines = [
                  "**Realisasi Goods Receipt tanggal"
                  f" {hari_ini.strftime('%d %B %Y')}**",
              ]
              ada_item_gr = False
              for p in period_range:
                lbl_bln = _format_bulan_tahun_en(p)
                items_gr = df_gr_real_valid[
                    df_gr_real_valid["Bulan_Period"] == p
                ]
                if not items_gr.empty:
                  ada_item_gr = True
                  poin2_lines.append("")
                  poin2_lines.append(f"**{lbl_bln}**")
                  for idx_item, (_, r_item) in enumerate(
                      items_gr.iterrows(), start=1
                  ):
                    no_po_item = (
                        str(r_item.get("Nomor_PO_Terbit", "")).strip() or "-"
                    )
                    no_pr_item = str(r_item.get("Nomor PR", "")).strip()
                    hdr_item = str(r_item.get("Header_PR", "")).strip()
                    judul_pengadaan = (
                        hdr_item if hdr_item else "(Judul tidak tersedia)"
                    )
                    poin2_lines.append(
                        f"{idx_item}. Nomor PO {no_po_item} atas Nomor PR"
                        f" {no_pr_item} pengadaan {judul_pengadaan}"
                    )

              if not ada_item_gr:
                poin2_lines.append("")
                poin2_lines.append(
                    "Belum ada Nomor PO yang terealisasi Goods Receipt."
                )

              st.success("  \n".join(poin2_lines))
    else:
      st.warning("Silakan pilih minimal satu Nomor PR di panel sebelah kiri.")

  except ValueError as ve:
    st.error(f"Validasi file gagal: {ve}")
  except Exception:
    st.error(
        "Terjadi kesalahan saat memproses file jadwal. Pastikan format dan struktur sheet sesuai template."
    )
else:
  st.info(
      "Silakan unggah file Excel jadwal pada panel **Data Source** di sebelah kiri untuk menampilkan Gantt Chart dan ringkasan PR."
  )