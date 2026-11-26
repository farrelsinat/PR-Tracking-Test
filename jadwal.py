import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Tracking Jadwal PR & Gantt Chart", layout="wide")

import hmac

# ... setelah st.set_page_config(...) ...
if "pin_valid" not in st.session_state:
  st.session_state.pin_valid = False

if not st.session_state.pin_valid:
  st.title("🔒 Verifikasi Akses")
  st.write("Masukkan PIN untuk membuka aplikasi tracking jadwal PR.")
  input_pin = st.text_input("PIN Akses", type="password")

  if st.button("Masuk"):
    benar_pin = str(st.secrets.get("APP_PIN", ""))
    if benar_pin and hmac.compare_digest(input_pin, benar_pin):
      st.session_state.pin_valid = True
      st.rerun()
    else:
      st.error("PIN salah atau belum diatur.")
  st.stop()
  
st.markdown("## 📋 Tracking Timeline Pengadaan Berdasarkan Nomor PR")
st.caption("Gantt Chart Jadwal Proses Pengadaan (Plan & Realisasi)" \
" - Fungsi Asset Management")

def _clean_id_str(val):
  """Mengubah angka/teks ID (Nomor PR / Nomor PO) menjadi string bersih tanpa '.0'."""
  if pd.isna(val):
    return ""
  num_val = pd.to_numeric(val, errors="coerce")
  if pd.notna(num_val):
    return str(int(round(num_val)))
  s = str(val).strip()
  return s[:-2] if s.endswith(".0") else s

def _format_usd(val):
  """Format nilai numerik ke format mata uang"""
  num_val = pd.to_numeric(val, errors="coerce")
  return f"USD {num_val:,.2f}" if pd.notna(num_val) else ""

def _is_none_text(val):
  """Mendeteksi apakah nilai mentah pada kolom Realisasi berisi teks 'none'."""
  if pd.isna(val):
    return False
  return str(val).strip().lower() in {"none", "nan", "nat"}

@st.cache_data
def load_schedule_data(uploaded_file):
  uploaded_file.seek(0)
  xls = pd.ExcelFile(uploaded_file)

  # 1. Baca sheet Jadwal
  jadwal_sheet = next(
      (s for s in xls.sheet_names if s.strip().lower() == "jadwal"),
      xls.sheet_names[0],
  )
  df = pd.read_excel(xls, sheet_name=jadwal_sheet)

  # 2. Baca mapping Nomor PR -> Header, Nomor PO, Nilai PR, Nilai PO dari sheet "Nomor PR"
  pr_info_map = {}
  pr_sheet = next(
      (s for s in xls.sheet_names if s.strip().lower() == "nomor pr"), None
  )
  if pr_sheet is not None:
    df_pr_sheet = pd.read_excel(xls, sheet_name=pr_sheet)
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
          clean_header = (
              ""
              if pd.isna(raw_header)
              else str(raw_header).replace("\n", " ").strip()
          )
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

  # 3. Normalisasi nama & isi kolom
  df.columns = df.columns.astype(str).str.strip()
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
    df["Realisasi_Input"] = pd.NaT

  df = df.dropna(subset=["Nomor PR"]).copy()
  df["Nomor PR"] = df["Nomor PR"].apply(_clean_id_str)
  df["Deskripsi"] = (
      df["Deskripsi"].astype(str).str.replace("\n", " ", regex=True).str.strip()
  )
  df["Hari"] = pd.to_numeric(df["Hari"], errors="coerce").fillna(1).astype(int)
  df["Plan_Input"] = pd.to_datetime(df["Plan_Input"], errors="coerce")

  # Tandai sel Realisasi yang secara eksplisit berisi teks "none"
  df["Realisasi_Is_None"] = df["Realisasi_Input"].apply(
      lambda v: str(v).strip().lower() == "none" if pd.notna(v) else False
  )
  df["Realisasi_Input"] = pd.to_datetime(
      df["Realisasi_Input"], errors="coerce"
  )

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

    # Acuan khusus baris T_Penerbitan PO / No PO (mengambil tanggal Plan sel T_Penerbitan PO)
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

    # Marker (hijau maupun merah) berpatokan pada tanggal Plan sel T_Penerbitan PO:
    # - Realisasi <= Plan PO -> "Durasi ke Penerbitan PO adalah {N} hari."
    # - Realisasi > Plan PO  -> "Durasi dari Penerbitan PO sudah lewat {N} hari."
    durasi_texts = []
    for _, r_m in df_group_calc.iterrows():
      r_tgl = r_m["Realisasi_Tgl"]
      if pd.isna(r_tgl) or pd.isna(tgl_plan_po):
        durasi_texts.append("-")
      elif r_tgl <= tgl_plan_po:
        sisa_hr = int((tgl_plan_po - r_tgl).days)
        durasi_texts.append(f"Durasi ke Penerbitan PO adalah {sisa_hr} hari.")
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

# SIDEBAR: UPLOAD & FILTER PR
st.sidebar.header("📂 Data Source")
uploaded_file = st.sidebar.file_uploader(
    "Upload File Excel - Jadwal", type=["xlsx", "xls"]
)

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

      tab_aktivitas, tab_tabel = st.tabs(
          ["Jadwal Aktivitas", "Resume Tabel Jadwal"]
      )

      # TAB 1: JADWAL AKTIVITAS (GANTT CHART & KARTU RINGKASAN)
      with tab_aktivitas:
        fig = px.timeline(
            df_filtered,
            x_start="Plan",
            x_end="Selesai",
            y="Deskripsi",
            color="Milestone",
            color_discrete_map={
                "🏁 Mulai (Pra Tender)": "#01831d",
                "⏳ Berjalan": "#afb41f",
                "📦 Delivery Time PO": "#dc2626",
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

        # Titik penanda realisasi (Hijau = di dalam / kiri bar, Merah = kanan bar)
        df_markers = df_filtered.dropna(subset=["Realisasi_Tgl"]).copy()
        if not df_markers.empty:
          df_markers["Tgl_Str"] = df_markers["Realisasi_Tgl"].dt.strftime(
              "%d %b %Y"
          )
          df_markers["PO_Str"] = df_markers["Tgl_Plan_PO"].apply(
              lambda d: (
                  d.strftime("%d %b %Y") if pd.notna(d) else "Belum ada Plan PO"
              )
          )

          marker_configs = [
              (
                  df_markers[df_markers["Status_Posisi_Bar"] == "hijau"],
                  "Realisasi (Di Dalam / Lebih Awal dari Bar)",
                  "#16a34a",
                  "Titik Realisasi (On Track / Lebih Awal)",
              ),
              (
                  df_markers[df_markers["Status_Posisi_Bar"] == "merah"],
                  "Realisasi (Lewat / Di Luar Bar Kanan)",
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
                          size=8,
                          line=dict(color="black", width=1.0),
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
            df_filtered[df_filtered["Nomor PR"] == selected_pr[0]]["Deskripsi"]
            .drop_duplicates()
            .tolist()
        )
        fig.update_yaxes(
            categoryorder="array",
            categoryarray=urutan_deskripsi,
            autorange="reversed",
        )
        fig.update_layout(
            plot_bgcolor="white",
            font=dict(size=11),
            margin=dict(t=60, b=50),
            xaxis=dict(
                showgrid=True, gridcolor="#e2e8f0", tickformat="%d %b %Y"
            ),
            legend_title="Status Milestone & Realisasi",
        )
        st.plotly_chart(
            fig, use_container_width=True, key="gantt_jadwal_aktivitas"
        )

        # RINGKASAN INFO TANGGAL & PROGRES PR
        st.markdown("### 📌 Informasi Ringkasan Tanggal Proses PR")

        total_pr_all = df_all["Nomor PR"].nunique()
        total_proporsi_progres = 0.0
        pr_terbit_po_count = 0
        total_pr_created = 0.0
        total_po_created = 0.0

        for _, grp_all in df_all.groupby("Nomor PR", sort=False):
          grp_reset = grp_all.reset_index(drop=True)
          last_idx_all = grp_reset["Realisasi_Tgl"].last_valid_index()
          if len(grp_reset) > 0 and last_idx_all is not None:
            total_proporsi_progres += (int(last_idx_all) + 1) / len(grp_reset)

          first_row = grp_reset.iloc[0]
          total_pr_created += float(first_row.get("Nilai_PR_Num", 0.0) or 0.0)
          total_po_created += float(first_row.get("Nilai_PO_Num", 0.0) or 0.0)

          if pd.notna(first_row.get("Tgl_Real_PO")) and bool(
              str(first_row.get("Nomor_PO_Terbit", "")).strip()
          ):
            pr_terbit_po_count += 1

        overall_progres_pct = (
            (total_proporsi_progres / total_pr_all) * 100
            if total_pr_all > 0
            else 0.0
        )
        st.success(
            "**RESUME**  \n"
            "• **Overall progress dari total"
            f" {total_pr_all} nomor PR dalam proses:"
            f" {overall_progres_pct:.2f}%**  \n"
            f"• **Terdapat {pr_terbit_po_count} nomor PR yang sudah terbit"
            " menjadi PO.**  \n"
            f"• **Total PR created: USD {total_pr_created:,.2f}**  \n"
            f"• **Total PO created: USD {total_po_created:,.2f}**"
        )

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

          header_pr_str = str(first_pr_row.get("Header_PR", "")).strip()
          tgl_plan_po = first_pr_row["Tgl_Plan_PO"]
          tgl_po_plan = first_pr_row["Tgl_Akhir_PO"]
          tgl_po_real = first_pr_row["Tgl_Real_PO"]
          no_po_terbit = str(first_pr_row.get("Nomor_PO_Terbit", "")).strip()
          nilai_pr_str = str(first_pr_row.get("Nilai_PR_USD", "")).strip()
          nilai_po_str = str(first_pr_row.get("Nilai_PO_USD", "")).strip()

          total_hari = (
              (tgl_po_plan - tgl_mulai_dt).days
              if pd.notna(tgl_po_plan)
              else (df_pr_group.iloc[-1]["Selesai"] - tgl_mulai_dt).days
          )

          # Status & tanggal terbit PO
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

          # Tanggal Realisasi U_Delivery Time Sesuai PO
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

          # Tanggal Realisasi V_Goods Receipt
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

          # Progres seri per PR (22 tahapan; >= 90% = 20 dari 22 tahapan -> hijau)
          total_tahapan_pr = len(df_pr_group)
          last_valid_idx = df_pr_group["Realisasi_Tgl"].last_valid_index()
          if last_valid_idx is not None and total_tahapan_pr > 0:
            tahapan_tercapai_seri = int(last_valid_idx)
            persen_progres_pr = (tahapan_tercapai_seri / total_tahapan_pr) * 100
            status_progres_line = (
                "**Status Progres dari Pra Tender - Terbit PO :"
                f" {persen_progres_pr:.2f}%** ({tahapan_tercapai_seri} dari"
                f" {total_tahapan_pr} tahapan)"
            )
          else:
            persen_progres_pr = 0.0
            status_progres_line = (
                "**Status Progres 0.00%** (Belum ada tanggal realisasi diinput)"
            )

          po_sudah_terbit = pd.notna(tgl_po_real) or (
              persen_progres_pr >= 90.0 and bool(no_po_terbit)
          )

          baris_info = [f"**Nomor PR: {pr}**"]
          if header_pr_str:
            baris_info.append(f"**Header: {header_pr_str}**")
          if nilai_pr_str:
            baris_info.append(f"**Nilai PR (USD): {nilai_pr_str}**")

          baris_info.extend([
              status_progres_line,
              f"• Tanggal Mulai Proses (Pra Tender): **{tgl_mulai_total}**",
              f"• Tanggal Target Plan Delivery: **{tgl_selesai_total}**",
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
                f" **{total_hari_real}** hari \n"
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

          # Evaluasi st.warning dengan acuan awal Plan A_Pra Tender & Plan T_Penerbitan PO
          is_none_mask = (
              df_pr_group["Realisasi_Is_None"]
              if "Realisasi_Is_None" in df_pr_group.columns
              else pd.Series([False] * len(df_pr_group))
          )
          df_pr_warn = df_pr_group[~is_none_mask].copy()

          plan_a_dt = first_pr_row["Plan_A_Pra_Tender"]
          selisih_hari_a = (
              int((hari_ini - plan_a_dt).days)
              if pd.notna(plan_a_dt)
              else None
          )
          ada_po_atau_real_po = bool(no_po_terbit) or pd.notna(tgl_po_real)
          ada_tanggal_realisasi = df_pr_warn["Realisasi_Tgl"].notna().any()
          deskripsi_kosong = (
              df_pr_warn.loc[df_pr_warn["Realisasi_Tgl"].isna(), "Deskripsi"]
              .astype(str)
              .str.strip()
              .tolist()
          )

          warning_lines = []
          tampilkan_warning = False

          if ada_po_atau_real_po and deskripsi_kosong:
            if pd.notna(tgl_po_real) and pd.notna(tgl_plan_po):
              selisih_po = int((tgl_po_real - tgl_plan_po).days)
              if selisih_po > 0:
                tampilkan_warning = True
                warning_lines.append(
                    "• Selisih Realisasi PO - Plan PO:"
                    f" **+{selisih_po} hari** (lewat dari Plan PO)"
                )
            else:
              tampilkan_warning = True
          elif ada_tanggal_realisasi and deskripsi_kosong:
            tampilkan_warning = True
          elif (not ada_po_atau_real_po) and (not ada_tanggal_realisasi):
            if selisih_hari_a is not None:
              tampilkan_warning = True

          if tampilkan_warning:
            if selisih_hari_a is not None:
              warning_lines.insert(
                  0,
                  "• Selisih tanggal hari ini"
                  f" ({hari_ini.strftime('%d %b %Y')}) terhadap Plan A_Pra Tender:"
                  f" **{selisih_hari_a} hari**",
              )
            if (
                ada_po_atau_real_po or ada_tanggal_realisasi
            ) and deskripsi_kosong:
              warning_lines.append(
                  "• Deskripsi realisasi masih kosong:"
                  f" **{', '.join(deskripsi_kosong)}**"
              )
            st.error("  \n".join(warning_lines))

      # TAB 2: RESUME TABEL JADWAL
      with tab_tabel:
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
    else:
      st.warning("Silakan pilih minimal satu Nomor PR di panel sebelah kiri.")

  except Exception as e:
    st.error(f"Terjadi kesalahan saat memproses file jadwal: {e}")
else:
  st.info(" ")
