import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Dashboard CODEX Moratelindo", layout="wide")

st.title("📊 Dashboard Analisa Ticket CODEX")
st.caption("Upload file CODEX (CSV/XLSX) → otomatis dianalisa & divisualisasikan")

uploaded = st.file_uploader("📁 Upload File CODEX (CSV / Excel)", type=["csv", "xlsx"])

# ==========================
# FUNGSI BANTUAN
# ==========================

def detect_col(cols, keywords):
    """Cari kolom berdasarkan keyword (tidak case sensitive, ignore spasi)."""
    for c in cols:
        name = c.lower().replace(" ", "")
        for key in keywords:
            if key in name:
                return c
    return None

def to_excel_bytes(df):
    """Convert DataFrame ke bytes Excel untuk download."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="HASIL")
    output.seek(0)
    return output

# ==========================
# MAIN
# ==========================

if uploaded:
    # ===== BACA FILE =====
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)

    st.subheader("🔍 Data Mentah")
    st.dataframe(df, use_container_width=True)

    cols = df.columns

    # ===== DETEKSI KOLOM-KOLOM PENTING =====
    col_subdiv   = detect_col(cols, ["subdiv", "sub_div", "division", "divisi"])
    col_opendate = detect_col(cols, ["opendate", "tglopen", "tanggalopen", "tanggal", "opentime", "created"])
    col_status   = detect_col(cols, ["status"])
    col_aging    = detect_col(cols, ["aging", "umur", "hari"])

    if not col_subdiv:
        st.error("❌ Kolom Sub Divisi / Divisi tidak ditemukan. Tolong cek nama kolom di file.")
        st.stop()

    # Normalisasi Sub Divisi
    df[col_subdiv] = df[col_subdiv].astype(str).str.upper().str.strip()

    # ===== HITUNG AGING =====
    if col_aging is not None:
        df["AGING"] = pd.to_numeric(df[col_aging], errors="coerce")
    elif col_opendate is not None:
        df[col_opendate] = pd.to_datetime(df[col_opendate], errors="coerce")
        today = pd.to_datetime("today").normalize()
        df["AGING"] = (today - df[col_opendate]).dt.days
    else:
        df["AGING"] = np.nan

    # ===== SIDEBAR FILTER =====
    st.sidebar.header("⚙️ Filter")

    # Filter Sub Divisi
    subdivisi_list = sorted(df[col_subdiv].dropna().unique())
    pilih_subdiv = st.sidebar.multiselect("Sub Divisi", ["ALL"] + list(subdivisi_list), default="ALL")

    df_filtered = df.copy()
    if "ALL" not in pilih_subdiv:
        df_filtered = df_filtered[df_filtered[col_subdiv].isin(pilih_subdiv)]

    # Filter Status (kalau ada)
    if col_status:
        status_list = sorted(df_filtered[col_status].dropna().astype(str).unique())
        pilih_status = st.sidebar.multiselect("Status", ["ALL"] + list(status_list), default="ALL")
        if "ALL" not in pilih_status:
            df_filtered = df_filtered[df_filtered[col_status].astype(str).isin(pilih_status)]

    # Filter tanggal (kalau ada kolom tanggal)
    if col_opendate:
        df_filtered[col_opendate] = pd.to_datetime(df_filtered[col_opendate], errors="coerce")
        min_date = df_filtered[col_opendate].min()
        max_date = df_filtered[col_opendate].max()
        if pd.notna(min_date) and pd.notna(max_date):
            start_date, end_date = st.sidebar.date_input(
                "Rentang Tanggal Open",
                value=(min_date.date(), max_date.date())
            )
            if isinstance(start_date, tuple):
                start_date, end_date = start_date  # safety
            mask_date = (df_filtered[col_opendate] >= pd.to_datetime(start_date)) & \
                        (df_filtered[col_opendate] <= pd.to_datetime(end_date))
            df_filtered = df_filtered[mask_date]

    # ===== SUMMARY KPI =====
    total = len(df_filtered)
    lt7   = len(df_filtered[df_filtered["AGING"] < 7])
    lt30  = len(df_filtered[(df_filtered["AGING"] >= 7) & (df_filtered["AGING"] < 30)])
    gt30  = len(df_filtered[df_filtered["AGING"] >= 30])

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Ticket (Filter Aktif)", total)
    col2.metric("Ticket < 7 Hari", lt7)
    col3.metric("Ticket 7–29 Hari", lt30)
    col4.metric("Ticket ≥ 30 Hari", gt30)

    st.markdown("---")

    # ===== LAYOUT GRAFIK =====
    g1, g2 = st.columns(2)

    # Grafik 1: Jumlah ticket per Sub Divisi
    with g1:
        st.subheader("📦 Jumlah Ticket per Sub Divisi")
        chart1 = df_filtered.groupby(col_subdiv)["AGING"].count().reset_index()
        chart1 = chart1.rename(columns={"AGING": "Jumlah Ticket"})
        if len(chart1) > 0:
            fig1 = px.bar(chart1, x=col_subdiv, y="Jumlah Ticket", text="Jumlah Ticket")
            fig1.update_layout(xaxis_title="Sub Divisi", yaxis_title="Jumlah Ticket")
            st.plotly_chart(fig1, use_container_width=True)
        else:
            st.info("Tidak ada data setelah filter.")

    # Grafik 2: Rata-rata aging per Sub Divisi
    with g2:
        st.subheader("⏱️ Rata-rata Aging per Sub Divisi")
        chart2 = df_filtered.groupby(col_subdiv)["AGING"].mean().reset_index()
        chart2["AGING"] = chart2["AGING"].round(1)
        chart2 = chart2.rename(columns={"AGING": "Rata-rata Aging (hari)"})
        if len(chart2) > 0:
            fig2 = px.bar(chart2, x=col_subdiv, y="Rata-rata Aging (hari)", text="Rata-rata Aging (hari)")
            fig2.update_layout(xaxis_title="Sub Divisi", yaxis_title="Rata-rata Aging (hari)")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("Tidak ada data untuk dihitung rata-rata aging-nya.")

    st.markdown("---")

    # Grafik 3: Tren ticket per hari (kalau ada tanggal)
    if col_opendate:
        st.subheader("📆 Tren Ticket per Hari")
        df_trend = df_filtered.dropna(subset=[col_opendate]).copy()
        if len(df_trend) > 0:
            df_trend["OpenDateOnly"] = df_trend[col_opendate].dt.date
            trend = df_trend.groupby("OpenDateOnly")["AGING"].count().reset_index()
            trend = trend.rename(columns={"AGING": "Jumlah Ticket"})
            fig3 = px.line(trend, x="OpenDateOnly", y="Jumlah Ticket", markers=True)
            fig3.update_layout(xaxis_title="Tanggal Open", yaxis_title="Jumlah Ticket")
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("Tidak ada data tanggal untuk dibuat tren.")

    # Grafik 4: Pie distribusi aging
    st.subheader("🥧 Distribusi Aging")
    df_pie = df_filtered.copy()
    df_pie["Aging Group"] = pd.cut(
        df_pie["AGING"],
        bins=[-1, 6, 29, 99999],
        labels=["<7 Hari", "7–29 Hari", "≥30 Hari"]
    )
    if df_pie["Aging Group"].notna().any():
        fig4 = px.pie(df_pie, names="Aging Group")
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("Distribusi aging belum bisa dihitung (data aging kosong).")

    st.markdown("---")

    # ===== TABEL AKHIR =====
    st.subheader("📄 Tabel Hasil (Setelah Filter)")
    st.dataframe(df_filtered, use_container_width=True)

    # ===== REKOMENDASI OTOMATIS SEDERHANA =====
    st.subheader("💡 Rekomendasi Otomatis (Simple Insight)")
    if len(df_filtered) > 0:
        avg_aging_by_subdiv = df_filtered.groupby(col_subdiv)["AGING"].mean().sort_values(ascending=False)
        worst_subdiv = avg_aging_by_subdiv.index[0]
        worst_aging = round(avg_aging_by_subdiv.iloc[0], 1)

        text = f"""
- Sub Divisi dengan **rata-rata aging tertinggi**: **{worst_subdiv}** (~{worst_aging} hari)  
- Prioritas percepatan follow-up sebaiknya difokuskan ke Sub Divisi tersebut.  
- Pertimbangkan:
  - review ulang SOP eskalasi di {worst_subdiv}  
  - cek apakah beban kerja engineer tidak seimbang  
  - tambahkan reminder / notifikasi untuk ticket > 30 hari
"""
        st.markdown(text)
    else:
        st.info("Tidak ada data untuk dibuat rekomendasi.")

    # ===== DOWNLOAD HASIL =====
    st.subheader("⬇️ Download Hasil")
    excel_bytes = to_excel_bytes(df_filtered)
    st.download_button(
        label="Download Hasil (Excel)",
        data=excel_bytes,
        file_name="hasil_dashboard_codex.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Silakan upload file CODEX (CSV / Excel) terlebih dahulu.")
