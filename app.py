import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="Dashboard Ticket CODEX", layout="wide")

# =========================================================
# TITLE
# =========================================================
st.title("📊 Dashboard Analisa Ticket CODEX — Versi Lengkap")

uploaded = st.file_uploader("📁 Upload File CODEX (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.info("Silakan upload file CODEx (.xlsx) terlebih dahulu.")
    st.stop()

# Load Data
df = pd.read_excel(uploaded)
df.columns = [c.strip() for c in df.columns]

# Normalisasi kolom wajib
required_cols = ["NO-TICKET", "HOSTNAME", "INTERFACE", "STATUS", 
                 "ASSIGN DIVISION", "CREATE TICKET"]

for col in required_cols:
    if col not in df.columns:
        st.error(f"Kolom **{col}** tidak ditemukan. Periksa file Excel.")
        st.stop()

# Convert tanggal → datetime
df["CREATE TICKET"] = pd.to_datetime(df["CREATE TICKET"], errors="coerce")

# Hitung age of ticket
df["AGE_DAYS"] = (datetime.now() - df["CREATE TICKET"]).dt.days

# Sub Divisi extracted
df["SUB_DIVISI"] = df["ASSIGN DIVISION"].str.split("-").str[-1].str.strip()

# PIC extracted
df["PIC"] = df["ASSIGN DIVISION"].str.split("-").str[0].str.strip()

# Status ticket OPEN / CLOSE (jika ada 2 kolom STATUS)
if "STATUS.1" in df.columns:
    df["STATUS_TICKET"] = df["STATUS.1"]
else:
    df["STATUS_TICKET"] = df["STATUS"]

# ===================================================================
# SIDEBAR FILTER
# ===================================================================
st.sidebar.header("⚙️ Filter")

subdivisi_opt = ["ALL"] + sorted(df["SUB_DIVISI"].dropna().unique().tolist())
status_opt = ["ALL"] + sorted(df["STATUS_TICKET"].dropna().unique().tolist())
pic_opt = ["ALL"] + sorted(df["PIC"].dropna().unique().tolist())

bulan_opt = ["ALL"] + sorted(df["CREATE TICKET"].dt.month_name().unique().tolist())
tahun_opt = ["ALL"] + sorted(df["CREATE TICKET"].dt.year.unique().tolist())

f_sub = st.sidebar.selectbox("Sub Divisi", subdivisi_opt)
f_status = st.sidebar.selectbox("Status Ticket", status_opt)
f_pic = st.sidebar.selectbox("PIC", pic_opt)
f_bulan = st.sidebar.selectbox("Bulan", bulan_opt)
f_tahun = st.sidebar.selectbox("Tahun", tahun_opt)

# Apply filter
filtered = df.copy()

if f_sub != "ALL":
    filtered = filtered[filtered["SUB_DIVISI"] == f_sub]

if f_status != "ALL":
    filtered = filtered[filtered["STATUS_TICKET"] == f_status]

if f_pic != "ALL":
    filtered = filtered[filtered["PIC"] == f_pic]

if f_bulan != "ALL":
    filtered = filtered[filtered["CREATE TICKET"].dt.month_name() == f_bulan]

if f_tahun != "ALL":
    filtered = filtered[filtered["CREATE TICKET"].dt.year == f_tahun]

st.success(f"Total data setelah filter: {len(filtered)}")

# ============================================================
# SUMMARY NUMBER CARD
# ============================================================
st.subheader("📌 Ringkasan Data")
c1, c2, c3, c4 = st.columns(4)

c1.metric("Total Ticket", len(filtered))
c2.metric("Critical", filtered["STATUS"].str.contains("Critical", case=False).sum())
c3.metric("Warning", filtered["STATUS"].str.contains("Warning", case=False).sum())
c4.metric("Average Age (Days)", round(filtered["AGE_DAYS"].mean(), 1))

# ============================================================
# GRAFIK STATUS
# ============================================================
st.subheader("📊 Distribusi Ticket Berdasarkan Kategori STATUS")
count_status = filtered["STATUS"].value_counts().reset_index()
count_status.columns = ["STATUS", "JUMLAH"]

fig_status = px.bar(count_status, x="STATUS", y="JUMLAH", color="STATUS",
                    color_discrete_sequence=px.colors.qualitative.Set1)
st.plotly_chart(fig_status, use_container_width=True)

# ============================================================
# TABLE RINCIAN STATUS
# ============================================================
st.subheader("📄 Tabel Rincian Status Ticket")

summary = filtered.groupby("STATUS").agg(
    JUMLAH=("STATUS", "count"),
    AVG_AGE=("AGE_DAYS", "mean"),
    MAX_AGE=("AGE_DAYS", "max")
).reset_index()

summary["AVG_AGE"] = summary["AVG_AGE"].round(1)
st.dataframe(summary, use_container_width=True)

# ============================================================
# HISTOGRAM AGE DAYS
# ============================================================
st.subheader("⏳ Distribusi Umur Ticket (AGE_DAYS)")

fig_age = px.histogram(filtered, x="AGE_DAYS", nbins=30, color="STATUS",
                       color_discrete_sequence=px.colors.qualitative.Set2)
st.plotly_chart(fig_age, use_container_width=True)

# ============================================================
# PRODUKTIVITAS PIC PER BULAN
# ============================================================
st.subheader("🧑‍💻 Performance PIC per Bulan")

filtered["MONTH"] = filtered["CREATE TICKET"].dt.month_name()
pic_perf = filtered.groupby(["PIC", "MONTH"]).size().reset_index(name="JUMLAH")

fig_pic = px.bar(pic_perf, x="PIC", y="JUMLAH", color="MONTH", barmode="group")
st.plotly_chart(fig_pic, use_container_width=True)

# ============================================================
# TABEL KINERJA PIC LENGKAP
# ============================================================
st.subheader("📋 Tabel Performance PIC")

pic_table = filtered.groupby("PIC").agg(
    TOTAL=("PIC", "count"),
    OPEN=("STATUS_TICKET", lambda x: (x == "Open").sum()),
    CLOSED=("STATUS_TICKET", lambda x: (x == "Close").sum()),
    AVG_AGE=("AGE_DAYS", "mean"),
    MAX_AGE=("AGE_DAYS", "max")
).reset_index()

pic_table["SLA_%"] = round((pic_table["CLOSED"] / pic_table["TOTAL"]) * 100, 1)
pic_table["AVG_AGE"] = pic_table["AVG_AGE"].round(1)

st.dataframe(pic_table, use_container_width=True)

# ============================================================
# REKOMENDASI OTOMATIS
# ============================================================
st.subheader("💡 Rekomendasi Perbaikan Kinerja PIC")

for _, row in pic_table.iterrows():
    pic = row["PIC"]
    avg_age = row["AVG_AGE"]
    sla = row["SLA_%"]

    st.markdown(f"### 🔧 {pic}")

    rekom = []

    if avg_age > 60:
        rekom.append("- Ticket lama > **60 hari**. Perlu daily follow-up & koordinasi lintas divisi.")

    if sla < 70:
        rekom.append("- SLA penyelesaian < **70%**. Perlu penambahan ritme closing ticket.")

    if row["OPEN"] > row["CLOSED"]:
        rekom.append("- Jumlah ticket OPEN lebih banyak dari CLOSED → potensi backlog.")

    if len(rekom) == 0:
        rekom.append("✔ Performance sangat baik & stabil.")

    for r in rekom:
        st.write(r)

    st.write("**Alasan bisnis:** Penyelesaian ticket cepat mengurangi risiko alarm berulang, mengurangi downtime, dan mempercepat troubleshooting NOC/BB/Access.\n---")


# ============================================================
# FULL DATA TABLE
# ============================================================
st.subheader("📑 Data Lengkap Ticket CODEX")
st.dataframe(filtered, use_container_width=True)
