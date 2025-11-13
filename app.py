import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# ============================
# CONFIG PAGE
# ============================
st.set_page_config(page_title="Dashboard Ticket CODEX", layout="wide")
st.title("📊 Dashboard Analisa Ticket CODEX (Versi Lengkap)")

# ============================
# UPLOAD FILE
# ============================
uploaded = st.file_uploader("📁 Upload File CODEX (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("Silakan upload file CODEX (.xlsx) terlebih dahulu.")
    st.stop()

# ============================
# LOAD DATA
# ============================
df = pd.read_excel(uploaded)

# Normalisasi nama kolom
df.columns = [c.strip() for c in df.columns]

# Pastikan kolom wajib ada
required_cols = ["CREATE TICKET", "STATUS", "NO-TICKET"]
missing = [c for c in required_cols if c not in df.columns]
if len(missing) > 0:
    st.error(f"Kolom wajib tidak ditemukan: {missing}")
    st.stop()

# ============================
# FUNGSI HITUNG AGING
# ============================
def hitung_aging(row):
    status = str(row.get("STATUS", "")).lower().strip()

    create_date = row.get("CREATE TICKET", None)
    if pd.isna(create_date):
        return 0

    try:
        create_date = pd.to_datetime(create_date)
    except:
        return 0

    age = (datetime.now() - create_date).days
    return max(age, 0)

# Tambahkan kolom AGING
df["AGE_DAYS"] = df.apply(hitung_aging, axis=1)

# ============================
# FILTER
# ============================
st.sidebar.header("⚙️ Filter")

# Filter Sub Divisi jika ada
subdiv_col = None
for c in df.columns:
    if "DIV" in c.upper():
        subdiv_col = c
        break

if subdiv_col:
    subdiv_list = ["ALL"] + sorted(df[subdiv_col].dropna().unique().tolist())
    subdiv = st.sidebar.selectbox("Sub Divisi", subdiv_list)

    if subdiv != "ALL":
        df = df[df[subdiv_col] == subdiv]

# Filter Status
status_list = ["ALL"] + sorted(df["STATUS"].dropna().unique().tolist())
status = st.sidebar.selectbox("Status Ticket", status_list)

if status != "ALL":
    df = df[df["STATUS"] == status]

# ============================
# STATISTIK RINGKAS
# ============================
st.subheader("📌 Ringkasan Data")

col1, col2, col3 = st.columns(3)
col1.metric("Total Ticket", len(df))
col2.metric("Ticket OPEN", len(df[df["STATUS"].str.lower() == "open"]))
col3.metric("Ticket CLOSED", len(df[df["STATUS"].str.lower() == "close"]))

# ============================
# DISTRIBUSI AGING
# ============================
st.subheader("⏳ Distribusi Umur Ticket (Aging)")

fig_age = px.histogram(
    df,
    x="AGE_DAYS",
    nbins=20,
    title="Distribusi Umur Ticket (Hari)",
    color="STATUS"
)
st.plotly_chart(fig_age, use_container_width=True)

# ============================
# TABEL DETAIL
# ============================
st.subheader("📋 Data Lengkap Ticket CODEX")
st.dataframe(df, use_container_width=True)

# Download button
@st.cache_data
def konversi_excel(df):
    return df.to_excel("output.xlsx", index=False)

st.download_button(
    "📥 Download Data Hasil Olahan (.xlsx)",
    data=df.to_csv(index=False).encode("utf-8"),
    file_name="codex_processed.csv",
    mime="text/csv"
)
