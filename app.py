import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="Dashboard Ticket CODEX", layout="wide")

st.title("📊 Dashboard Analisa Ticket CODEX (Versi Lengkap)")

uploaded = st.file_uploader("📂 Upload File CODEx (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("Silakan upload file CODEx (.xlsx) terlebih dahulu.")
    st.stop()

# ============================
# LOAD DATA
# ============================
df = pd.read_excel(uploaded)

# Normalisasi nama kolom
df.columns = [c.strip() for c in df.columns]

# pastikan kolom tanggal ada
if "CREATE TICKET" not in df.columns:
    st.error("Kolom 'CREATE TICKET' tidak ditemukan!")
    st.stop()

# Convert tanggal
df["CREATE TICKET"] = pd.to_datetime(df["CREATE TICKET"], errors="coerce")

if "CLOSE TICKET" in df.columns:
    df["CLOSE TICKET"] = pd.to_datetime(df["CLOSE TICKET"], errors="coerce")
else:
    df["CLOSE TICKET"] = None

today = datetime.now()

# ============================
# PERHITUNGAN AGING
# ============================
def hitung_aging(row):
    if row["STATUS"].lower() == "open":
        return (today - row["CREATE TICKET"]).days
    elif pd.notnull(row["CLOSE TICKET"]):
        return (row["CLOSE TICKET"] - row["CREATE TICKET"]).days
    else:
        return 0

df["AGE_DAYS"] = df.apply(hitung_aging, axis=1)

# ============================
# SLA LEVEL
# ============================
def sla_level(age):
    if age > 30:
        return "Critical (>30 hari)"
    elif age > 7:
        return "Major (7–30 hari)"
    elif age > 3:
        return "Minor (3–7 hari)"
    else:
        return "Normal (<3 hari)"

df["SLA_LEVEL"] = df["AGE_DAYS"].apply(sla_level)

# ============================
# SIDEBAR FILTER
# ============================
st.sidebar.header("⚙️ Filter")

subdivisi_list = ["ALL"] + sorted(df["ASSIGN DIVISION"].dropna().unique().tolist())
status_list = ["ALL"] + sorted(df["STATUS"].dropna().unique().tolist())
sla_list = ["ALL"] + sorted(df["SLA_LEVEL"].unique())

f_subdiv = st.sidebar.selectbox("Sub Divisi", subdivisi_list)
f_status = st.sidebar.selectbox("Status Ticket", status_list)
f_sla = st.sidebar.selectbox("SLA Level", sla_list)

filtered_df = df.copy()

if f_subdiv != "ALL":
    filtered_df = filtered_df[filtered_df["ASSIGN DIVISION"] == f_subdiv]

if f_status != "ALL":
    filtered_df = filtered_df[filtered_df["STATUS"] == f_status]

if f_sla != "ALL":
    filtered_df = filtered_df[filtered_df["SLA_LEVEL"] == f_sla]

# ============================
# TAMPILKAN DATA
# ============================
st.subheader("📁 Data Ticket (Setelah Filter)")
st.dataframe(filtered_df, use_container_width=True)

# ============================
# GRAFIK STATUS OPEN / CLOSE
# ============================
st.subheader("📊 Distribusi Status Ticket")
fig_pie = px.pie(df, names="STATUS", title="Persentase Ticket Open / Close")
st.plotly_chart(fig_pie, use_container_width=True)

# ============================
# GRAFIK SLA LEVEL
# ============================
st.subheader("🚦 SLA Level Ticket")
fig_sla = px.bar(
    df["SLA_LEVEL"].value_counts().reset_index(),
    x="index",
    y="SLA_LEVEL",
    labels={"index": "SLA Level", "SLA_LEVEL": "Jumlah"},
    color="index",
)
st.plotly_chart(fig_sla, use_container_width=True)

# ============================
# GRAFIK AGING
# ============================
st.subheader("📈 Aging Ticket (Open Only)")
df_open = df[df["STATUS"] == "Open"]

if len(df_open) > 0:
    fig_age = px.bar(
        df_open,
        x="NO-TICKET",
        y="AGE_DAYS",
        color="AGE_DAYS",
        text="AGE_DAYS",
        title="Umur Ticket (Hari) – Ticket OPEN",
    )
    st.plotly_chart(fig_age, use_container_width=True)
else:
    st.info("Tidak ada ticket open.")

# ============================
# TOP 10 TICKET TERTUA
# ============================
st.subheader("🔥 TOP 10 Ticket Paling Lama")
top10 = df_open.sort_values("AGE_DAYS", ascending=False).head(10)
st.dataframe(top10, use_container_width=True)

# ============================
# RATA-RATA TICKET PER DIVISI
# ============================
if "ASSIGN DIVISION" in df.columns:
    st.subheader("🏢 Rata-rata Aging per Assign Division")
    avg_age = df_open.groupby("ASSIGN DIVISION")["AGE_DAYS"].mean().reset_index()

    fig_avg = px.bar(
        avg_age,
        x="ASSIGN DIVISION",
        y="AGE_DAYS",
        title="Rata-rata Umur Ticket per Divisi",
        color="AGE_DAYS",
    )
    st.plotly_chart(fig_avg, use_container_width=True)
