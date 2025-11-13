import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches
import io

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(page_title="Dashboard Ticket CODEX", layout="wide")
st.title("📊 Dashboard Analisa Ticket CODEX (Versi Lengkap)")

# =========================
# FILE UPLOADER (1 FILE SAJA)
# =========================
uploaded = st.file_uploader("📁 Upload File CODEX (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("Silakan upload file CODEX (.xlsx) terlebih dahulu.")
    st.stop()

df = pd.read_excel(uploaded)

# =========================
# NORMALISASI KOLOM
# =========================
df.columns = [c.strip().upper() for c in df.columns]

# Pastikan ada kolom tanggal
tanggal_col = "CREATE TICKET"
if tanggal_col not in df.columns:
    st.error(f"Kolom '{tanggal_col}' tidak ditemukan. Pastikan nama kolom sama.")
    st.stop()

# Convert waktu
df[tanggal_col] = pd.to_datetime(df[tanggal_col])

# Hitung umur tiket
def hitung_umur(row):
    if isinstance(row[tanggal_col], pd.Timestamp):
        return (datetime.now() - row[tanggal_col]).days
    return 0

df["AGE_DAYS"] = df.apply(hitung_umur, axis=1)

# Normalisasi STATUS
df["STATUS_TICKET"] = df["STATUS"].str.strip().str.lower()

# =========================
# FILTER SAMPING
# =========================
st.sidebar.header("⚙️ Filter")

subdivisi_list = ["ALL"] + sorted(df["ASSIGN DIVISION"].dropna().unique().tolist())
subdivisi = st.sidebar.selectbox("Sub Divisi", subdivisi_list)

status_list = ["ALL", "open", "close"]
status_ticket = st.sidebar.selectbox("Status Ticket", status_list)

df_filtered = df.copy()

# Filter sub divisi
if subdivisi != "ALL":
    df_filtered = df_filtered[df_filtered["ASSIGN DIVISION"] == subdivisi]

# Filter status
if status_ticket != "ALL":
    df_filtered = df_filtered[df_filtered["STATUS_TICKET"] == status_ticket]

# =========================
# ANALISA 1 — Summary Kategori (Count)
# =========================
count_status = df_filtered.groupby("STATUS").size().reset_index(name="JUMLAH")

st.subheader("📌 Grafik Jumlah Ticket Berdasarkan Kategori")
fig1 = px.bar(
    count_status,
    x="STATUS",
    y="JUMLAH",
    color="STATUS",
    text="JUMLAH"
)
st.plotly_chart(fig1, use_container_width=True)

# =========================
# ANALISA 2 — AGE HISTOGRAM
# =========================
st.subheader("⏱️ Distribusi Umur Ticket (Hari)")
fig_age = px.histogram(df_filtered, x="AGE_DAYS", nbins=20, color="STATUS")
st.plotly_chart(fig_age, use_container_width=True)

# =========================
# ANALISA 3 — Produktivitas PIC per Bulan
# =========================
st.subheader("👨‍💻 Produktivitas PIC (Ticket yang diselesaikan)")

df_closed = df[df["STATUS_TICKET"] == "close"].copy()
df_closed["BULAN"] = df_closed["CREATE TICKET"].dt.to_period("M").astype(str)

pic_summary = df_closed.groupby(["ASSIGN DIVISION", "BULAN"]).size().reset_index(name="TIKET_SELESAI")

fig_pic = px.bar(
    pic_summary,
    x="BULAN",
    y="TIKET_SELESAI",
    color="ASSIGN DIVISION",
    barmode="group",
    text="TIKET_SELESAI"
)

st.plotly_chart(fig_pic, use_container_width=True)

st.dataframe(pic_summary, use_container_width=True)

# =========================
# PPT GENERATOR FUNCTION
# =========================
def generate_ppt(df_summary, fig_age, fig_pic):
    prs = Presentation("template PPT Moratel.pptx")

    # SLIDE 1 — SUMMARY TABLE
    slide1 = prs.slides.add_slide(prs.slide_layouts[1])
    slide1.shapes.title.text = "Summary Ticket CODEX"

    rows, cols = df_summary.shape
    table = slide1.shapes.add_table(
        rows + 1, cols,
        Inches(0.5), Inches(1.5), Inches(9), Inches(0.8)
    ).table

    # header
    for j, col in enumerate(df_summary.columns):
        table.cell(0, j).text = col

    # isi
    for i in range(rows):
        for j in range(cols):
            table.cell(i+1, j).text = str(df_summary.iloc[i, j])

    # SLIDE 2 — AGE
    img_bytes = io.BytesIO()
    fig_age.write_image(img_bytes, format="png")
    img_bytes.seek(0)
    slide2 = prs.slides.add_slide(prs.slide_layouts[5])
    slide2.shapes.add_picture(img_bytes, Inches(1), Inches(1), width=Inches(8))

    # SLIDE 3 — PIC GRAPH
    img2_bytes = io.BytesIO()
    fig_pic.write_image(img2_bytes, format="png")
    img2_bytes.seek(0)
    slide3 = prs.slides.add_slide(prs.slide_layouts[5])
    slide3.shapes.add_picture(img2_bytes, Inches(1), Inches(1), width=Inches(8))

    # Save PPT
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# =========================
# BUTTON DOWNLOAD PPT
# =========================
st.subheader("📥 Download Presentasi PPT")

if st.button("Generate PPT"):
    ppt_data = generate_ppt(
        df_summary=count_status,
        fig_age=fig_age,
        fig_pic=fig_pic
    )

    st.success("PPT berhasil dibuat! Silakan download.")
    st.download_button(
        label="Download PPT",
        data=ppt_data,
        file_name="Analisa_Ticket_CODEX.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
