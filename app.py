import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches
from html2image import Html2Image

hti = Html2Image()

# ==============================
# PAGE CONFIG
# ==============================
st.set_page_config(page_title="Dashboard Ticket CODEX", layout="wide")

st.title("📊 Dashboard Analisa Ticket CODEX (Versi Lengkap)")

uploaded = st.file_uploader("📁 Upload File CODEX (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("Silakan upload file CODEX (.xlsx) terlebih dahulu.")
    st.stop()

# ==============================
# LOAD DATA
# ==============================
df = pd.read_excel(uploaded)

# Normalisasi kolom
df.columns = [c.strip() for c in df.columns]

# Pastikan kolom tanggal ada
if "CREATE TICKET" not in df.columns:
    st.error("Kolom 'CREATE TICKET' tidak ditemukan.")
    st.stop()

# Convert tanggal
df["CREATE TICKET"] = pd.to_datetime(df["CREATE TICKET"], errors="coerce")

# Hitung umur ticket (hari)
today = datetime.now()
df["AGE_DAYS"] = (today - df["CREATE TICKET"]).dt.days

# ==============================
# FILTER
# ==============================
st.sidebar.title("🔍 Filter Data")

sub_divisi_list = ["ALL"] + sorted(df["ASSIGN DIVISION"].dropna().unique())
status_list = ["ALL"] + sorted(df["STATUS"].dropna().unique())

subdiv_filter = st.sidebar.selectbox("Sub Divisi", sub_divisi_list)
status_filter = st.sidebar.selectbox("Status Ticket", status_list)

filtered_df = df.copy()
if subdiv_filter != "ALL":
    filtered_df = filtered_df[filtered_df["ASSIGN DIVISION"] == subdiv_filter]
if status_filter != "ALL":
    filtered_df = filtered_df[filtered_df["STATUS"] == status_filter]

# ==============================
# GRAFIK UMUR TICKET
# ==============================
fig_age = px.histogram(
    filtered_df,
    x="AGE_DAYS",
    color="STATUS",
    nbins=30,
    title="Distribusi Umur Ticket (AGE_DAYS)",
    color_discrete_sequence=px.colors.qualitative.Set2
)

st.plotly_chart(fig_age, use_container_width=True)

# ==============================
# HITUNG JUMLAH STATUS
# ==============================
count_status = filtered_df.groupby("STATUS").size().reset_index(name="JUMLAH")

st.subheader("📌 Ringkasan Ticket berdasarkan Status")
st.dataframe(count_status)

# ==============================
# HITUNG KINERJA PIC
# ==============================
df["MONTH"] = df["CREATE TICKET"].dt.to_period("M").astype(str)
pic_month = (
    df.groupby(["ASSIGN DIVISION", "MONTH"])
    .size()
    .reset_index(name="JUMLAH")
)

fig_pic = px.bar(
    pic_month,
    x="MONTH",
    y="JUMLAH",
    color="ASSIGN DIVISION",
    title="Produktivitas PIC per Bulan",
)

st.plotly_chart(fig_pic, use_container_width=True)

# ==============================
# DOWNLOAD PPT FUNCTIONS
# ==============================

def fig_to_png(fig):
    """Convert Plotly figure to PNG using html2image."""
    html = fig.to_html(include_plotlyjs="cdn")
    hti.screenshot(html_str=html, save_as="temp.png", size=(900, 500))

    with open("temp.png", "rb") as f:
        return f.read()


def generate_ppt(df_summary, fig_age, fig_pic):
    prs = Presentation("template PPT Moratel.pptx")

    # Slide 1 – Ringkasan Status
    slide1 = prs.slides.add_slide(prs.slide_layouts[1])
    slide1.shapes.title.text = "Ringkasan Ticket CODEX"
    body = slide1.shapes.placeholders[1].text_frame

    for idx, row in df_summary.iterrows():
        body.text += f"{row['STATUS']}: {row['JUMLAH']} Ticket\n"

    # Slide 2 – Grafik AGE_DAYS
    age_img = fig_to_png(fig_age)
    slide2 = prs.slides.add_slide(prs.slide_layouts[5])
    slide2.shapes.title.text = "Distribusi Umur Ticket"
    slide2.shapes.add_picture(io.BytesIO(age_img), Inches(1), Inches(1), width=Inches(8))

    # Slide 3 – Grafik PIC
    pic_img = fig_to_png(fig_pic)
    slide3 = prs.slides.add_slide(prs.slide_layouts[5])
    slide3.shapes.title.text = "Performa PIC per Bulan"
    slide3.shapes.add_picture(io.BytesIO(pic_img), Inches(1), Inches(1), width=Inches(8))

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()

# ==============================
# DOWNLOAD PPT BUTTON
# ==============================
st.subheader("📥 Download Presentasi PPT")

if st.button("Generate PPT"):
    ppt_file = generate_ppt(count_status, fig_age, fig_pic)
    st.success("PPT berhasil dibuat!")

    st.download_button(
        label="📥 Download PPT",
        data=ppt_file,
        file_name="Laporan_Ticket_CODEX.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

# ==============================
# DATA TABEL LENGKAP
# ==============================
st.subheader("📑 Data Lengkap Ticket CODEX")
st.dataframe(filtered_df)
