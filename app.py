import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches
import io

st.set_page_config(page_title="Dashboard Ticket CODEX (Full Ver)", layout="wide")

# ==========================
# HEADER
# ==========================
st.title("📊 Dashboard Analisa Ticket CODEX — Versi Lengkap")

# ==========================
# UPLOAD FILE
# ==========================
uploaded = st.file_uploader("📁 Upload File CODEX (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("Silakan upload file CODEX (.xlsx) terlebih dahulu.")
    st.stop()

# ==========================
# LOAD DATA
# ==========================
df = pd.read_excel(uploaded)

# Normalisasi kolom
df.columns = [c.strip() for c in df.columns]

# Pastikan tanggal dalam format datetime
df["CREATE TICKET"] = pd.to_datetime(df["CREATE TICKET"], errors="coerce")

# Hitung Umur Ticket
today = pd.Timestamp.today()

def hitung_age(row):
    if str(row["STATUS"]).lower() == "open":
        return (today - row["CREATE TICKET"]).days
    else:
        return (row["CREATE TICKET"].max() - row["CREATE TICKET"]).days

df["AGE_DAYS"] = (today - df["CREATE TICKET"]).dt.days

# ==========================
# FILTER
# ==========================
st.sidebar.header("🔍 Filter Data")

sub_list = ["ALL"] + sorted(df["ASSIGN DIVISION"].dropna().unique().tolist())
status_list = ["ALL"] + sorted(df["STATUS"].dropna().unique().tolist())

sub_filter = st.sidebar.selectbox("Sub Divisi", sub_list)
status_filter = st.sidebar.selectbox("Status Ticket", status_list)

dff = df.copy()
if sub_filter != "ALL":
    dff = dff[dff["ASSIGN DIVISION"] == sub_filter]
if status_filter != "ALL":
    dff = dff[dff["STATUS"] == status_filter]

# ==========================
# STATISTIK 1 — Jumlah Ticket per Status
# ==========================
st.subheader("📌 Statistik Ticket Berdasarkan STATUS")

count_status = dff["STATUS"].value_counts().reset_index()
count_status.columns = ["STATUS", "JUMLAH"]

fig_status = px.bar(
    count_status,
    x="STATUS",
    y="JUMLAH",
    color="STATUS",
    title="Jumlah Ticket Berdasarkan Status"
)

st.plotly_chart(fig_status, use_container_width=True)

# ==========================
# STATISTIK 2 — Umur Ticket (AGE)
# ==========================
st.subheader("⏱ Distribusi Umur Ticket (Days)")

fig_age = px.histogram(
    dff,
    x="AGE_DAYS",
    nbins=20,
    color="STATUS",
    title="Distribusi Umur Ticket"
)

st.plotly_chart(fig_age, use_container_width=True)

# ==========================
# STATISTIK 3 — Performa PIC / Bulan
# ==========================
st.subheader("👷 Performa PIC per Bulan")

df["MONTH"] = df["CREATE TICKET"].dt.to_period("M").astype(str)

pic_perf = df.groupby(["ASSIGN DIVISION", "MONTH"]).size().reset_index(name="JUMLAH")

fig_pic = px.bar(
    pic_perf,
    x="MONTH",
    y="JUMLAH",
    color="ASSIGN DIVISION",
    title="Performa PIC / Bulan"
)

st.plotly_chart(fig_pic, use_container_width=True)

# ==========================
# SARAN PERBAIKAN (AUTO)
# ==========================
st.subheader("📝 Rekomendasi Perbaikan Kerja PIC")

recommend_text = """
1. **PIC perlu mempercepat penyelesaian ticket yang sudah berumur > 30 hari**, karena berdampak pada stabilitas jaringan & eskalasi pelanggan.
2. Ticket dengan status **Warning/Alert** harus diberi prioritas harian.
3. Mapping ulang kapasitas interface yang sering muncul di ticket.
4. PIC disarankan membuat **checkpoint mingguan** untuk mencegah ticket aging menumpuk.
"""

st.warning(recommend_text)

# ==========================
# TAMPILKAN DATA LENGKAP
# ==========================
st.subheader("📄 Data Lengkap Ticket CODEX")
st.dataframe(dff, use_container_width=True)

# ==========================
# EXPORT PPT
# ==========================

def generate_ppt(df_summary, fig_age, fig_pic):
    prs = Presentation("template PPT Moratel.pptx")

    # SLIDE 1 — Summary Status
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = "Ringkasan Ticket CODEX"

    txt = slide.shapes.placeholders[1].text_frame
    for idx, row in df_summary.iterrows():
        txt.text += f"{row['STATUS']}: {row['JUMLAH']} Ticket\n"

    # SLIDE 2 — Age Distribution
    img_bytes = fig_age.to_image(format="png")
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Distribusi Umur Ticket"
    slide.shapes.add_picture(io.BytesIO(img_bytes), Inches(1), Inches(1), width=Inches(8))

    # SLIDE 3 — PIC Performance
    img2 = fig_pic.to_image(format="png")
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Performa PIC / Bulan"
    slide.shapes.add_picture(io.BytesIO(img2), Inches(1), Inches(1), width=Inches(8))

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()

st.subheader("📥 Download Presentasi PPT")

if st.button("Generate PPT"):
    ppt_bytes = generate_ppt(count_status, fig_age, fig_pic)
    st.download_button(
        label="⬇ Download PPT",
        data=ppt_bytes,
        file_name="Presentasi-CODEX.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
