import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches, Pt
import io

st.set_page_config(page_title="Dashboard CODEX", layout="wide")
st.title("📊 Dashboard Ticket CODEX (Versi Stabil – Tanpa Browser)")

# Upload
uploaded = st.file_uploader("📁 Upload File CODEX (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("Silakan upload file CODEX (.xlsx) terlebih dahulu.")
    st.stop()

# Load Data
df = pd.read_excel(uploaded)
df.columns = [c.strip() for c in df.columns]

df["CREATE TICKET"] = pd.to_datetime(df["CREATE TICKET"], errors="coerce")
df["AGE_DAYS"] = (datetime.now() - df["CREATE TICKET"]).dt.days
df["MONTH"] = df["CREATE TICKET"].dt.to_period("M").astype(str)

# Filter
st.sidebar.title("🔍 Filter")
sub_list = ["ALL"] + sorted(df["ASSIGN DIVISION"].dropna().unique())
status_list = ["ALL"] + sorted(df["STATUS"].dropna().unique())

sub_f = st.sidebar.selectbox("Sub Divisi", sub_list)
sta_f = st.sidebar.selectbox("Status Ticket", status_list)

fdf = df.copy()
if sub_f != "ALL":
    fdf = fdf[fdf["ASSIGN DIVISION"] == sub_f]
if sta_f != "ALL":
    fdf = fdf[fdf["STATUS"] == sta_f]

# Grafik
fig_age = px.histogram(fdf, x="AGE_DAYS", color="STATUS", nbins=40, title="Distribusi Umur Ticket")
st.plotly_chart(fig_age, use_container_width=True)

# Summary Status
count_status = fdf.groupby("STATUS").size().reset_index(name="JUMLAH")
st.subheader("📌 Ringkasan Ticket")
st.dataframe(count_status)

# PIC per bulan
pic_month = df.groupby(["ASSIGN DIVISION", "MONTH"]).size().reset_index(name="JUMLAH")

fig_pic = px.bar(pic_month, x="MONTH", y="JUMLAH", color="ASSIGN DIVISION",
                 title="Produktivitas PIC per Bulan")
st.plotly_chart(fig_pic, use_container_width=True)

# Generate PPT (TANPA GRAFIK)
def generate_ppt(df_sum, df_pic):

    prs = Presentation()
    
    # Cover
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Laporan Ticket CODEX"
    slide.placeholders[1].text = "Generated otomatis oleh Dashboard Streamlit"

    # Slide Ringkasan
    slide2 = prs.slides.add_slide(prs.slide_layouts[1])
    slide2.shapes.title.text = "Ringkasan Ticket"

    body = slide2.placeholders[1].text_frame
    body.text = ""
    for _, row in df_sum.iterrows():
        body.add_paragraph().text = f"{row['STATUS']} : {row['JUMLAH']} Ticket"

    # Slide PIC Bulanan
    slide3 = prs.slides.add_slide(prs.slide_layouts[1])
    slide3.shapes.title.text = "Produktivitas PIC per Bulan"
    body2 = slide3.placeholders[1].text_frame
    body2.text = ""

    for _, row in df_pic.iterrows():
        body2.add_paragraph().text = (
            f"{row['ASSIGN DIVISION']} – {row['MONTH']} : {row['JUMLAH']} Ticket"
        )

    # Slide Saran
    slide4 = prs.slides.add_slide(prs.slide_layouts[1])
    slide4.shapes.title.text = "Saran Peningkatan Kinerja"

    saran = """
1. PIC dengan jumlah ticket tinggi perlu evaluasi beban kerja dan SOP.
2. Ticket usia > 30 hari harus mendapatkan prioritas untuk percepatan root-cause.
3. Perlu reminder otomatis harian untuk ticket yang belum ditutup.
4. PIC wajib update status minimal setiap 24 jam.
    """

    t = slide4.placeholders[1].text_frame
    t.text = saran

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()

# Tombol Download PPT
st.subheader("📥 Download Presentasi PPT")

if st.button("Generate PPT"):
    ppt = generate_ppt(count_status, pic_month)

    st.success("PPT berhasil dibuat!")

    st.download_button(
        "📥 Download PPT",
        data=ppt,
        file_name="Laporan_Ticket_CODEX.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

# Tabel Data
st.subheader("📑 Data Lengkap")
st.dataframe(fdf)
