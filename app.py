import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches
import io

st.set_page_config(page_title="Dashboard Ticket CODEX", layout="wide")
st.title("📊 Dashboard Analisa Ticket CODEX (Versi Lengkap)")


# ==========================
# UPLOAD FILE
# ==========================
uploaded = st.file_uploader("📁 Upload File CODEX (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("Silakan upload file CODEX (.xlsx) terlebih dahulu.")
    st.stop()

# Baca data
df = pd.read_excel(uploaded)

# Normalisasi nama kolom
df.columns = [c.strip() for c in df.columns]

# Pastikan kolom tanggal valid
if "CREATE TICKET" not in df.columns:
    st.error("Kolom 'CREATE TICKET' tidak ditemukan di excel.")
    st.stop()

df["CREATE TICKET"] = pd.to_datetime(df["CREATE TICKET"], errors="coerce")

# AGE DAYS
today = datetime.now()
df["AGE_DAYS"] = (today - df["CREATE TICKET"]).dt.days


# ==========================
# FILTER
# ==========================
col1, col2 = st.sidebar.columns(1)

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
# GRAFIK AGE DAYS
# ==========================
st.subheader("📈 Distribusi Umur Ticket (AGE DAYS)")

fig_age = px.histogram(
    dff,
    x="AGE_DAYS",
    color="STATUS",
    nbins=20,
    title="Distribusi Umur Ticket"
)
st.plotly_chart(fig_age, use_container_width=True)



# ==========================
# TABLE SUMMARY STATUS
# ==========================
st.subheader("📌 Ringkasan Ticket per Status")

count_status = dff["STATUS"].value_counts().reset_index()
count_status.columns = ["STATUS", "JUMLAH"]

st.dataframe(count_status, use_container_width=True)



# ==========================
# ANALISA PIC (Assignee)
# ==========================
st.subheader("👨‍🔧 Jumlah Ticket per PIC per Bulan")

# Tambah kolom bulan
dff["BULAN"] = dff["CREATE TICKET"].dt.to_period("M").astype(str)

pic_month = dff.pivot_table(
    index="ASSIGN DIVISION",
    columns="BULAN",
    values="NO",
    aggfunc="count",
    fill_value=0
)

st.dataframe(pic_month, use_container_width=True)

fig_pic = px.bar(
    pic_month.reset_index().melt(id_vars="ASSIGN DIVISION"),
    x="ASSIGN DIVISION",
    y="value",
    color="BULAN",
    title="Jumlah Ticket per PIC per Bulan",
    barmode="group"
)
st.plotly_chart(fig_pic, use_container_width=True)



# ==========================
# SARAN OTOMATIS
# ==========================
st.subheader("📝 Saran Perbaikan Kinerja PIC")

saran_list = []
for divisi in pic_month.index:
    total = pic_month.loc[divisi].sum()

    if total >= 15:
        saran = f"🔥 **{divisi}** menangani banyak ticket ({total}). Disarankan menambah engineer atau mempercepat penutupan."
    elif total >= 8:
        saran = f"⚠ **{divisi}** memiliki beban kerja sedang ({total}). Perlu monitoring agar backlog tidak meningkat."
    else:
        saran = f"✔ **{divisi}** beban kerja rendah ({total}). Kinerja stabil."
    
    saran_list.append(saran)

for s in saran_list:
    st.write("-", s)



# ==========================
# GENERATE PPT
# ==========================
st.subheader("📥 Download Presentasi PPT")

def generate_ppt(df_summary, fig_age, fig_pic):
    prs = Presentation("template PPT Moratel.pptx")

    # SLIDE 1 - SUMMARY
    slide = prs.slides[0]
    body = slide.shapes.placeholders[1].text_frame
    body.text = "Ringkasan Ticket CODEX:\n"

    for _, row in df_summary.iterrows():
        body.text += f"- {row['STATUS']}: {row['JUMLAH']}\n"

    # SLIDE 2 - Grafik AGE DAYS
    img_stream1 = io.BytesIO()
    fig_age.write_image(img_stream1, format="png")
    img_stream1.seek(0)

    slide2 = prs.slides.add_slide(prs.slide_layouts[5])
    slide2.shapes.add_picture(img_stream1, Inches(1), Inches(1), width=Inches(8))

    # SLIDE 3 - Grafik PIC
    img_stream2 = io.BytesIO()
    fig_pic.write_image(img_stream2, format="png")
    img_stream2.seek(0)

    slide3 = prs.slides.add_slide(prs.slide_layouts[5])
    slide3.shapes.add_picture(img_stream2, Inches(1), Inches(1), width=Inches(8))

    # SAVE PPT TO BYTES
    output = io.BytesIO()
    prs.save(output)
    return output


if st.button("Generate PPT"):
    ppt = generate_ppt(count_status, fig_age, fig_pic)
    st.success("Berhasil generate PPT!")

    st.download_button(
        label="📥 Download PPT",
        data=ppt,
        file_name="Analisa Ticket CODEX.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
