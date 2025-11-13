import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import tempfile

st.set_page_config(page_title="Dashboard Ticket CODEX", layout="wide")

# ============================
# TITLE
# ============================
st.title("📊 Dashboard Analisa Ticket CODEX — Versi Lengkap & Export PPT")

st.info("Upload 2 file Excel (bulan ini & bulan lalu).")


# ============================
# UPLOAD FILE
# ============================
file1 = st.file_uploader("📁 Upload File CODEX Bulan 1 (.xlsx)", type=["xlsx"])
file2 = st.file_uploader("📁 Upload File CODEX Bulan 2 (.xlsx)", type=["xlsx"])
ppt_template = "template PPT Moratel.pptx"

if file1 is None or file2 is None:
    st.warning("Upload **dua file Excel** terlebih dahulu.")
    st.stop()

# ============================
# LOAD DATA
# ============================
df1 = pd.read_excel(file1)
df2 = pd.read_excel(file2)

df = pd.concat([df1, df2], ignore_index=True)
df.columns = [c.strip() for c in df.columns]

# Normalisasi kolom wajib
required_cols = ["NO-TICKET", "HOSTNAME", "INTERFACE", "STATUS",
                 "ASSIGN DIVISION", "CREATE TICKET"]

for col in required_cols:
    if col not in df.columns:
        st.error(f"Kolom **{col}** tidak ditemukan dalam file Excel.")
        st.stop()

df["CREATE TICKET"] = pd.to_datetime(df["CREATE TICKET"], errors="coerce")
df["AGE_DAYS"] = (datetime.now() - df["CREATE TICKET"]).dt.days

df["SUB_DIVISI"] = df["ASSIGN DIVISION"].str.split("-").str[-1].str.strip()
df["PIC"] = df["ASSIGN DIVISION"].str.split("-").str[0].str.strip()

if "STATUS.1" in df.columns:
    df["STATUS_TICKET"] = df["STATUS.1"]
else:
    df["STATUS_TICKET"] = df["STATUS"]


# ============================
# SUMMARY NUMBERS
# ============================
st.subheader("📌 Ringkasan Data Ticket")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Ticket", len(df))
c2.metric("Critical", df["STATUS"].str.contains("Critical", case=False).sum())
c3.metric("Warning", df["STATUS"].str.contains("Warning", case=False).sum())
c4.metric("Average Age (Days)", round(df["AGE_DAYS"].mean(), 1))


# ============================
# GRAFIK STATUS
# ============================
st.subheader("📊 Grafik Status Ticket")

count_status = df["STATUS"].value_counts().reset_index()
count_status.columns = ["STATUS", "JUMLAH"]

fig_status = px.bar(count_status, x="STATUS", y="JUMLAH",
                    color="STATUS", title="Distribusi Kategori Status")

st.plotly_chart(fig_status, use_container_width=True)


# ============================
# GRAFIK AGE DAYS
# ============================
st.subheader("⏳ Distribusi Umur Ticket (Age Days)")

fig_age = px.histogram(df, x="AGE_DAYS", nbins=30, color="STATUS",
                       title="Distribusi AGE Ticket")

st.plotly_chart(fig_age, use_container_width=True)


# ============================
# PERFORMANCE PIC
# ============================
st.subheader("🧑‍💻 Grafik Performance PIC per Bulan")

df["MONTH"] = df["CREATE TICKET"].dt.month_name()

pic_perf = df.groupby(["PIC", "MONTH"]).size().reset_index(name="JUMLAH")

fig_pic = px.bar(pic_perf, x="PIC", y="JUMLAH",
                 color="MONTH", barmode="group",
                 title="Performance PIC per Bulan")

st.plotly_chart(fig_pic, use_container_width=True)


# ============================
# TABEL PIC
# ============================
st.subheader("📋 Tabel Performance PIC")

pic_table = df.groupby("PIC").agg(
    TOTAL=("PIC", "count"),
    OPEN=("STATUS_TICKET", lambda x: (x == "Open").sum()),
    CLOSED=("STATUS_TICKET", lambda x: (x == "Close").sum()),
    AVG_AGE=("AGE_DAYS", "mean"),
    MAX_AGE=("AGE_DAYS", "max")
).reset_index()

pic_table["SLA_%"] = round((pic_table["CLOSED"] / pic_table["TOTAL"]) * 100, 1)
pic_table["AVG_AGE"] = pic_table["AVG_AGE"].round(1)

st.dataframe(pic_table, use_container_width=True)


# ============================
# GENERATE PPT FUNCTION
# ============================
def generate_ppt(summary_df, pic_df, fig1, fig2, fig3):
    prs = Presentation(ppt_template)

    # Convert plotly figure → PNG
    temp_imgs = []
    for fig in [fig1, fig2, fig3]:
        tmpfile = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        fig.write_image(tmpfile.name)
        temp_imgs.append(tmpfile.name)

    # Slide 1 is cover (template)
    
    # Slide 2 – Summary
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Summary Data Ticket"
    body = slide.shapes.placeholders[1]
    body.text = summary_df.to_string(index=False)

    # Slide 3 – Tabel PIC
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Performance PIC"
    slide.shapes.placeholders[1].text = pic_df.to_string(index=False)

    # Slide 4 – Grafik Status
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Grafik Status Ticket"
    slide.shapes.add_picture(temp_imgs[0], Inches(1), Inches(1), width=Inches(8))

    # Slide 5 – Grafik Age
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Distribusi Umur Ticket"
    slide.shapes.add_picture(temp_imgs[1], Inches(1), Inches(1), width=Inches(8))

    # Slide 6 – Grafik PIC
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Grafik PIC per Bulan"
    slide.shapes.add_picture(temp_imgs[2], Inches(1), Inches(1), width=Inches(8))

    output = BytesIO()
    prs.save(output)
    return output


# ============================
# ADD DOWNLOAD PPT BUTTON
# ============================
st.subheader("📥 Download Presentasi PPT")

if st.button("Generate PPT"):
    ppt_data = generate_ppt(summary=count_status,
                            summary_df=count_status,
                            pic_df=pic_table,
                            fig1=fig_status,
                            fig2=fig_age,
                            fig3=fig_pic)

    st.download_button(
        label="📩 Download PPT Laporan CODEx",
        data=ppt_data,
        file_name="Laporan-CODEx-Moratel.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

