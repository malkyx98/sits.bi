# app.py
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt

# -------------------------------
# PAGE CONFIG & STYLE
# -------------------------------
st.set_page_config(page_title="🤖 AI-Powered Global Excel Analyzer", layout="wide")

st.markdown("""
<style>
body {
    background-color: #1f1f2e;
    color: #f0f0f0;
    font-family: 'Segoe UI', sans-serif;
}
h1, h2, h3, h4 {
    color: #00b0f6;
}
.stButton>button {
    background-color: #00b0f6;
    color: white;
}
.stDownloadButton>button {
    background-color: #ff7f50;
    color: white;
}
</style>
""", unsafe_allow_html=True)

st.title("🤖 AI-Powered Global Excel Analyzer")
st.markdown("Upload ANY Excel file → Get instant AI-driven KPIs, trends, heatmaps & reports")

# -------------------------------
# FILE UPLOAD
# -------------------------------
uploaded_file = st.file_uploader("Upload Excel File (.xlsx, .xls)", type=["xlsx","xls"])

if uploaded_file is None:
    st.info("📂 Please upload an Excel file to start analysis.")
    st.stop()

# -------------------------------
# READ EXCEL FILE
# -------------------------------
try:
    df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Failed to read Excel file: {e}")
    st.stop()

st.success(f"File uploaded! Rows: {df.shape[0]}, Columns: {df.shape[1]}")

# -------------------------------
# DETECT COLUMN TYPES
# -------------------------------
numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
date_cols = df.select_dtypes(include=['datetime64', 'datetime']).columns.tolist()
categorical_cols = df.select_dtypes(include=['object']).columns.tolist()

# ID/Name columns heuristic
id_cols = [col for col in categorical_cols if df[col].nunique() > df.shape[0]*0.7]

# Status columns heuristic
status_cols = [col for col in categorical_cols if df[col].nunique() <= 10 and col not in id_cols]

st.markdown("### Detected Column Roles")
st.markdown(f"🔢 Numeric: {numeric_cols if numeric_cols else 'None'}")
st.markdown(f"🏷️ Categorical: { [c for c in categorical_cols if c not in id_cols+status_cols] or 'None' }")
st.markdown(f"📅 Date: {date_cols if date_cols else 'None'}")
st.markdown(f"🆔 ID/Name: {id_cols if id_cols else 'None'}")
st.markdown(f"✅ Status: {status_cols if status_cols else 'None'}")

# -------------------------------
# KPI TABLES
# -------------------------------
st.markdown("### 📊 Automatic KPI & Summary Tables")
kpi_table = pd.DataFrame()

for col in numeric_cols:
    kpi_table[col] = [df[col].sum(), df[col].mean(), df[col].max(), df[col].min()]
kpi_table.index = ['Total', 'Average', 'Max', 'Min']

if not kpi_table.empty:
    st.dataframe(kpi_table)
else:
    st.info("No numeric columns detected for KPI calculation.")

# -------------------------------
# VISUALIZATIONS
# -------------------------------
st.markdown("### 📈 Automatic Visualizations")

# Numeric columns: histogram & boxplot
for col in numeric_cols:
    st.subheader(f"Distribution - {col}")
    fig, ax = plt.subplots(figsize=(8,3))
    sns.histplot(df[col].dropna(), kde=True, ax=ax, color="#00b0f6")
    st.pyplot(fig)

# Correlation heatmap for numeric
if len(numeric_cols) > 1:
    st.subheader("Correlation Heatmap")
    corr = df[numeric_cols].corr()
    fig, ax = plt.subplots(figsize=(8,6))
    sns.heatmap(corr, annot=True, cmap="coolwarm", ax=ax)
    st.pyplot(fig)

# Categorical counts bar chart
for col in categorical_cols:
    st.subheader(f"Top Values - {col}")
    counts = df[col].value_counts().nlargest(10)
    fig = px.bar(counts, x=counts.index, y=counts.values, color=counts.values,
                 color_continuous_scale='Viridis')
    st.plotly_chart(fig, use_container_width=True)

# -------------------------------
# EXPORT REPORTS
# -------------------------------
st.markdown("### 📥 Download Processed Reports")

# Excel Export
excel_buffer = BytesIO()
with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Processed_Data', index=False)
st.download_button("Download Excel", excel_buffer.getvalue(), "processed_report.xlsx")

# PowerPoint Export
prs = Presentation()
slide_layout = prs.slide_layouts[5]
slide = prs.slides.add_slide(slide_layout)
slide.shapes.title.text = "Automatic KPI & Summary"

rows, cols_count = kpi_table.shape if not kpi_table.empty else (2,2)
table = slide.shapes.add_table(rows+1, cols_count, Inches(0.5), Inches(1.5), Inches(9), Inches(1)).table

# Header
if not kpi_table.empty:
    for j, col_name in enumerate(kpi_table.columns):
        table.cell(0,j).text = str(col_name)
    # Values
    for i in range(rows):
        for j in range(cols_count):
            val = kpi_table.iloc[i,j]
            table.cell(i+1,j).text = str(val)

ppt_buffer = BytesIO()
prs.save(ppt_buffer)
st.download_button("Download PowerPoint", ppt_buffer.getvalue(), "processed_report.pptx")


