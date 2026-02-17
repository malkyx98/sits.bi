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
# Page Config & Style
# -------------------------------
st.set_page_config(page_title="🤖 AI-Powered Global Excel Analyzer", layout="wide")
st.markdown("""
    <style>
        body {
            background-color: #1e1e2f;
            color: #ffffff;
        }
        .stButton>button {
            background-color: #ff6f61;
            color: white;
        }
        .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
            color: #00d4ff;
        }
        .stDataFrame {
            background-color: #2a2a3c;
            color: #ffffff;
        }
    </style>
""", unsafe_allow_html=True)

st.title("🤖 AI-Powered Global Excel Analyzer")
st.markdown("Upload **ANY Excel file** → Get instant AI-driven KPIs, tables & visualizations")

# -------------------------------
# File Upload
# -------------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
if uploaded_file is None:
    st.info("📂 Please upload an Excel file to analyze.")
    st.stop()

df = pd.read_excel(uploaded_file)
st.write(f"Uploaded `{uploaded_file.name}`: Rows = {df.shape[0]}, Columns = {df.shape[1]}")

# -------------------------------
# Auto Detect Column Types
# -------------------------------
numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
date_cols = df.select_dtypes(include='datetime').columns.tolist()
if not date_cols:
    # Try to parse any object column as date
    for col in df.columns:
        try:
            df[col] = pd.to_datetime(df[col])
            date_cols.append(col)
        except:
            continue
categorical_cols = df.select_dtypes(include='object').columns.tolist()

id_cols = [c for c in df.columns if c not in numeric_cols + categorical_cols + date_cols]

st.subheader("🔍 Detected Column Roles")
st.markdown(f"**Numeric:** {numeric_cols if numeric_cols else 'None'}")
st.markdown(f"**Categorical:** {categorical_cols if categorical_cols else 'None'}")
st.markdown(f"**Date:** {date_cols if date_cols else 'None'}")
st.markdown(f"**ID / Name:** {id_cols if id_cols else 'None'}")

# -------------------------------
# KPI Summary
# -------------------------------
st.subheader("📊 Automatic KPI & Summary Tables")
if numeric_cols:
    kpi_df = df[numeric_cols].describe().T
    st.dataframe(kpi_df)
else:
    st.info("No numeric columns detected for KPIs.")

# -------------------------------
# Categorical Counts
# -------------------------------
st.subheader("📈 Categorical Value Counts")
for col in categorical_cols[:5]:  # limit to first 5 for display
    counts = df[col].value_counts().head(10)
    st.markdown(f"**{col} - Top Values**")
    st.bar_chart(counts)

# -------------------------------
# Heatmap for numeric correlations
# -------------------------------
if numeric_cols:
    st.subheader("🔥 Correlation Heatmap")
    corr = df[numeric_cols].corr()
    plt.figure(figsize=(8,6))
    sns.heatmap(corr, annot=True, cmap="coolwarm", fmt=".2f")
    st.pyplot(plt.gcf())
else:
    st.info("No numeric columns for correlation heatmap.")

# -------------------------------
# Export Processed Reports
# -------------------------------
st.subheader("📥 Download Processed Reports")
output = BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Processed_Data')
st.download_button("Download Excel", output.getvalue(), f"{uploaded_file.name.split('.')[0]}_processed.xlsx")

# -------------------------------
# Optional: PowerPoint Export
# -------------------------------
try:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(1))
    title.text = "Global Excel Analyzer Report"
    title.text_frame.paragraphs[0].font.size = Pt(24)
    # Just add first table as example
    if numeric_cols:
        rows, cols = kpi_df.shape
        table = slide.shapes.add_table(rows+1, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(1.5)).table
        # header
        for j, c in enumerate(kpi_df.columns):
            table.cell(0,j).text = str(c)
        # values
        for i in range(rows):
            for j in range(cols):
                table.cell(i+1,j).text = str(kpi_df.iloc[i,j])
except Exception as e:
    st.warning(f"PowerPoint export skipped: {str(e)}")

