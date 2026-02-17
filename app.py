# app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px

# -----------------------------
# PAGE CONFIG
# -----------------------------
st.set_page_config(page_title="🤖 AI-Powered Global Excel Analyzer", layout="wide")

# -----------------------------
# CUSTOM CSS - AI THEME
# -----------------------------
st.markdown("""
<style>
/* Background & main text */
body {
    background-color: #0e1117;
    color: #ffffff;
}

/* Header style */
h1, h2, h3 {
    color: #00ffff;
    font-family: 'Segoe UI', sans-serif;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: #111827;
    color: #ffffff;
}

/* Buttons */
.stButton>button {
    background-color: #00ffff;
    color: #000000;
    font-weight: bold;
}

/* KPI Cards */
.kpi-card {
    background-color: #1f2937;
    padding: 10px 15px;
    border-radius: 10px;
    margin: 5px;
    display: inline-block;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# APP TITLE
# -----------------------------
st.title("🤖 AI-Powered Global Excel Analyzer")
st.markdown("Upload ANY Excel file → Get instant AI-driven KPIs, trends, and reports")

# -----------------------------
# FILE UPLOAD
# -----------------------------
uploaded_file = st.file_uploader("📂 Upload Excel File (.xlsx, .xls)", type=['xlsx','xls'])

if uploaded_file is None:
    st.info("Please upload an Excel file to proceed.")
    st.stop()

# Read Excel
try:
    df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Failed to read Excel file: {e}")
    st.stop()

st.success(f"File uploaded! Rows: {df.shape[0]}, Columns: {df.shape[1]}")

# -----------------------------
# COLUMN ROLE DETECTION
# -----------------------------
def detect_column_roles(df):
    numeric_cols, date_cols, categorical_cols, id_cols, status_cols = [], [], [], [], []
    for col in df.columns:
        s = df[col].dropna()
        # Numeric detection
        if pd.api.types.is_numeric_dtype(s):
            numeric_cols.append(col)
        # Date detection
        elif pd.api.types.is_datetime64_any_dtype(s):
            date_cols.append(col)
        elif any(keyword in col.lower() for keyword in ['date','start','end','time']):
            try:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                date_cols.append(col)
            except:
                categorical_cols.append(col)
        # ID/Name detection
        elif any(keyword in col.lower() for keyword in ['id','name','ref','user']):
            id_cols.append(col)
        # Status detection
        elif any(keyword in col.lower() for keyword in ['status','sla','done','closed','pending']):
            status_cols.append(col)
        else:
            categorical_cols.append(col)
    return numeric_cols, categorical_cols, date_cols, id_cols, status_cols

numeric_cols, categorical_cols, date_cols, id_cols, status_cols = detect_column_roles(df)

# -----------------------------
# COLUMN ROLES DISPLAY
# -----------------------------
st.markdown("### 🔍 Detected Column Roles")
col1, col2 = st.columns(2)
with col1:
    st.markdown(f"**🔢 Numeric:** {numeric_cols if numeric_cols else 'None'}")
    st.markdown(f"**🏷️ Categorical:** {categorical_cols if categorical_cols else 'None'}")
with col2:
    st.markdown(f"**📅 Date:** {date_cols if date_cols else 'None'}")
    st.markdown(f"**🆔 ID/Name:** {id_cols if id_cols else 'None'}")
    st.markdown(f"**✅ Status:** {status_cols if status_cols else 'None'}")

# -----------------------------
# KPI CARDS
# -----------------------------
st.markdown("### 📊 Numeric KPIs")
if numeric_cols:
    kpi_cols = st.columns(len(numeric_cols))
    for i, col in enumerate(numeric_cols):
        total = df[col].sum()
        avg = df[col].mean()
        mx = df[col].max()
        mn = df[col].min()
        with kpi_cols[i]:
            st.markdown(f"<div class='kpi-card'><h4>{col}</h4>Total: {total}<br>Avg: {avg:.2f}<br>Max: {mx}<br>Min: {mn}</div>", unsafe_allow_html=True)
else:
    st.info("No numeric columns detected for KPI calculation.")

# -----------------------------
# CATEGORICAL SUMMARY
# -----------------------------
st.markdown("### 🏷️ Categorical Summary")
for col in categorical_cols:
    counts = df[col].value_counts().head(10)
    st.markdown(f"**Top values for {col}:**")
    st.dataframe(counts)
    if len(counts) > 0:
        try:
            fig = px.bar(counts, x=counts.index, y=counts.values, labels={'x': col, 'y': 'Count'}, color=counts.values)
            st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.warning(f"Skipping chart for {col}: {e}")

# -----------------------------
# DATE TRENDS
# -----------------------------
st.markdown("### 📆 Date Trends")
for col in date_cols:
    df[col+'_date'] = pd.to_datetime(df[col], errors='coerce')
    if df[col+'_date'].notna().sum() > 0:
        trend = df.groupby(df[col+'_date'].dt.to_period('M')).size()
        st.markdown(f"**Monthly Trend for {col}:**")
        try:
            fig = px.line(trend, x=trend.index.astype(str), y=trend.values, markers=True,
                          labels={'x':'Month', 'y':'Count'}, line_shape='spline')
            st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.warning(f"Skipping trend chart for {col}: {e}")

# -----------------------------
# DOWNLOAD PROCESSED REPORT
# -----------------------------
st.markdown("### 📥 Download Processed Reports")
output = BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Processed_Data', index=False)
output.seek(0)
st.download_button("Download Excel Report", output.getvalue(), "AI_Excel_Report.xlsx")

st.success("✅ AI Analysis Complete! Your KPIs, trends, and summaries are ready.")

