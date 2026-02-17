# app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px

# -----------------------------
# PAGE CONFIG
# -----------------------------
st.set_page_config(page_title="🌈 AI Colorful Excel Analyzer", layout="wide")

# -----------------------------
# CUSTOM COLORFUL CSS
# -----------------------------
st.markdown("""
<style>
body {
    background: linear-gradient(to right, #f9f9f9, #e0f7fa);
    color: #000;
    font-family: 'Segoe UI', sans-serif;
}
h1, h2, h3 {
    font-weight: bold;
    background: linear-gradient(90deg, #ff4b4b, #ffa64b, #4b6eff);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}
.stButton>button {
    background: linear-gradient(90deg, #ff4b4b, #ffa64b, #4b6eff);
    color: #fff;
    font-weight: bold;
    border-radius: 15px;
}
.kpi-card {
    background: linear-gradient(135deg, #f7971e, #ffd200, #6a11cb);
    color: #fff;
    padding: 15px;
    border-radius: 15px;
    text-align: center;
    margin-bottom: 10px;
}
.dataframe th {
    background: #ff4b4b;
    color: white;
}
.dataframe td {
    background: #fff3e0;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# TITLE
# -----------------------------
st.title("🌈 AI-Powered Colorful Global Excel Analyzer")
st.markdown("Upload ANY Excel file → Get instant AI-driven KPIs, trends, and rainbow charts!")

# -----------------------------
# FILE UPLOAD
# -----------------------------
uploaded_file = st.file_uploader("📂 Upload Excel File (.xlsx, .xls)", type=['xlsx','xls'])

if uploaded_file is None:
    st.info("Please upload an Excel file to start analysis.")
    st.stop()

try:
    df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Failed to read Excel file: {e}")
    st.stop()

st.success(f"File uploaded! Rows: {df.shape[0]}, Columns: {df.shape[1]}")

# -----------------------------
# DETECT COLUMN ROLES (AI-style)
# -----------------------------
def detect_column_roles(df):
    numeric_cols, date_cols, categorical_cols, id_cols, status_cols = [], [], [], [], []
    for col in df.columns:
        s = df[col].dropna()
        if pd.api.types.is_numeric_dtype(s):
            numeric_cols.append(col)
        elif pd.api.types.is_datetime64_any_dtype(s):
            date_cols.append(col)
        elif any(k in col.lower() for k in ['date','time']):
            try:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                date_cols.append(col)
            except:
                categorical_cols.append(col)
        elif any(k in col.lower() for k in ['id','name','ref','user']):
            id_cols.append(col)
        elif any(k in col.lower() for k in ['status','sla','done','closed','pending']):
            status_cols.append(col)
        else:
            categorical_cols.append(col)
    return numeric_cols, categorical_cols, date_cols, id_cols, status_cols

numeric_cols, categorical_cols, date_cols, id_cols, status_cols = detect_column_roles(df)

# -----------------------------
# SHOW COLUMN ROLES
# -----------------------------
st.markdown("### 🔍 Column Roles")
col1, col2 = st.columns(2)
with col1:
    st.markdown(f"**🔢 Numeric:** {numeric_cols if numeric_cols else 'None'}")
    st.markdown(f"**🏷️ Categorical:** {categorical_cols if categorical_cols else 'None'}")
with col2:
    st.markdown(f"**📅 Date:** {date_cols if date_cols else 'None'}")
    st.markdown(f"**🆔 ID/Name:** {id_cols if id_cols else 'None'}")
    st.markdown(f"**✅ Status:** {status_cols if status_cols else 'None'}")

# -----------------------------
# KPI CARDS WITH RAINBOW GRADIENTS
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
            st.markdown(
                f"<div class='kpi-card'><h4>{col}</h4>Total: {total}<br>Avg: {avg:.2f}<br>Max: {mx}<br>Min: {mn}</div>",
                unsafe_allow_html=True
            )
else:
    st.info("No numeric columns found for KPIs.")

# -----------------------------
# CATEGORICAL SUMMARY WITH COLORFUL CHARTS
# -----------------------------
st.markdown("### 🏷️ Categorical Analysis")
palette = px.colors.qualitative.Plotly
for idx, col in enumerate(categorical_cols):
    counts = df[col].value_counts().head(10)
    st.markdown(f"**Top 10 values for {col}:**")
    st.dataframe(counts)
    if len(counts) > 0:
        try:
            fig = px.bar(counts, x=counts.index, y=counts.values, 
                         labels={'x': col, 'y': 'Count'},
                         color=counts.index,
                         color_discrete_sequence=palette)
            st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.warning(f"Skipping chart for {col}: {e}")

# -----------------------------
# DATE TRENDS WITH RAINBOW LINES
# -----------------------------
st.markdown("### 📆 Date Trends")
rainbow_colors = px.colors.qualitative.D3
for col in date_cols:
    df[col+'_dt'] = pd.to_datetime(df[col], errors='coerce')
    if df[col+'_dt'].notna().sum() > 0:
        trend = df.groupby(df[col+'_dt'].dt.to_period('M')).size()
        st.markdown(f"**Monthly trend for {col}:**")
        try:
            fig = px.line(trend, x=trend.index.astype(str), y=trend.values, markers=True,
                          labels={'x':'Month','y':'Count'},
                          line_shape='spline',
                          color_discrete_sequence=rainbow_colors)
            st.plotly_chart(fig, use_container_width=True)
        except:
            pass

# -----------------------------
# DOWNLOAD PROCESSED REPORT
# -----------------------------
st.markdown("### 📥 Download Processed Report")
output = BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Processed_Data', index=False)
output.seek(0)
st.download_button("Download Excel Report", output.getvalue(), "AI_Colorful_Excel_Report.xlsx")

st.success("✅ Analysis Complete! Enjoy your rainbow-powered AI insights 🌈")
