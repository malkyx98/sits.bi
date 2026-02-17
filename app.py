import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# -----------------------------
# PAGE CONFIG
# -----------------------------
st.set_page_config(page_title="Universal Excel Analyzer", layout="wide")
st.title("📊 Universal Excel Auto Analyzer")
st.write("Upload ANY Excel file → Get instant analytics")

# -----------------------------
# FILE UPLOAD
# -----------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
if uploaded_file is None:
    st.info("Please upload an Excel file to start analysis.")
    st.stop()

df = pd.read_excel(uploaded_file)
st.success("File uploaded successfully!")

# -----------------------------
# DATA PREVIEW
# -----------------------------
with st.expander("Preview Data"):
    st.dataframe(df.head())

# -----------------------------
# SMART COLUMN DETECTION
# -----------------------------
def detect_columns_smart(df):
    detected = {"date_cols": [], "numeric_cols": [], "categorical_cols": []}
    for col in df.columns:
        series = df[col].dropna()
        # Detect date columns
        try:
            parsed = pd.to_datetime(series, errors='coerce')
            if parsed.notna().sum() / len(series) > 0.6:
                detected["date_cols"].append(col)
                continue
        except:
            pass
        # Detect numeric columns (only if > 80% values are numbers)
        numeric_series = pd.to_numeric(series, errors='coerce')
        numeric_ratio = numeric_series.notna().sum() / len(series)
        if numeric_ratio > 0.8:
            detected["numeric_cols"].append(col)
        else:
            detected["categorical_cols"].append(col)
    return detected

detected = detect_columns_smart(df)

# -----------------------------
# USER OVERRIDE FOR COLUMN TYPES
# -----------------------------
st.subheader("🔍 Detected Column Types")
c1, c2, c3 = st.columns(3)

numeric_cols = st.multiselect("Select Numeric Columns", detected["numeric_cols"], default=detected["numeric_cols"])
categorical_cols = st.multiselect("Select Categorical Columns", detected["categorical_cols"], default=detected["categorical_cols"])
date_cols = st.multiselect("Select Date Columns", detected["date_cols"], default=detected["date_cols"])

c1.success(f"🔢 Numeric Columns\n\n{numeric_cols}")
c2.warning(f"🏷️ Categorical Columns\n\n{categorical_cols}")
c3.info(f"📅 Date Columns\n\n{date_cols}")

# -----------------------------
# AUTO KPI GENERATION
# -----------------------------
st.markdown("## 📊 Automatic KPIs")

def auto_kpis(df, numeric_cols):
    kpis = {}
    for col in numeric_cols:
        kpis[col] = {
            "Total": df[col].sum(),
            "Average": df[col].mean(),
            "Max": df[col].max(),
            "Min": df[col].min()
        }
    return kpis

kpis = auto_kpis(df, numeric_cols)
for col, stats in kpis.items():
    st.markdown(f"### {col}")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total", round(stats["Total"], 2))
    k2.metric("Average", round(stats["Average"], 2))
    k3.metric("Max", stats["Max"])
    k4.metric("Min", stats["Min"])

# -----------------------------
# AUTOMATIC CHARTS
# -----------------------------
st.markdown("## 📈 Automatic Visualizations")

# TIME SERIES CHART
if date_cols and numeric_cols:
    date_col = st.selectbox("Select Date Column for Trend", date_cols)
    num_col = st.selectbox("Select Numeric Column for Trend", numeric_cols, key="trend")
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    fig = px.line(df.sort_values(date_col), x=date_col, y=num_col, title="Trend Over Time")
    st.plotly_chart(fig, use_container_width=True)

# CATEGORY ANALYSIS
if categorical_cols and numeric_cols:
    st.markdown("### Category Analysis")
    cat_col = st.selectbox("Select Category Column", categorical_cols)
    num_col2 = st.selectbox("Select Numeric Column for Category Analysis", numeric_cols, key="cat")
    grouped = df.groupby(cat_col)[num_col2].sum().reset_index()
    fig2 = px.bar(grouped, x=cat_col, y=num_col2, title="Category Distribution")
    st.plotly_chart(fig2, use_container_width=True)

# CORRELATION HEATMAP
if len(numeric_cols) > 1:
    st.markdown("## 🔥 Correlation Matrix")
    corr = df[numeric_cols].corr()
    fig3 = px.imshow(corr, text_auto=True, title="Correlation Heatmap")
    st.plotly_chart(fig3, use_container_width=True)

# -----------------------------
# RAW DATA DOWNLOAD
# -----------------------------
st.markdown("## 📥 Download Processed Data")
csv = df.to_csv(index=False).encode("utf-8")
st.download_button("Download CSV", csv, "processed_data.csv")
