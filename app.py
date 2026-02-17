import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# -----------------------------
# PAGE CONFIG
# -----------------------------
st.set_page_config(page_title="Global Excel Analyzer", layout="wide")
st.title("📊 Global Excel Analyzer")
st.write("Upload ANY Excel file → Get instant KPIs and visualizations automatically")

# -----------------------------
# FILE UPLOAD
# -----------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx","xls"])
if uploaded_file is None:
    st.info("Please upload an Excel file to start analysis.")
    st.stop()

df = pd.read_excel(uploaded_file)
st.success(f"File uploaded! Rows: {df.shape[0]}, Columns: {df.shape[1]}")

# -----------------------------
# DATA PREVIEW
# -----------------------------
with st.expander("Preview Data"):
    st.dataframe(df.head())

# -----------------------------
# SMART COLUMN DETECTION
# -----------------------------
def detect_columns(df):
    numeric_cols, categorical_cols, date_cols = [], [], []
    for col in df.columns:
        series = df[col].dropna()
        # Date detection
        try:
            parsed = pd.to_datetime(series, errors='coerce')
            if parsed.notna().sum() / len(series) > 0.5:
                date_cols.append(col)
                continue
        except:
            pass
        # Numeric detection
        numeric_series = pd.to_numeric(series, errors='coerce')
        if numeric_series.notna().sum() / len(series) > 0.7:
            numeric_cols.append(col)
        else:
            categorical_cols.append(col)
    return numeric_cols, categorical_cols, date_cols

numeric_cols, categorical_cols, date_cols = detect_columns(df)

st.subheader("Detected Column Types")
c1,c2,c3 = st.columns(3)
c1.info(f"🔢 Numeric Columns:\n{numeric_cols}")
c2.warning(f"🏷️ Categorical Columns:\n{categorical_cols}")
c3.success(f"📅 Date Columns:\n{date_cols}")

# -----------------------------
# AUTOMATIC KPIs FOR NUMERIC COLUMNS
# -----------------------------
st.subheader("📊 Numeric KPIs")
if numeric_cols:
    for col in numeric_cols:
        col_total = df[col].sum()
        col_avg = df[col].mean()
        col_max = df[col].max()
        col_min = df[col].min()
        k1,k2,k3,k4 = st.columns(4)
        k1.metric(f"{col} Total", round(col_total,2))
        k2.metric(f"{col} Avg", round(col_avg,2))
        k3.metric(f"{col} Max", col_max)
        k4.metric(f"{col} Min", col_min)
else:
    st.info("No numeric columns detected.")

# -----------------------------
# FREQUENCY COUNTS FOR CATEGORICAL
# -----------------------------
st.subheader("📈 Categorical Counts")
if categorical_cols:
    for col in categorical_cols:
        counts = df[col].value_counts().head(10)
        st.markdown(f"**{col}** - Top 10 values")
        st.bar_chart(counts)
else:
    st.info("No categorical columns detected.")

# -----------------------------
# TRENDS FOR DATE COLUMNS
# -----------------------------
st.subheader("📆 Date Trends")
if date_cols:
    for col in date_cols:
        df[col+'_parsed'] = pd.to_datetime(df[col], errors='coerce')
        trend = df.groupby(df[col+'_parsed'].dt.to_period("M")).size().reset_index(name='Count')
        trend[col+'_parsed'] = trend[col+'_parsed'].astype(str)
        fig = px.line(trend, x=col+'_parsed', y='Count', title=f"Trend over time ({col})")
        st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No date columns detected.")

# -----------------------------
# CORRELATION HEATMAP FOR NUMERIC
# -----------------------------
st.subheader("🔥 Correlation Heatmap")
if len(numeric_cols) > 1:
    corr = df[numeric_cols].corr()
    fig_corr = px.imshow(corr, text_auto=True, color_continuous_scale='RdBu_r', title="Correlation Heatmap")
    st.plotly_chart(fig_corr, use_container_width=True)
else:
    st.info("Need at least 2 numeric columns for correlation analysis.")

# -----------------------------
# DOWNLOAD PROCESSED DATA
# -----------------------------
st.subheader("📥 Download Processed Data")
csv = df.to_csv(index=False).encode("utf-8")
st.download_button("Download CSV", csv, "processed_data.csv")
