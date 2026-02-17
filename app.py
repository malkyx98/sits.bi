import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# -----------------------------
# PAGE CONFIG
# -----------------------------
st.set_page_config(page_title="Global Excel Analyzer v3", layout="wide")
st.title("📊 Global Excel Analyzer v3")
st.write("Upload ANY Excel file → Get interactive KPIs and visualizations automatically")

# -----------------------------
# FILE UPLOAD
# -----------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx","xls"])
if uploaded_file is None:
    st.info("Please upload an Excel file to start analysis.")
    st.stop()

# Read Excel safely
try:
    df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Failed to read Excel file: {e}")
    st.stop()

st.success(f"File uploaded! Rows: {df.shape[0]}, Columns: {df.shape[1]}")

# -----------------------------
# CLEAN COLUMN NAMES
# -----------------------------
df.columns = [str(col).strip() if col is not None else f"Column_{i}" 
              for i, col in enumerate(df.columns)]

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
        if len(series) == 0:
            continue
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
        if numeric_series.notna().any():
            numeric_cols.append(col)
        else:
            categorical_cols.append(col)
    return numeric_cols, categorical_cols, date_cols

numeric_cols, categorical_cols, date_cols = detect_columns(df)

st.subheader("Detected Column Types")
c1,c2,c3 = st.columns(3)
c1.info(f"🔢 Numeric Columns:\n{numeric_cols if numeric_cols else 'None'}")
c2.warning(f"🏷️ Categorical Columns:\n{categorical_cols if categorical_cols else 'None'}")
c3.success(f"📅 Date Columns:\n{date_cols if date_cols else 'None'}")

# -----------------------------
# INTERACTIVE FILTERS
# -----------------------------
st.sidebar.subheader("🔎 Filters")

# Numeric filters
for col in numeric_cols:
    min_val = float(df[col].min())
    max_val = float(df[col].max())
    step = (max_val - min_val) / 100 if max_val > min_val else 1
    df = df[(df[col] >= st.sidebar.slider(f"{col} min", min_val, max_val, min_val, step)) &
            (df[col] <= st.sidebar.slider(f"{col} max", min_val, max_val, max_val, step))]

# Categorical filters
for col in categorical_cols:
    options = df[col].dropna().unique().tolist()
    selected = st.sidebar.multiselect(f"{col} filter", options, default=options)
    df = df[df[col].isin(selected)]

# Date filters
for col in date_cols:
    try:
        df[col+'_parsed'] = pd.to_datetime(df[col], errors='coerce')
        min_date = df[col+'_parsed'].min()
        max_date = df[col+'_parsed'].max()
        selected_range = st.sidebar.date_input(f"{col} range", [min_date, max_date])
        if len(selected_range) == 2:
            df = df[(df[col+'_parsed'] >= pd.to_datetime(selected_range[0])) &
                    (df[col+'_parsed'] <= pd.to_datetime(selected_range[1]))]
    except:
        continue

# -----------------------------
# NUMERIC KPIs
# -----------------------------
st.subheader("📊 Numeric KPIs")
if numeric_cols:
    for col in numeric_cols:
        numeric_series = pd.to_numeric(df[col], errors='coerce')
        k1,k2,k3,k4 = st.columns(4)
        k1.metric(f"{col} Total", round(numeric_series.sum(),2))
        k2.metric(f"{col} Avg", round(numeric_series.mean(),2))
        k3.metric(f"{col} Max", numeric_series.max())
        k4.metric(f"{col} Min", numeric_series.min())
else:
    st.info("No numeric columns detected.")

# -----------------------------
# CATEGORICAL COUNTS
# -----------------------------
st.subheader("📈 Categorical Counts (Top 10)")
if categorical_cols:
    for col in categorical_cols:
        counts = df[col].dropna().astype(str).value_counts().head(10)
        if counts.empty:
            continue
        st.markdown(f"**{col}** - Top 10 values")
        st.bar_chart(counts)
else:
    st.info("No categorical columns detected.")

# -----------------------------
# DATE TRENDS
# -----------------------------
st.subheader("📆 Date Trends")
if date_cols:
    for col in date_cols:
        if col+'_parsed' not in df.columns:
            df[col+'_parsed'] = pd.to_datetime(df[col], errors='coerce')
        trend = df.groupby(df[col+'_parsed'].dt.to_period("M")).size().reset_index(name='Count')
        trend[col+'_parsed'] = trend[col+'_parsed'].astype(str)
        if trend.empty:
            continue
        fig = px.line(trend, x=col+'_parsed', y='Count', title=f"Trend over time ({col})")
        st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No date columns detected.")

# -----------------------------
# CORRELATION HEATMAP
# -----------------------------
st.subheader("🔥 Correlation Heatmap")
if len(numeric_cols) > 1:
    corr = df[numeric_cols].apply(pd.to_numeric, errors='coerce').corr()
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

