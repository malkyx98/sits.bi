import streamlit as st
import pandas as pd
import plotly.express as px
import os
import base64
from src.data_processing import process_data

# ==================================================
# 1. SYSTEM CONFIGURATION & STYLES
# ==================================================
st.set_page_config(
    page_title="Performance PRO | Analytics",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Initialize Session States
if "current_page" not in st.session_state:
    st.session_state.current_page = "OVERVIEW"
if "file_ref" not in st.session_state:
    st.session_state.file_ref = None

# Advanced CSS for Premium Look
st.markdown("""
<style>
    header[data-testid="stHeader"] { visibility: hidden; }
    [data-testid="stSidebar"] { display: none; }
    
    .stApp {
        background: linear-gradient(to bottom, #001f4d, #ffffff) !important;
        background-attachment: fixed;
    }

    .block-container { padding-top: 1rem !important; }
    
    .header-wrapper {
        display: flex; align-items: center; justify-content: flex-start;
        gap: 15px; width: 100%; padding: 10px 0;
    }

    .nav-logo-text { 
        font-size: 22px; font-weight: 700; color: #FFFFFF; 
        letter-spacing: -0.5px;
    }
    .nav-logo-text span { color: #FF8C00; }
    
    /* KPI Cards */
    .kpi-box {
        background: rgba(255, 255, 255, 0.9); 
        backdrop-filter: blur(10px);
        border-radius: 12px; padding: 15px; text-align: center; 
        border-bottom: 4px solid #FF8C00;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.1);
        transition: transform 0.3s ease;
    }
    .kpi-box:hover { transform: translateY(-5px); }
    .kpi-label { font-size: 11px; font-weight: 700; color: #666; text-transform: uppercase; }
    .kpi-value { font-size: 28px; font-weight: 800; color: #001f4d; margin: 2px 0; }
    .kpi-range { font-size: 10px; color: #999; font-style: italic; }

    /* Portal Cards */
    .portal-card {
        background: white; padding: 25px; border-radius: 15px; 
        text-align: center; box-shadow: 0 10px 20px rgba(0,0,0,0.05);
        border: 1px solid #f0f0f0;
        min-height: 220px;
    }
    .circle-icon {
        width: 60px; height: 60px; border-radius: 50%; 
        display: flex; align-items: center; justify-content: center; 
        font-size: 22px; font-weight: 800; color: white; margin: 0 auto 15px auto;
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
    }

    /* Professional Buttons */
    .stButton > button {
        background-color: #FF8C00 !important; border: none !important;
        color: white !important; border-radius: 8px; font-size: 12px !important;
        font-weight: 700 !important; transition: all 0.3s ease !important;
    }
    .stButton > button:hover { 
        background-color: #001f4d !important;
        box-shadow: 0 4px 12px rgba(255, 140, 0, 0.4) !important;
    }

    /* Status Section Headers */
    .status-header-complete {
        background-color: #2e7d32;
        color: white;
        padding: 10px 15px;
        border-radius: 8px 8px 0 0;
        margin-top: 25px;
        font-weight: bold;
        font-size: 14px;
    }
    .status-header-pending {
        background-color: #d32f2f;
        color: white;
        padding: 10px 15px;
        border-radius: 8px 8px 0 0;
        margin-top: 25px;
        font-weight: bold;
        font-size: 14px;
    }
</style>
""", unsafe_allow_html=True)

def get_image_base64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except: return None

# ==================================================
# 2. HEADER & NAVIGATION
# ==================================================
col_logo, col_menu = st.columns([2.5, 3.5])

with col_logo:
    img_b64 = get_image_base64("logo.png")
    logo_html = f'<img src="data:image/png;base64,{img_b64}" width="140">' if img_b64 else ""
    st.markdown(f"""
        <div class="header-wrapper">
            {logo_html}
            <div class="nav-logo-text">Performance <span>PRO</span></div>
        </div>
    """, unsafe_allow_html=True)

with col_menu:
    if st.session_state.file_ref:
        m = st.columns(4)
        if m[0].button("OVERVIEW"): st.session_state.current_page = "OVERVIEW"
        if m[1].button("DEPARTMENT"): st.session_state.current_page = "DEPT"
        if m[2].button("SUPERVISOR"): st.session_state.current_page = "SUP"
        if m[3].button("NEW UPLOAD"):
            st.session_state.file_ref = None
            st.rerun()

# ==================================================
# 3. DATA LOADING & GLOBAL DEFINITIONS
# ==================================================
if st.session_state.file_ref is None:
    st.markdown("<div style='height: 100px;'></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown("<h2 style='text-align:center; color:white;'>Intelligence Portal</h2>", unsafe_allow_html=True)
        up = st.file_uploader("", type=["xlsx"])
        if up:
            st.session_state.file_ref = up
            st.rerun()
    st.stop()

@st.cache_data
def load_and_clean(file):
    df_raw = pd.read_excel(file, header=None)
    mask = df_raw.apply(lambda r: r.astype(str).str.contains("Score", case=False).any(), axis=1)
    h_idx = mask.idxmax() if mask.any() else 0
    df = pd.read_excel(file, header=h_idx)
    df.columns = df.columns.str.strip()
    return process_data(df)

df = load_and_clean(st.session_state.file_ref)

score_col = "Average Score"
sup_col = next((c for c in ["Immediate Supervisor", "Supervisor", "Manager"] if c in df.columns), df.columns[0])

# ==================================================
# 4. PAGE LOGIC
# ==================================================

if st.session_state.current_page == "OVERVIEW":
    k_cols = st.columns(5)
    brackets = [
        ("Outstanding", len(df[df[score_col] >= 90]), "100-90"),
        ("Excellent", len(df[(df[score_col] >= 80) & (df[score_col] < 90)]), "89-80"),
        ("Good", len(df[(df[score_col] >= 60) & (df[score_col] < 80)]), "79-60"),
        ("Satisfactory", len(df[(df[score_col] >= 50) & (df[score_col] < 60)]), "59-50"),
        ("Poor", len(df[df[score_col] < 50]), "49-0")
    ]
    for i, (label, val, rng) in enumerate(brackets):
        with k_cols[i]:
            st.markdown(f'''<div class="kpi-box"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-range">{rng}</div></div>''', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    g1, g2, g3 = st.columns([1.2, 1.2, 0.8])
    with g1:
        st.markdown("<p style='color:white; font-weight:bold;'>Departmental Benchmarks</p>", unsafe_allow_html=True)
        d_avg = df.groupby("Department")[score_col].mean().sort_values().reset_index()
        fig = px.bar(d_avg, x=score_col, y="Department", orientation='h', color_discrete_sequence=['#FF8C00'])
        fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white", height=350)
        st.plotly_chart(fig, use_container_width=True)
    with g2:
        st.markdown("<p style='color:white; font-weight:bold;'>Elite Performers (Top 10)</p>", unsafe_allow_html=True)
        top = df.sort_values(score_col, ascending=False).head(10)
        fig = px.bar(top, x="Emp Name", y=score_col, color_discrete_sequence=['#FFA500'])
        fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white", height=350)
        st.plotly_chart(fig, use_container_width=True)
    with g3:
        st.markdown("<p style='color:white; font-weight:bold;'>KPI Document Completion</p>", unsafe_allow_html=True)
        avg_val = df[score_col].mean()
        fig = px.pie(values=[avg_val, 100-avg_val], hole=0.75, color_discrete_sequence=['#FF8C00', 'rgba(255,255,255,0.2)'])
        fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', font_color="white", showlegend=False, height=350)
        fig.add_annotation(text=f"{round(avg_val,1)}%", x=0.5, y=0.5, font_size=24, font_color="white", showarrow=False)
        st.plotly_chart(fig, use_container_width=True)

elif st.session_state.current_page == "DEPT":
    st.markdown("<h3 style='color: white;'>Department Analytics Portal</h3>", unsafe_allow_html=True)
    search_d = st.text_input("Search Department...", "").lower()
    depts = sorted([d for d in df["Department"].dropna().unique() if search_d in d.lower()])
    cols = st.columns(4)
    for idx, dept in enumerate(depts):
        d_df = df[df["Department"] == dept]
        dept_avg_val = round(d_df[score_col].mean(), 1)
        with cols[idx % 4]:
            st.markdown(f'''<div class="portal-card"><div class="circle-icon" style="background: #FF8C00;">{dept[0].upper()}</div><div style="font-size:14px; font-weight:800; color:#001f4d;">{dept}</div><div style="font-size:12px; color:#666; margin-bottom:15px;">Staff: {len(d_df)} | Avg: {dept_avg_val}%</div></div>''', unsafe_allow_html=True)
            if st.button("View Full Metrics", key=f"d_{dept}"):
                st.dataframe(d_df.sort_values(score_col, ascending=False), use_container_width=True)

elif st.session_state.current_page == "SUP":
    st.markdown("<h3 style='color: white;'>Supervisor Leadership Portal</h3>", unsafe_allow_html=True)
    search_s = st.text_input("Find Supervisor...", "").lower()
    sups = sorted([s for s in df[sup_col].dropna().unique() if search_s in str(s).lower()])
    
    cols = st.columns(4)
    for idx, sup in enumerate(sups):
        s_df = df[df[sup_col] == sup]
        initials = "".join([n[0] for n in str(sup).split()[:2]]).upper()
        
        comp_df = s_df[s_df[score_col] > 0].sort_values(score_col, ascending=False)
        pend_df = s_df[s_df[score_col] == 0]

        with cols[idx % 4]:
            st.markdown(f'''
                <div class="portal-card">
                    <div class="circle-icon" style="background: #001f4d; border: 2px solid #FF8C00;">{initials}</div>
                    <div style="font-size:14px; font-weight:800; color:#001f4d;">{sup}</div>
                    <div style="font-size:11px; color:#666; margin-bottom:15px;">
                        Done: {len(comp_df)} | Pending: {len(pend_df)}
                    </div>
                </div>
            ''', unsafe_allow_html=True)
            
            if st.button("Review Team", key=f"s_{idx}"):
                # Pending Section
                st.markdown(f'<div class="status-header-pending">Pending Reviews for {sup} ({len(pend_df)})</div>', unsafe_allow_html=True)
                if not pend_df.empty:
                    st.dataframe(pend_df, use_container_width=True)
                    # ADDED: CSV Download for Pending List
                    csv_data = pend_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="Download Pending List",
                        data=csv_data,
                        file_name=f"Pending_Reviews_{sup.replace(' ', '_')}.csv",
                        mime="text/csv",
                        key=f"dl_{idx}"
                    )
                else:
                    st.success("All reviews are completed for this team.")

                # Completed Section
                st.markdown(f'<div class="status-header-complete">Completed Reviews for {sup} ({len(comp_df)})</div>', unsafe_allow_html=True)
                if not comp_df.empty:
                    st.dataframe(comp_df, use_container_width=True)
                else:
                    st.info("No completed reviews yet.")
