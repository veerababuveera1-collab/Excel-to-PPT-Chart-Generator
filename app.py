import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

# --- 1. SETTINGS & ELITE STYLING ---
st.set_page_config(page_title="Enterprise Governance | DataSlide BI", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f8f9fa; }
    .main-header {
        background: linear-gradient(135deg, #001f3f, #004080);
        color: white; padding: 1.5rem; border-radius: 15px;
        text-align: center; margin-bottom: 2rem;
    }
    .metric-card {
        background-color: white; padding: 1.2rem; border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1); border-left: 5px solid #004080;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FAST NETWORKING DAYS LOGIC ---
def calculate_aging(df):
    # Standardizing dates
    df['Discovery_Date'] = pd.to_datetime(df['Discovery_Date'], errors='coerce')
    df['Closed_Date'] = pd.to_datetime(df['Closed_Date'], errors='coerce')
    
    # Calculate Business Days (Networking Days)
    # Using np.busday_count for speed
    start_dates = df['Discovery_Date'].dt.date.values.astype('datetime64[D]')
    # If not closed, use today's date
    today = np.datetime64('today')
    end_dates = df['Closed_Date'].dt.date.fillna(pd.Timestamp.now()).values.astype('datetime64[D]')
    
    # Holidays 2026
    hols = ['2026-01-01', '2026-01-26', '2026-08-15']
    
    try:
        return np.busday_count(start_dates, end_dates, holidays=hols)
    except:
        return (pd.to_datetime(end_dates) - pd.to_datetime(start_dates)).days

# --- 3. SIDEBAR: DATA SYNC ---
with st.sidebar:
    st.title("📂 Governance Portal")
    uploaded_file = st.file_uploader("Synchronize Defect Master", type=["xlsx"])
    
    if not uploaded_file:
        st.info("Please upload your Master Excel to begin.")
        st.stop()

    df_raw = pd.read_excel(uploaded_file)
    df_raw.columns = [c.strip() for c in df_raw.columns]
    
    # Environment Selector
    env_list = df_raw['Environment'].unique().tolist()
    selected_env = st.selectbox("🎯 Select Environment View", env_list)
    
    # Filter Data
    df = df_raw[df_raw['Environment'] == selected_env].copy()
    df['Aging_Days'] = calculate_aging(df)
    df['Week'] = df['Discovery_Date'].dt.strftime('Wk-%U')

# --- 4. EXECUTIVE HEADER ---
st.markdown(f'<div class="main-header"><h1>🚀 {selected_env} Defect Governance</h1></div>', unsafe_allow_html=True)

# --- 5. TOP METRICS ---
m1, m2, m3, m4 = st.columns(4)
with m1:
    st.markdown(f'<div class="metric-card"><h3>Total Defects</h3><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
with m2:
    active = len(df[df['Status'] != 'Closed'])
    st.markdown(f'<div class="metric-card"><h3>Active Backlog</h3><h2>{active}</h2></div>', unsafe_allow_html=True)
with m3:
    critical = len(df[df['Severity'] == 'Critical'])
    st.markdown(f'<div class="metric-card"><h3>Critical Risks</h3><h2 style="color:red;">{critical}</h2></div>', unsafe_allow_html=True)
with m4:
    avg_age = round(df['Aging_Days'].mean(), 1)
    st.markdown(f'<div class="metric-card"><h3>Avg Biz Aging</h3><h2>{avg_age} Days</h2></div>', unsafe_allow_html=True)

# --- 6. TABS (OLD + NEW FUNCTIONALITIES) ---
t_old, t_new, t_audit = st.tabs(["📊 Executive Summary (Old Look)", "📈 Lifecycle Trend (New Look)", "📋 Detailed Audit Log"])

with t_old:
    # PARETO CHART (Old Functionality)
    st.subheader("Defect Distribution (Pareto)")
    pareto = df.groupby('App_Area').size().sort_values(ascending=False).reset_index(name='count')
    pareto['cum_pct'] = 100 * (pareto['count'].cumsum() / pareto['count'].sum())
    
    fig_p = go.Figure()
    fig_p.add_trace(go.Bar(x=pareto['App_Area'], y=pareto['count'], name="Volume", marker_color='#004080'))
    fig_p.add_trace(go.Scatter(x=pareto['App_Area'], y=pareto['cum_pct'], name="Cumulative %", yaxis="y2", line=dict(color="orange", width=3)))
    fig_p.update_layout(yaxis2=dict(overlaying="y", side="right", range=[0, 110]), showlegend=False)
    st.plotly_chart(fig_p, use_container_width=True)

    # TREEMAP (Old Functionality)
    st.plotly_chart(px.treemap(df, path=['Severity', 'App_Area'], values='Fix_Cost', title="Defect Weight by Cost"), use_container_width=True)

with t_new:
    # LIFECYCLE TREND (Upgraded functionality from your photo)
    st.subheader("Weekly Lifecycle Trend")
    pivot = df.groupby(['Week', 'Status']).size().unstack(fill_value=0)
    
    # Red Line Backlog Logic
    pivot['Open_Backlog'] = 0
    running = 0
    for idx in pivot.index:
        created = pivot.loc[idx, 'Created'] if 'Created' in pivot.columns else 0
        closed = pivot.loc[idx, 'Closed'] if 'Closed' in pivot.columns else 0
        moved = pivot.loc[idx, 'Moved'] if 'Moved' in pivot.columns else 0
        running = (running + created) - (closed + moved)
        pivot.loc[idx, 'Open_Backlog'] = running

    fig_l = go.Figure()
    colors = {'Created': '#3498db', 'Closed': '#2ecc71', 'Moved': '#f1c40f'}
    for s in ['Created', 'Closed', 'Moved']:
        if s in pivot.columns:
            fig_l.add_trace(go.Bar(name=s, x=pivot.index, y=pivot[s], marker_color=colors[s]))
    
    fig_l.add_trace(go.Scatter(name='Open Backlog (Red Line)', x=pivot.index, y=pivot['Open_Backlog'], 
                               line=dict(color='red', width=4), mode='lines+markers'))
    
    st.plotly_chart(fig_l, use_container_width=True)
    
    # Matrix Table (Fixed the error here by removing style.gradient)
    st.markdown("### Weekly Matrix Summary")
    st.dataframe(pivot.T, use_container_width=True)

with t_audit:
    st.subheader("Detailed Traceability Matrix")
    # Showing all your requested fields
    st.dataframe(df[['Defect_ID', 'Status', 'Severity', 'Assignee', 'Reporter', 'Root_Cause', 'Remarks', 'Aging_Days']], use_container_width=True)
