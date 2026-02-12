import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime

# --- 1. SETTINGS & ELITE STYLING ---
st.set_page_config(page_title="Enterprise Governance | DataSlide BI", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f8f9fa; }
    .main-header {
        background: linear-gradient(135deg, #001f3f, #004080);
        color: white; padding: 2rem; border-radius: 15px;
        text-align: center; margin-bottom: 2rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
    }
    .metric-card {
        background-color: white; padding: 1.5rem; border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1); border-left: 5px solid #004080;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. THE NETWORKING DAYS ENGINE (With Holidays) ---
def get_biz_days(start_date, end_date):
    # Dummy Corporate Holidays 2026 - You can add more here
    holidays_2026 = ['2026-01-01', '2026-01-26', '2026-08-15', '2026-10-02', '2026-12-25']
    try:
        # np.busday_count counts business days between two dates
        start = start_date.astype('datetime64[D]')
        # If not closed, use current date
        if pd.isnull(end_date):
            end = np.datetime64('today')
        else:
            end = end_date.astype('datetime64[D]')
        
        count = np.busday_count(start, end, holidays=holidays_2026)
        return max(0, count)
    except:
        return 0

# --- 3. SIDEBAR: DATA SYNCHRONIZATION ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/1162/1162456.png", width=80)
    st.title("Governance Portal")
    uploaded_file = st.file_uploader("📂 Synchronize Defect Master", type=["xlsx"])
    
    if not uploaded_file:
        st.warning("Awaiting Executive Data Upload...")
        st.stop()

    # Data Processing
    df = pd.read_excel(uploaded_file)
    
    # Standardizing Columns (Case Insensitive)
    df.columns = [c.strip() for c in df.columns]
    df['Discovery_Date'] = pd.to_datetime(df['Discovery_Date'])
    df['Closed_Date'] = pd.to_datetime(df['Closed_Date'])
    df['Week'] = df['Discovery_Date'].dt.strftime('Wk-%U')
    
    # Environment Switcher (The Core Upgrade)
    st.divider()
    env_list = df['Environment'].unique().tolist()
    selected_env = st.selectbox("🎯 Select Environment View", env_list)
    
    # Filtering Data
    df_filtered = df[df['Environment'] == selected_env].copy()
    
    # Calculate Aging using Networking Days logic
    df_filtered['Aging_Days'] = df_filtered.apply(lambda x: get_biz_days(x['Discovery_Date'], x['Closed_Date']), axis=1)

# --- 4. EXECUTIVE HEADER ---
st.markdown(f"""
    <div class="main-header">
        <h1>🚀 {selected_env} Defect Governance Dashboard</h1>
        <p>Real-time Release Stability & Aging Analysis | Networking Days Enabled</p>
    </div>
    """, unsafe_allow_html=True)

# --- 5. TOP LEVEL METRICS ---
m1, m2, m3, m4 = st.columns(4)
with m1:
    st.markdown(f'<div class="metric-card"><h3>Total Defects</h3><h2>{len(df_filtered)}</h2></div>', unsafe_allow_html=True)
with m2:
    open_count = len(df_filtered[df_filtered['Status'] != 'Closed'])
    st.markdown(f'<div class="metric-card"><h3>Active Backlog</h3><h2>{open_count}</h2></div>', unsafe_allow_html=True)
with m3:
    criticals = len(df_filtered[df_filtered['Severity'] == 'Critical'])
    st.markdown(f'<div class="metric-card"><h3>Critical Risks</h3><h2 style="color:red;">{criticals}</h2></div>', unsafe_allow_html=True)
with m4:
    avg_aging = round(df_filtered['Aging_Days'].mean(), 1)
    st.markdown(f'<div class="metric-card"><h3>Avg Biz Aging</h3><h2>{avg_aging} Days</h2></div>', unsafe_allow_html=True)

# --- 6. THE TABBED INTERFACE (Old + New) ---
t_old, t_new, t_audit = st.tabs(["📊 Executive Summary (Old Look)", "📈 Lifecycle Trend (New Look)", "📋 Detailed Audit Log"])

with t_old:
    # PARETO ANALYSIS (Original Functionality)
    st.subheader("Defect Distribution (Pareto)")
    pareto_data = df_filtered.groupby('App_Area').size().sort_values(ascending=False).reset_index(name='count')
    pareto_data['cumulative'] = 100 * (pareto_data['count'].cumsum() / pareto_data['count'].sum())
    
    fig_p = go.Figure()
    fig_p.add_trace(go.Bar(x=pareto_data['App_Area'], y=pareto_data['count'], name="Volume", marker_color='#004080'))
    fig_p.add_trace(go.Scatter(x=pareto_data['App_Area'], y=pareto_data['cumulative'], name="% Cumulative", yaxis="y2", line=dict(color="orange", width=3)))
    fig_p.update_layout(yaxis2=dict(overlaying="y", side="right", range=[0, 110]), title="Pareto Analysis by App Area")
    st.plotly_chart(fig_p, use_container_width=True)

    # TREEMAP & COST (Original Functionality)
    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(px.treemap(df_filtered, path=['Severity', 'App_Area'], values='Fix_Cost', title="Risk Weight by Cost"), use_container_width=True)
    with c2:
        st.plotly_chart(px.pie(df_filtered, names='Root_Cause', title="Root Cause Analysis", hole=0.4), use_container_width=True)

with t_new:
    # THE RED LINE LIFECYCLE (The Upgrade from your Photo)
    st.subheader(f"Weekly Lifecycle Velocity: {selected_env}")
    
    # Pivot for Matrix
    pivot = df_filtered.groupby(['Week', 'Status']).size().unstack(fill_value=0)
    
    # Open Backlog Logic (The Red Line)
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
    
    fig_l.add_trace(go.Scatter(name='Total Open Backlog', x=pivot.index, y=pivot['Open_Backlog'], 
                               line=dict(color='red', width=4), mode='lines+markers+text',
                               text=pivot['Open_Backlog'], textposition="top center"))
    
    fig_l.update_layout(barmode='group', title="Inflow vs Outflow vs Backlog")
    st.plotly_chart(fig_l, use_container_width=True)
    
    st.markdown("### Weekly Performance Matrix")
    st.dataframe(pivot.T.style.background_gradient(axis=1, cmap='YlGnBu'), use_container_width=True)

with t_audit:
    st.markdown("### 🔍 Full Traceability Matrix")
    # Using all your fields here
    cols_to_show = ['Defect_ID', 'Status', 'Severity', 'Assignee', 'Reporter', 'Aging_Days', 'Remarks', 'KPI_Status']
    st.dataframe(df_filtered[cols_to_show].style.applymap(
        lambda x: 'color: red; font-weight: bold' if x == 'Breached' else '', subset=['KPI_Status']
    ), use_container_width=True)
