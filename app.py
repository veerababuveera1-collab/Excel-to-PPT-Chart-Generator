import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

# --- 1. SETTINGS & EXECUTIVE THEME ---
st.set_page_config(page_title="Governance Command Center", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f4f7f9; }
    .main-header {
        background: linear-gradient(135deg, #001529, #003366);
        color: white; padding: 2rem; border-radius: 15px;
        text-align: center; margin-bottom: 2rem; box-shadow: 0 4px 12px rgba(0,0,0,0.2);
    }
    .metric-box {
        background-color: white; padding: 1.5rem; border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-top: 5px solid #003366;
        text-align: center;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. ENGINE: FAST NETWORKING DAYS ---
def get_biz_aging(df):
    df['Discovery_Date'] = pd.to_datetime(df['Discovery_Date'], errors='coerce')
    df['Closed_Date'] = pd.to_datetime(df['Closed_Date'], errors='coerce')
    
    start_dates = df['Discovery_Date'].dt.date.values.astype('datetime64[D]')
    # If not closed, calculate aging up to today (Feb 13, 2026)
    today = np.datetime64('2026-02-13') 
    end_dates = df['Closed_Date'].dt.date.fillna(pd.Timestamp('2026-02-13')).values.astype('datetime64[D]')
    
    # Standard 2026 Corporate Holidays
    hols = ['2026-01-01', '2026-01-26', '2026-08-15', '2026-10-02', '2026-12-25']
    
    try:
        return np.busday_count(start_dates, end_dates, holidays=hols)
    except:
        return (pd.to_datetime(end_dates) - pd.to_datetime(start_dates)).days

# --- 3. DATA ORCHESTRATION ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3281/3281306.png", width=80)
    st.title("Strategic Governance")
    uploaded_file = st.file_uploader("Upload Unified Defect Master", type=["xlsx"])
    
    if not uploaded_file:
        st.info("Awaiting Master Data Sync...")
        st.stop()

    raw_df = pd.read_excel(uploaded_file)
    raw_df.columns = [c.strip() for c in raw_df.columns]
    
    # Environment Selector (The Filter)
    env_choice = st.selectbox("🎯 Select View", raw_df['Environment'].unique())
    df = raw_df[raw_df['Environment'] == env_choice].copy()
    
    # Enrich Data
    df['Aging_Days'] = get_biz_aging(df)
    df['Week'] = df['Discovery_Date'].dt.strftime('Wk-%U')
    # SLA Classification
    df['SLA_Status'] = pd.cut(df['Aging_Days'], bins=[-1, 3, 7, 999], labels=['Green', 'Amber', 'Red'])

# --- 4. COMMAND CENTER INTERFACE ---
st.markdown(f'<div class="main-header"><h1>🚀 {env_choice} Command Center</h1><p>End-to-End Defect Lifecycle & Accountability</p></div>', unsafe_allow_html=True)

# Top Metrics Row
c1, c2, c3, c4 = st.columns(4)
with c1: st.markdown(f'<div class="metric-box"><h4>Total Volume</h4><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
with c2: st.markdown(f'<div class="metric-box"><h4>Active Backlog</h4><h2 style="color:#e67e22;">{len(df[df["Status"]!="Closed"])}</h2></div>', unsafe_allow_html=True)
with c3: st.markdown(f'<div class="metric-box"><h4>Critical Risk</h4><h2 style="color:#c0392b;">{len(df[df["Severity"]=="Critical"])}</h2></div>', unsafe_allow_html=True)
with c4: st.markdown(f'<div class="metric-box"><h4>Avg Aging</h4><h2>{round(df["Aging_Days"].mean(),1)} Days</h2></div>', unsafe_allow_html=True)

# --- 5. TABS: THE 360° VIEW ---
t_account, t_trend, t_old, t_audit = st.tabs([
    "👥 Accountability & SLA", 
    "📈 Lifecycle Trend", 
    "📊 Pareto & Cost (Old Look)", 
    "📋 Detailed Audit Log"
])

with t_account:
    st.subheader("Ownership Leaderboard")
    col_a, col_b = st.columns([2, 1])
    with col_a:
        leader_df = df.groupby(['Assignee', 'Severity']).size().reset_index(name='Count')
        fig_l = px.bar(leader_df, y='Assignee', x='Count', color='Severity', orientation='h', 
                       title="Defects by Assignee", barmode='stack',
                       color_discrete_map={'Critical': '#c0392b', 'High': '#e67e22', 'Medium': '#f1c40f', 'Low': '#2ecc71'})
        st.plotly_chart(fig_l, use_container_width=True)
    with col_b:
        sla_counts = df['SLA_Status'].value_counts()
        fig_s = px.pie(names=sla_counts.index, values=sla_counts.values, hole=0.5,
                       color=sla_counts.index, title="Aging SLA Health",
                       color_discrete_map={'Green': '#2ecc71', 'Amber': '#f1c40f', 'Red': '#c0392b'})
        st.plotly_chart(fig_s, use_container_width=True)

with t_trend:
    st.subheader("The 'Red Line' Backlog Analysis")
    pivot = df.groupby(['Week', 'Status']).size().unstack(fill_value=0)
    pivot['Backlog'] = 0
    running = 0
    for idx in pivot.index:
        running = (running + pivot.loc[idx].get('Created', 0)) - (pivot.loc[idx].get('Closed', 0) + pivot.loc[idx].get('Moved', 0))
        pivot.loc[idx, 'Backlog'] = running

    fig_t = go.Figure()
    if 'Created' in pivot.columns: fig_t.add_trace(go.Bar(name='Created', x=pivot.index, y=pivot['Created'], marker_color='#3498db'))
    if 'Closed' in pivot.columns: fig_t.add_trace(go.Bar(name='Closed', x=pivot.index, y=pivot['Closed'], marker_color='#2ecc71'))
    fig_t.add_trace(go.Scatter(name='Active Backlog', x=pivot.index, y=pivot['Backlog'], 
                               line=dict(color='red', width=4), mode='lines+markers+text', 
                               text=pivot['Backlog'], textposition="top center"))
    st.plotly_chart(fig_t, use_container_width=True)

with t_old:
    st.subheader("Executive Pareto & Cost Analysis")
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        # Re-creating your exact Pareto logic
        pareto = df.groupby('App_Area').size().sort_values(ascending=False).reset_index(name='count')
        st.plotly_chart(px.bar(pareto, x='App_Area', y='count', title="Defect Density by Area"), use_container_width=True)
    with col_p2:
        # Your original Treemap logic
        st.plotly_chart(px.treemap(df, path=['Severity', 'App_Area'], values='Fix_Cost', 
                                   title="Financial Risk Exposure"), use_container_width=True)

with t_audit:
    st.subheader("End-to-End Traceability")
    # Clean table with all old fields
    st.dataframe(df[['Defect_ID', 'Status', 'Severity', 'Assignee', 'Reporter', 'Aging_Days', 'Remarks', 'Root_Cause']], use_container_width=True)
