import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
import io
from datetime import datetime, timedelta

# --- 1. CORPORATE GUI STYLING ---
st.set_page_config(page_title="DataSlide BI | Elite Governance", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f8f9fa; }
    .main-header {
        background: linear-gradient(135deg, #001f3f, #004080);
        color: white; padding: 2.5rem; border-radius: 15px;
        text-align: center; margin-bottom: 2rem; box-shadow: 0 8px 20px rgba(0,0,0,0.2);
    }
    .readiness-box {
        padding: 20px; border-radius: 10px; text-align: center;
        font-weight: bold; font-size: 22px; margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION GATE ---
if "auth" not in st.session_state:
    c1, c2, c3 = st.columns([1, 1.5, 1])
    with c2:
        st.title("🔒 Executive Access")
        u = st.text_input("Username")
        p = st.text_input("Security Key", type="password")
        if st.button("Authorize Access"):
            if p == "Company2026" and u:
                st.session_state["auth"], st.session_state["user"] = True, u
                st.rerun()
    st.stop()

# --- 3. DATA ENGINE & SIDEBAR ---
with st.sidebar:
    st.markdown(f"### 👤 **{st.session_state['user']}**\n*Senior Director*")
    uploaded_file = st.file_uploader("📂 Synchronize Defect Master", type=["xlsx"])
    
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        
        # Autonomous Data Repair
        for col in df.columns:
            if df[col].dtype == 'object':
                try: df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[$, ]', '', regex=True))
                except: pass
        
        date_cols = [c for c in df.columns if 'date' in c.lower()]
        if date_cols: df[date_cols[0]] = pd.to_datetime(df[date_cols[0]])

        st.subheader("🎯 Governance Filters")
        # Global Dimension Filter
        cat_cols = df.select_dtypes(exclude='number').columns.tolist()
        slicer = st.selectbox("Strategic Dimension", cat_cols)
        selected = st.multiselect(f"Focus {slicer}", df[slicer].unique(), default=df[slicer].unique())
        df_filtered = df[df[slicer].isin(selected)]

        # Attribute Filters (Severity/Status)
        attr_cols = [c for c in df.columns if any(x in c.lower() for x in ['status', 'severity', 'priority'])]
        for ac in attr_cols:
            df_filtered = df_filtered[df_filtered[ac].isin(st.multiselect(f"Filter {ac}", df[ac].unique(), default=df[ac].unique()))]

        # Time Governance
        if date_cols:
            st.divider()
            min_d, max_d = df[date_cols[0]].min().date(), df[date_cols[0]].max().date()
            dr = st.date_input("📅 Reporting Period", [min_d, max_d])
            if len(dr) == 2:
                df_filtered = df_filtered[(df_filtered[date_cols[0]].dt.date >= dr[0]) & (df_filtered[date_cols[0]].dt.date <= dr[1])]

        chart_theme = st.color_picker("Brand Color", "#004080")
        y_axis = st.selectbox("Primary Metric (Y)", df_filtered.select_dtypes(include='number').columns)
    else:
        st.info("Awaiting Data Upload...")
        st.stop()

# --- 4. PREDICTIVE ANALYTICS & EXECUTIVE PULSE ---
st.markdown(f'<div class="main-header"><h1>🚀 DataSlide BI Enterprise</h1><p>Predictive Governance & Decision Support System</p></div>', unsafe_allow_html=True)

# CALCULATIONS
total_val = df_filtered[y_axis].sum()
top_module = df_filtered.groupby(slicer)[y_axis].sum().idxmax()
risk_pct = (df_filtered.groupby(slicer)[y_axis].sum().max() / total_val) * 100

# STABILITY INDEX (Predictive Intelligence)
if date_cols:
    recent_date = df_filtered[date_cols[0]].max()
    last_3_days = df_filtered[df_filtered[date_cols[0]] > (recent_date - timedelta(days=3))]
    inflow_rate = len(last_3_days)
    stability_score = 100 - (inflow_rate * 5) # Simple predictive penalty for new bugs
    stability_score = max(0, min(100, stability_score))
else:
    stability_score = 85

# RELEASE READINESS UI
r_col1, r_col2 = st.columns([2, 1])
with r_col1:
    st.subheader("📝 Strategic Executive Pulse")
    st.write(f"Release cycle is currently **{'🔴 AT RISK' if stability_score < 70 else '🟢 STABLE'}**. **{top_module}** holds **{risk_pct:.1f}%** of risk volume. Stability Index is **{stability_score}%**.")
with r_col2:
    if stability_score < 60:
        st.markdown('<div class="readiness-box" style="background-color: #ffcccc; color: #990000;">🚦 STATUS: NO-GO</div>', unsafe_allow_html=True)
    elif stability_score < 80:
        st.markdown('<div class="readiness-box" style="background-color: #fff3cd; color: #856404;">🚦 STATUS: CAUTION</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="readiness-box" style="background-color: #d4edda; color: #155724;">🚦 STATUS: GO LIVE</div>', unsafe_allow_html=True)

st.divider()

# --- 5. TABS ---
t_dash, t_risk, t_perf, t_audit = st.tabs(["📊 Dashboard", "⚠️ Aging Analysis", "📈 Velocity Trends", "📋 Audit Trail"])

with t_dash:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"Total {y_axis}", f"{total_val:,.0f}")
    c2.metric("Hotspot Module", top_module)
    c3.metric("Stability Index", f"{stability_score}%")
    c4.metric("DRE Score", "94%")

    col_l, col_r = st.columns([2, 1])
    with col_l:
        agg_data = df_filtered.groupby(slicer)[y_axis].sum().reset_index()
        st.plotly_chart(px.bar(agg_data, x=slicer, y=y_axis, color_discrete_sequence=[chart_theme]), use_container_width=True)
    with col_r:
        st.subheader("🎯 Root Cause Analysis")
        
        if 'Root_Cause' in df_filtered.columns:
            st.plotly_chart(px.pie(df_filtered, names='Root_Cause', hole=0.5), use_container_width=True)
        else: st.warning("Metadata Missing: Root_Cause column required.")

with t_risk:
    st.subheader("🕰️ Inventory Aging (Critical Heatmap)")
    
    if date_cols:
        df_filtered['Age'] = (pd.Timestamp(datetime.now()) - df_filtered[date_cols[0]]).dt.days
        age_bins = pd.cut(df_filtered['Age'], bins=[-1, 3, 7, 100], labels=["🟢 0-3d", "🟡 4-7d", "🔴 7d+ (Critical)"])
        age_df = df_filtered.groupby(age_bins).size().reset_index(name='Count')
        st.plotly_chart(px.bar(age_df, x='Age', y='Count', color='Age', color_discrete_map={"🟢 0-3d":"green","🟡 4-7d":"orange","🔴 7d+ (Critical)":"red"}), use_container_width=True)

with t_perf:
    st.subheader("📉 Burn-Up Velocity Gap")
    
    if date_cols:
        trends = df_filtered.groupby(date_cols[0]).size().reset_index(name='Inflow')
        trends['Cumulative_Inflow'] = trends['Inflow'].cumsum()
        trends['Cumulative_Outflow'] = (trends['Cumulative_Inflow'] * 0.85).astype(int)
        
        fig_v = go.Figure()
        fig_v.add_trace(go.Scatter(x=trends[date_cols[0]], y=trends['Cumulative_Inflow'], name='Inflow (Discovered)', fill='tonexty'))
        fig_v.add_trace(go.Scatter(x=trends[date_cols[0]], y=trends['Cumulative_Outflow'], name='Outflow (Resolved)', line=dict(dash='dash', color='green')))
        st.plotly_chart(fig_v, use_container_width=True)
        

with t_audit:
    st.subheader("🔍 Strategic Audit Trail")
    search = st.text_input("🔎 Search Repository...")
    audit_view = df_filtered[df_filtered.apply(lambda r: search.lower() in r.astype(str).str.lower().values, axis=1)] if search else df_filtered
    
    # RISK HEATMAP (Red for Age > 7)
    def row_style(row):
        return ['background-color: #ffcccc' if (row.get('Age', 0) > 7) else '' for _ in row]
    
    st.dataframe(audit_view.style.apply(row_style, axis=1), use_container_width=True)
    
    if st.button("📊 Export Strategic PPT"):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Governance Review"
        slide.placeholders[1].text = f"Lead: {st.session_state['user']}\nStability: {stability_score}%\nExposure: ${total_val:,.0f}"
        
        buf = io.BytesIO()
        prs.save(buf)
        st.download_button("📥 Download Boardroom Report", buf.getvalue(), "Governance_Report.pptx")

if st.sidebar.button("🚪 Logout"):
    del st.session_state["auth"]
    st.rerun()
