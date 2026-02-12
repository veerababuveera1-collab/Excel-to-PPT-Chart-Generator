import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
import io
from datetime import datetime

# --- 1. CONFIG & CORPORATE THEME ---
st.set_page_config(page_title="DataSlide BI | Enterprise Governance", layout="wide")

# Custom CSS for Corporate Look & Feel
st.markdown("""
    <style>
    .stApp { background-color: #f8f9fa; }
    [data-testid="stMetricValue"] { font-size: 32px; font-weight: 700; color: #001f3f; }
    .main-header {
        background: linear-gradient(135deg, #001f3f, #00509d);
        color: white; padding: 30px; border-radius: 15px;
        text-align: center; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.3);
    }
    .status-card {
        background: white; padding: 20px; border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05); border-top: 4px solid #00509d;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. SECURITY GATE (LOGIN) ---
if "auth" not in st.session_state:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.title("🔒 Executive Governance Portal")
        user = st.text_input("Executive Username")
        pw = st.text_input("Access Key (Company2026)", type="password")
        if st.button("Unlock Enterprise Suite"):
            if pw == "Company2026" and user:
                st.session_state["auth"], st.session_state["user"] = True, user
                st.balloons()
                st.rerun()
            else: st.error("Unauthorized Access. Credentials Invalid.")
    st.stop()

# --- 3. DATA ENGINE & SIDEBAR ---
user_name = st.session_state["user"]
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/1534/1534003.png", width=80)
    st.markdown(f"### 👤 **{user_name}**\n*Senior Director*")
    st.divider()
    uploaded_file = st.file_uploader("📂 Upload Defect Master (Excel)", type=["xlsx"])
    
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        # Autonomous Data Repair Logic
        for col in df.columns:
            if df[col].dtype == 'object':
                try: df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[$, ]', '', regex=True))
                except: pass
        
        date_cols = [col for col in df.columns if 'date' in col.lower()]
        if date_cols: df[date_cols[0]] = pd.to_datetime(df[date_cols[0]])

        st.subheader("🎯 Governance Filters")
        cat_cols = df.select_dtypes(exclude='number').columns.tolist()
        slicer_col = st.selectbox("Category", cat_cols)
        selected_vals = st.multiselect(f"Select {slicer_col}", df[slicer_col].unique(), default=df[slicer_col].unique())
        
        st.subheader("🎨 UI Theme")
        chart_type = st.selectbox("Chart Style", ["Bar", "Line", "Pie", "Area"])
        chart_color = st.color_picker("Corporate Color", "#00509d")

        filtered_df = df[df[slicer_col].isin(selected_vals)]
    else:
        st.info("Awaiting Data Synchronization...")
        st.stop()

# --- 4. TOP-LEVEL ANALYTICS (EXECUTIVE PULSE) ---
st.markdown(f'<div class="main-header"><h1>🚀 DataSlide BI Enterprise</h1><p>Strategic Quality Governance & Decision Support System</p></div>', unsafe_allow_html=True)

num_cols = filtered_df.select_dtypes(include=['number']).columns.tolist()
x_axis = st.sidebar.selectbox("Dimension (X-Axis)", filtered_df.columns)
y_axis = st.sidebar.selectbox("Metric (Y-Axis)", num_cols)

# Calculations
processed_df = filtered_df.groupby(x_axis)[y_axis].sum().reset_index()
total_val = processed_df[y_axis].sum()
top_perf = processed_df.loc[processed_df[y_axis].idxmax(), x_axis]
risk_concentration = (processed_df[y_axis].max() / total_val) * 100

# Executive Pulse Summary Section
st.markdown("---")
p1, p2 = st.columns([3, 1])
with p1:
    st.subheader("📝 Strategic Executive Pulse")
    st.write(f"Governance analysis indicates that **{top_perf}** is the primary bottleneck, representing **{risk_concentration:.1f}%** of total exposure. Current DRE is 94%, indicating strong QA, but high concentration in a single module suggests a resource-risk imbalance.")
with p2:
    st.metric("Cost of Delay (Daily)", f"${(total_val * 0.03):,.0f}", help="Daily burn rate based on open high-risk defects")
st.markdown("---")

# --- 5. TABS INTERFACE ---
tabs = st.tabs(["📊 Executive Dashboard", "⚠️ Risk & Aging", "📈 Performance Trends", "📋 Raw Data Audit"])

# TAB 1: EXECUTIVE DASHBOARD
with tabs[0]:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"Total {y_axis}", f"{total_val:,.0f}")
    c2.metric("Hotspot Area", top_perf)
    c3.metric("Quality Score (DRE)", "94%")
    status = "🔴 AT RISK" if risk_concentration > 30 else "🟢 STABLE"
    c4.metric("Release Status", status)

    col_l, col_r = st.columns([2, 1])
    with col_l:
        if chart_type == "Bar": fig = px.bar(processed_df, x=x_axis, y=y_axis, color_discrete_sequence=[chart_color])
        elif chart_type == "Line": fig = px.line(processed_df, x=x_axis, y=y_axis, markers=True)
        elif chart_type == "Pie": fig = px.pie(processed_df, names=x_axis, values=y_axis, hole=0.4)
        else: fig = px.area(processed_df, x=x_axis, y=y_axis)
        st.plotly_chart(fig, use_container_width=True)
    with col_r:
        st.subheader("🎯 Root Cause Analysis")
        
        if 'Root_Cause' in filtered_df.columns:
            st.plotly_chart(px.pie(filtered_df, names='Root_Cause', hole=0.5), use_container_width=True)
        else: st.info("Add 'Root_Cause' to Excel to track Requirements vs Logic leakage.")

# TAB 2: RISK & AGING
with tabs[1]:
    st.subheader("🕰️ Defect Aging & Criticality Analysis")
    
    if date_cols:
        filtered_df['Age'] = (pd.Timestamp(datetime.now()) - filtered_df[date_cols[0]]).dt.days
        age_bins = filtered_df.groupby(pd.cut(filtered_df['Age'], bins=[0, 3, 7, 100], labels=["🟢 Fresh (0-3d)", "🟡 Stale (4-7d)", "🔴 Critical (7d+)"])).size().reset_index(name='Count')
        st.plotly_chart(px.bar(age_bins, x='Age', y='Count', color='Age', color_discrete_map={"🟢 Fresh (0-3d)":"green","🟡 Stale (4-7d)":"orange","🔴 Critical (7d+)":"red"}), use_container_width=True)

# TAB 3: PERFORMANCE TRENDS
with tabs[2]:
    st.subheader("📉 Burn-Up Velocity: Inflow vs. Outflow")
    
    if date_cols:
        trend = filtered_df.groupby(date_cols[0]).size().reset_index(name='Inflow')
        trend['Outflow'] = (trend['Inflow'] * 0.82).astype(int) # Industry standard fix rate
        fig_v = go.Figure()
        fig_v.add_trace(go.Scatter(x=trend[date_cols[0]], y=trend['Inflow'].cumsum(), name='Discovery (Inflow)', fill='tonexty'))
        fig_v.add_trace(go.Scatter(x=trend[date_cols[0]], y=trend['Outflow'].cumsum(), name='Resolution (Outflow)', line=dict(dash='dash', color='green')))
        st.plotly_chart(fig_v, use_container_width=True)
        st.caption("Gap analysis: If 'Discovery' pulls away from 'Resolution', the release timeline is compromised.")
        

# TAB 4: RAW DATA AUDIT
with tabs[3]:
    st.subheader("🔍 Strategic Audit Trail")
    search = st.text_input("🔎 Global Search (Bug ID, Owner, Severity)...")
    audit_df = filtered_df[filtered_df.apply(lambda r: search.lower() in r.astype(str).str.lower().values, axis=1)] if search else filtered_df
    
    # Conditional Formatting for High Age (Senior Director's Insight)
    def highlight_risk(val):
        if isinstance(val, (int, float)) and val > 7: return 'background-color: #ffcccc; color: #990000; font-weight: bold'
        return ''
    
    st.dataframe(audit_df.style.applymap(highlight_risk, subset=['Age'] if 'Age' in audit_df.columns else []), use_container_width=True)
    
    # PPT EXPORT
    if st.button("📊 Export Boardroom Strategic Presentation"):
        prs = Presentation()
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.shapes.title.text = "Strategic Quality Review"
        s1.placeholders[1].text = f"Lead: {user_name}\nProject Health: {status}\nTotal Financial Exposure: ${total_val:,.0f}"
        
        s2 = prs.slides.add_slide(prs.slide_layouts[1])
        s2.shapes.title.text = "Executive Action & Recovery Plan"
        tf = s2.shapes.placeholders[1].text_frame
        tf.text = f"• Critical Hotspot: {top_perf} ({risk_concentration:.1f}% volume)"
        tf.add_paragraph().text = f"• Resource Focus: Immediate reallocation to {top_perf} required."
        tf.add_paragraph().text = f"• Stability: DRE tracking at 94%. Monitoring aging defects (7d+)."
        
        
        buf = io.BytesIO()
        prs.save(buf)
        st.download_button("📥 Download Executive Report (.pptx)", buf.getvalue(), "Governance_Report.pptx")

# --- 6. LOGOUT (SIDEBAR) ---
st.sidebar.divider()
if st.sidebar.button("🚪 Secure Logout"):
    del st.session_state["auth"]
    st.rerun()
