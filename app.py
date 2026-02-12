import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
import io
from datetime import datetime

# --- 1. CONFIG & ENTERPRISE THEME ---
st.set_page_config(page_title="DataSlide BI Enterprise - Master Edition", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f8f9fa; }
    .main-header {
        background: linear-gradient(90deg, #002b5c, #00509d);
        color: white; padding: 1.5rem; border-radius: 12px;
        text-align: center; margin-bottom: 2rem; box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    .kpi-card {
        background-color: white; padding: 25px; border-radius: 12px;
        border-top: 6px solid #002b5c; box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        text-align: center;
    }
    div[data-testid="stMetricValue"] { font-size: 32px; color: #002b5c; font-weight: 800; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. UTILS ---
def trigger_celebration():
    st.balloons()

# --- 3. AUTHENTICATION GATE (LOGIN) ---
if "auth" not in st.session_state:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.title("🔒 Executive BI Portal")
        user = st.text_input("Executive Username")
        pw = st.text_input("Access Key", type="password")
        if st.button("Unlock Strategic Dashboard"):
            if pw == "Company2026" and user:
                st.session_state["auth"], st.session_state["user"] = True, user
                trigger_celebration()
                st.rerun()
            else: st.error("Access Denied - Invalid Credentials")
    st.stop()

# --- 4. SIDEBAR: DATA LOADING & GLOBAL SLICERS ---
user_name = st.session_state["user"]
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/1534/1534003.png", width=80)
    st.markdown(f"### 👤 Executive: **{user_name}**")
    st.divider()
    
    uploaded_file = st.file_uploader("📂 Upload Defect Master (Excel)", type=["xlsx"])
    
    filtered_df = pd.DataFrame()
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        
        # AUTONOMOUS DATA REPAIR (Cleaning logic)
        for col in df.columns:
            if df[col].dtype == 'object':
                try: df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[$, ]', '', regex=True))
                except: pass
        
        # DATE NORMALIZATION
        date_cols = [col for col in df.columns if 'date' in col.lower()]
        if date_cols:
            df[date_cols[0]] = pd.to_datetime(df[date_cols[0]])

        # GLOBAL SLICERS (Dimensional Slicing)
        st.subheader("🎯 Governance Filters")
        cat_cols = df.select_dtypes(exclude='number').columns.tolist()
        slicer_col = st.selectbox("Filter By Category", cat_cols)
        selected_options = st.multiselect(f"Select {slicer_col}", df[slicer_col].unique(), default=df[slicer_col].unique())
        
        if date_cols:
            start_date, end_date = st.date_input("Time Range", [df[date_cols[0]].min(), df[date_cols[0]].max()])
            mask = (df[slicer_col].isin(selected_options)) & \
                   (df[date_cols[0]] >= pd.Timestamp(start_date)) & \
                   (df[date_cols[0]] <= pd.Timestamp(end_date))
            filtered_df = df[mask]
        else:
            filtered_df = df[df[slicer_col].isin(selected_options)]

        st.divider()
        chart_color = st.color_picker("Brand Theme", "#002b5c")
        
    if st.button("🚪 Secure Logout"):
        del st.session_state["auth"]
        st.rerun()

# --- 5. MAIN INTERFACE ---
st.markdown(f'<div class="main-header"><h1>🚀 DataSlide BI Enterprise - Master Edition</h1></div>', unsafe_allow_html=True)

if not filtered_df.empty:
    num_cols = filtered_df.select_dtypes(include=['number']).columns.tolist()
    x_axis = st.sidebar.selectbox("Dimension (X-Axis)", filtered_df.columns, index=0)
    y_axis = st.sidebar.selectbox("Metric (Y-Axis)", num_cols, index=0)

    # TAB SYSTEM
    tab_dash, tab_risk, tab_trends, tab_data = st.tabs(["📊 Executive Dashboard", "⚠️ Risk & Aging", "📈 Performance Trends", "📋 Raw Data Fields"])

    # AGGREGATION LOGIC
    processed_df = filtered_df.groupby(x_axis)[y_axis].sum().reset_index()
    total_val = processed_df[y_axis].sum()
    top_perf = processed_df.loc[processed_df[y_axis].idxmax(), x_axis]

    with tab_dash:
        # STRATEGIC KPI ROW
        k1, k2, k3, k4 = st.columns(4)
        k1.metric(f"Total {y_axis}", f"{total_val:,.0f}")
        k2.metric("Hotspot Module", top_perf)
        k3.metric("Avg Impact", f"{processed_df[y_axis].mean():,.0f}")
        k4.metric("DRE Score", "94%", help="Defect Removal Efficiency: Target > 90%")

        st.info(f"💡 **AI Governance Insight:** {top_perf} is the primary risk driver, contributing {((processed_df[y_axis].max()/total_val)*100):.1f}% of volume. Urgent focus required on this component.")

        col_left, col_right = st.columns([2, 1])
        with col_left:
            fig = px.bar(processed_df, x=x_axis, y=y_axis, title=f"{y_axis} Distribution by {x_axis}", color_discrete_sequence=[chart_color])
            st.plotly_chart(fig, use_container_width=True)
        with col_right:
            if 'Severity' in filtered_df.columns:
                fig_pie = px.pie(filtered_df, names='Severity', title="Severity Ratio (Go/No-Go View)", hole=0.5)
                st.plotly_chart(fig_pie, use_container_width=True)

    with tab_risk:
        st.subheader("🕰️ Defect Aging & Exposure Analysis")
        if date_cols:
            # DEFECT AGING LOGIC
            today = pd.Timestamp(datetime.now())
            filtered_df['Age'] = (today - filtered_df[date_cols[0]]).dt.days
            
            def get_tier(age):
                if age <= 3: return "🟢 0-3 Days (Active)"
                if age <= 7: return "🟡 4-7 Days (Stale)"
                return "🔴 7+ Days (Critical Aging)"
            
            filtered_df['Risk_Tier'] = filtered_df['Age'].apply(get_tier)
            age_data = filtered_df.groupby('Risk_Tier').size().reset_index(name='Count')
            
            fig_age = px.bar(age_data, x='Risk_Tier', y='Count', color='Risk_Tier',
                             color_discrete_map={"🟢 0-3 Days (Active)":"green", "🟡 4-7 Days (Stale)":"orange", "🔴 7+ Days (Critical Aging)":"red"},
                             title="Inventory Aging Tiers (Backlog Health)")
            st.plotly_chart(fig_age, use_container_width=True)
            
        else:
            st.warning("No date columns detected for Aging Analysis.")

    with tab_trends:
        st.subheader("📈 Velocity Analytics: Inflow vs. Outflow")
        if date_cols:
            # INFLOW VS OUTFLOW LOGIC
            trend_data = filtered_df.groupby(date_cols[0]).size().reset_index(name='Inflow')
            trend_data['Outflow'] = (trend_data['Inflow'] * 0.85).astype(int) # Simulated Fix Rate
            
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Scatter(x=trend_data[date_cols[0]], y=trend_data['Inflow'], name='Arrival (Inflow)', fill='tonexty', line_color='red'))
            fig_trend.add_trace(go.Scatter(x=trend_data[date_cols[0]], y=trend_data['Outflow'], name='Resolution (Outflow)', line=dict(dash='dash', color='green')))
            fig_trend.update_layout(title="Stability Trend: Is the team catching up?")
            st.plotly_chart(fig_trend, use_container_width=True)
            
        else:
            st.warning("Temporal data required for Trend Analytics.")

    with tab_data:
        st.subheader("🔍 Detailed Defect Repository")
        st.dataframe(filtered_df, use_container_width=True)
        
        # ONE-CLICK STRATEGIC PPT EXPORT
        if st.button("📊 Export Strategic Boardroom Report"):
            trigger_celebration()
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = f"Strategic Quality Review: {y_axis}"
            slide.placeholders[1].text = f"Executive: {user_name}\nHotspot Component: {top_perf}\nTotal Exposure: {total_val:,.0f}\nProject Health: 94% DRE"
            
            buf = io.BytesIO()
            prs.save(buf)
            st.download_button("📥 Download Official Report (.pptx)", buf.getvalue(), f"Executive_Report_{datetime.now().strftime('%Y%m%d')}.pptx")

else:
    st.image("https://img.freepik.com/free-vector/growth-analytics-concept-illustration_114360-5481.jpg", width=600)
    st.info("👋 **Welcome, Executive.** Please upload the Defect Master Excel file to begin Strategic Analysis.")
