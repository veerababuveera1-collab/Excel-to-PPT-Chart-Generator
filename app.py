import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
import io
from datetime import datetime

# --- 1. CONFIG & THEME (Enterprise Grade) ---
st.set_page_config(page_title="DataSlide BI Enterprise", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f0f2f6; }
    .main-header {
        background: linear-gradient(90deg, #063970, #154c79);
        color: white; padding: 1.5rem; border-radius: 10px;
        text-align: center; margin-bottom: 1rem; box-shadow: 0 4px 10px rgba(0,0,0,0.1);
    }
    .kpi-card {
        background-color: white; padding: 20px; border-radius: 10px;
        border-top: 5px solid #063970; box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        text-align: center;
    }
    div[data-testid="stMetricValue"] { font-size: 28px; color: #063970; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. CELEBRATION & UTILS ---
def trigger_celebration():
    st.balloons()
    st.markdown(f'<iframe src="https://www.soundjay.com/buttons/sounds/button-10.mp3" allow="autoplay" style="display:none"></iframe>', unsafe_allow_html=True)

# --- 3. AUTHENTICATION GATE ---
if "auth" not in st.session_state:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.title("🔒 Enterprise BI Portal")
        user = st.text_input("Username")
        pw = st.text_input("Access Key", type="password")
        if st.button("Unlock Dashboard"):
            if pw == "Company2026" and user:
                st.session_state["auth"], st.session_state["user"] = True, user
                trigger_celebration()
                st.rerun()
            else: st.error("Access Denied")
    st.stop()

# --- 4. SIDEBAR: GLOBAL SLICERS & CONTROLS ---
user_name = st.session_state["user"]
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/1534/1534003.png", width=80)
    st.markdown(f"### 👤 Executive: **{user_name}**")
    st.divider()
    
    uploaded_file = st.file_uploader("📂 Data Source (Excel)", type=["xlsx"])
    
    filtered_df = pd.DataFrame()
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        
        # AUTO-CLEAN NUMERICS (Fixes Currency/Text issues)
        for col in df.columns:
            if df[col].dtype == 'object':
                try: df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[$, ]', '', regex=True))
                except: pass
        
        # --- GLOBAL SLICERS ---
        st.subheader("🎯 Global Slicers")
        cat_cols = df.select_dtypes(exclude='number').columns.tolist()
        slicer_col = st.selectbox("Filter By Category", cat_cols)
        selected_options = st.multiselect(f"Select {slicer_col}", df[slicer_col].unique(), default=df[slicer_col].unique())
        
        # Temporal Intelligence (Date Detection)
        date_cols = [col for col in df.columns if 'date' in col.lower()]
        if date_cols:
            df[date_cols[0]] = pd.to_datetime(df[date_cols[0]])
            start_date, end_date = st.date_input("Time Range", [df[date_cols[0]].min(), df[date_cols[0]].max()])
            mask = (df[slicer_col].isin(selected_options)) & \
                   (df[date_cols[0]] >= pd.Timestamp(start_date)) & \
                   (df[date_cols[0]] <= pd.Timestamp(end_date))
            filtered_df = df[mask]
        else:
            filtered_df = df[df[slicer_col].isin(selected_options)]

        st.divider()
        chart_color = st.color_picker("Brand Theme Color", "#063970")
        chart_type = st.selectbox("Visualization", ["Bar", "Area", "Line", "Donut"])

    if st.button("🚪 Logout"):
        del st.session_state["auth"]
        st.rerun()

# --- 5. MAIN INTERFACE ---
st.markdown(f'<div class="main-header"><h1>🚀 DataSlide BI Enterprise</h1></div>', unsafe_allow_html=True)

if not filtered_df.empty:
    num_cols = filtered_df.select_dtypes(include=['number']).columns.tolist()
    x_axis = st.sidebar.selectbox("X-Axis (Dimension)", filtered_df.columns, index=0)
    y_axis = st.sidebar.selectbox("Y-Axis (Metric)", num_cols, index=0)

    # --- TABS SYSTEM ---
    tab_dash, tab_trends, tab_data = st.tabs(["📊 Executive Dashboard", "📈 Trend Analysis", "📋 Raw Data Fields"])

    with tab_dash:
        processed_df = filtered_df.groupby(x_axis)[y_axis].sum().reset_index()
        total_val = processed_df[y_axis].sum()
        top_perf = processed_df.loc[processed_df[y_axis].idxmax(), x_axis]
        
        # KPI Metrics
        k1, k2, k3 = st.columns(3)
        k1.metric(f"Total {y_axis}", f"{total_val:,.0f}")
        k2.metric("Leader", top_perf)
        k3.metric("Avg Benchmark", f"{processed_df[y_axis].mean():,.0f}")

        # Prescriptive AI Insight
        st.info(f"💡 **AI Insight:** {top_perf} is the primary driver, contributing {((processed_df[y_axis].max()/total_val)*100):.1f}% of volume. We recommend focusing resource allocation here.")

        # Chart Engine
        if chart_type == "Area": fig = px.area(processed_df, x=x_axis, y=y_axis)
        elif chart_type == "Bar": fig = px.bar(processed_df, x=x_axis, y=y_axis)
        elif chart_type == "Line": fig = px.line(processed_df, x=x_axis, y=y_axis, markers=True)
        else: fig = px.pie(processed_df, names=x_axis, values=y_axis, hole=0.4)
        
        fig.update_traces(marker_color=chart_color) if chart_type != "Donut" else fig.update_traces(marker=dict(colors=[chart_color]))
        st.plotly_chart(fig, use_container_width=True)

    with tab_trends:
        st.subheader("📅 Temporal Performance Evolution")
        if date_cols:
            trend_df = filtered_df.groupby(date_cols[0])[y_axis].sum().reset_index()
            fig_trend = px.line(trend_df, x=date_cols[0], y=y_axis, title="Performance Over Time", markers=True)
            st.plotly_chart(fig_trend, use_container_width=True)
        else:
            st.warning("No date column detected for trend analysis.")

    with tab_data:
        st.subheader("🔍 Granular Data Fields")
        st.dataframe(filtered_df, use_container_width=True)
        
        if st.button("📊 Export Executive PPT"):
            trigger_celebration()
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = f"Strategic Review: {y_axis}"
            slide.placeholders[1].text = f"Generated by: {user_name}\nTop Performer: {top_perf}\nTotal: {total_val:,.0f}"
            
            buf = io.BytesIO()
            prs.save(buf)
            st.download_button("📥 Download Official Report", buf.getvalue(), f"Report_{datetime.now().strftime('%Y%m%d')}.pptx")

else:
    st.image("https://img.freepik.com/free-vector/growth-analytics-concept-illustration_114360-5481.jpg", width=600)
    st.info("👋 **Welcome to the BI Command Center.** Upload an Excel file to begin.")
