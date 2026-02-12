import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches
import io
from datetime import datetime

# --- 1. ENTERPRISE UI STYLING ---
st.set_page_config(page_title="DataSlide BI Enterprise", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f4f7f9; }
    .main-header {
        background: linear-gradient(90deg, #002e5d, #0056b3);
        color: white; padding: 2rem; border-radius: 15px;
        text-align: center; margin-bottom: 2rem;
    }
    .insight-box {
        background-color: #e8f0fe; padding: 20px; border-radius: 10px;
        border-left: 5px solid #0056b3; margin: 10px 0;
    }
    div[data-testid="stMetricValue"] { font-size: 32px; font-weight: 800; color: #002e5d; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION ---
if "auth" not in st.session_state:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🔒 Enterprise BI Portal")
        user = st.text_input("Authorized User Name")
        pw = st.text_input("Access Key", type="password")
        if st.button("Unlock Dashboard"):
            if pw == "Company2026" and user:
                st.session_state["auth"], st.session_state["user"] = True, user
                st.rerun()
            else: st.error("Invalid Credentials")
    st.stop()

# --- 3. SIDEBAR & SETTINGS ---
user_name = st.session_state["user"]
with st.sidebar:
    st.title(f"👤 {user_name}")
    chart_color = st.color_picker("Corporate Theme Color", "#002e5d")
    chart_type = st.selectbox("Chart Style", ["Area", "Bar", "Line", "Donut"])
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])
    if st.button("Logout"):
        del st.session_state["auth"]
        st.rerun()

# --- 4. DASHBOARD HEADER ---
st.markdown(f'<div class="main-header"><h1>🚀 DataSlide BI Enterprise</h1><p>Executive Insights for {user_name}</p></div>', unsafe_allow_html=True)

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    num_col = df.select_dtypes(include='number').columns[0]
    cat_col = df.select_dtypes(exclude='number').columns[0]
    
    # --- 5. AI-SUMMARY LOGIC (THE "PRO" FEATURE) ---
    total = df[num_col].sum()
    top_performer = df.groupby(cat_col)[num_col].sum().idxmax()
    avg_val = df[num_col].mean()
    
    st.markdown(f"""
    <div class="insight-box">
        <h3>💡 Strategic AI Summary for {user_name}</h3>
        <ul>
            <li><b>Total Performance:</b> Your total {num_col} is <b>{total:,.0f}</b>.</li>
            <li><b>Top Contributor:</b> The <b>{top_performer}</b> segment is currently leading the portfolio.</li>
            <li><b>Executive Insight:</b> The average value per entry is <b>{avg_val:,.0f}</b>. 
                Focus resources on {top_performer} to maximize ROI.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    # --- 6. METRICS & CHARTS ---
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Volume", f"{total:,.0f}")
    m2.metric("Top Segment", top_performer)
    m3.metric("Avg Value", f"{avg_val:,.0f}")

    if chart_type == "Area": fig = px.area(df, x=cat_col, y=num_col)
    elif chart_type == "Bar": fig = px.bar(df, x=cat_col, y=num_col)
    elif chart_type == "Line": fig = px.line(df, x=cat_col, y=num_col, markers=True)
    else: fig = px.pie(df, names=cat_col, values=num_col, hole=0.5)
    
    fig.update_traces(marker_color=chart_color) if chart_type != "Donut" else fig.update_traces(marker=dict(colors=[chart_color]))
    st.plotly_chart(fig, use_container_width=True)

    # --- 7. PPT EXPORT ---
    if st.button("📊 Export Executive PPT"):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = f"Monthly {num_col} Report"
        slide.placeholders[1].text = f"Prepared by: {user_name}\nKey Insight: {top_performer} is the leading segment."
        
        buf = io.BytesIO()
        prs.save(buf)
        st.download_button("📥 Download Presentation", buf.getvalue(), f"{user_name}_Report.pptx")
else:
    st.info("Please upload an Excel file to generate AI Insights.")
