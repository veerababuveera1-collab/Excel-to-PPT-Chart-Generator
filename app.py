import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import io

# --- 1. PAGE CONFIG & BRANDING ---
st.set_page_config(page_title="DataSlide Pro", layout="wide", page_icon="📊")

st.markdown("""
    <style>
    .stApp { background-color: #f4f7f9; }
    .main-header {
        background-color: #004b95;
        padding: 30px;
        border-radius: 15px;
        color: white;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .stButton>button {
        width: 100%;
        border-radius: 25px;
        background-color: #004b95;
        color: white;
        font-weight: bold;
        border: none;
        height: 3em;
    }
    .stMetric {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<div class="main-header"><h1>🚀 DataSlide Pro</h1><p>Automated Business Intelligence & PowerPoint Reporting</p></div>', unsafe_allow_html=True)

# --- 2. DATA ENGINE ---
@st.cache_data
def get_default_data():
    return pd.DataFrame({
        'Department': ['Sales', 'Marketing', 'R&D', 'HR', 'Ops'],
        'Revenue': [120000, 85000, 95000, 40000, 110000],
        'Expenses': [80000, 90000, 100000, 35000, 95000]
    })

# --- 3. SIDEBAR ---
with st.sidebar:
    st.header("📂 Data Center")
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
    else:
        df = get_default_data()
        st.info("Using demo data.")

# --- 4. DASHBOARD UI ---
c1, c2 = st.columns([1, 2])

with c1:
    st.subheader("🛠️ Settings")
    all_cols = df.columns.tolist()
    num_cols = df.select_dtypes(include=['number']).columns.tolist()
    
    x_axis = st.selectbox("X-Axis", all_cols)
    y_axis = st.selectbox("Y-Axis", num_cols)
    chart_type = st.radio("Style", ["Bar", "Line", "Area"], horizontal=True)
    chart_color = st.color_picker("Color", "#004b95")
    slide_title = st.text_input("Slide Title", value=f"{y_axis} Analysis")

with c2:
    st.subheader("📊 Visual Preview")
    if chart_type == "Bar":
        fig = px.bar(df, x=x_axis, y=y_axis, template="plotly_white", color_discrete_sequence=[chart_color])
    elif chart_type == "Line":
        fig = px.line(df, x=x_axis, y=y_axis, template="plotly_white", color_discrete_sequence=[chart_color])
    else:
        fig = px.area(df, x=x_axis, y=y_axis, template="plotly_white", color_discrete_sequence=[chart_color])
    st.plotly_chart(fig, use_container_width=True)

# --- 5. NEW: SUMMARY TABLE & KPIs ---
st.divider()
st.subheader("📋 Executive Summary Table")

# Row for Metrics
m1, m2, m3 = st.columns(3)
m1.metric(f"Total {y_axis}", f"{df[y_axis].sum():,.0f}")
m2.metric(f"Average {y_axis}", f"{df[y_axis].mean():,.0f}")
m3.metric(f"Max {y_axis}", f"{df[y_axis].max():,.0f}")

# The Summary Table
st.dataframe(df[[x_axis, y_axis]].style.format({y_axis: "{:,.2f}"}), use_container_width=True)

# --- 6. EXPORT LOGIC ---
def build_ppt(data, x, y, title, color):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title

    plt.figure(figsize=(10, 6))
    if chart_type == "Bar": plt.bar(data[x], data[y], color=color)
    else: plt.plot(data[x], data[y], color=color, marker='o')
    
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', dpi=300)
    plt.close()
    img_buf.seek(0)
    slide.shapes.add_picture(img_buf, Inches(1), Inches(1.5), width=Inches(8))

    ppt_buf = io.BytesIO()
    prs.save(ppt_buf)
    ppt_buf.seek(0)
    return ppt_buf

if st.button("🚀 Export to PowerPoint"):
    ppt_out = build_ppt(df, x_axis, y_axis, slide_title, chart_color)
    st.download_button("📥 Download PPT", data=ppt_out, file_name="Report.pptx")
    st.balloons()
