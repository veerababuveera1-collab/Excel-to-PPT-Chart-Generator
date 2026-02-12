import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import io

# --- 1. PAGE CONFIG & BRANDING ---
st.set_page_config(page_title="DataSlide Pro", layout="wide", page_icon="📊")

# Custom CSS for a professional "Enterprise" look
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
    .stButton>button:hover {
        background-color: #0078d4;
        border: none;
    }
    div[data-testid="stExpander"] {
        background-color: white;
        border-radius: 10px;
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

# --- 3. SIDEBAR & FILE INPUT ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/732/732220.png", width=50) # Generic Excel Icon
    st.header("Data Control Center")
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    st.divider()
    
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.success("File successfully loaded!")
    else:
        df = get_default_data()
        st.info("💡 Using demo data. Upload your own to generate custom reports.")

# --- 4. INTERACTIVE DASHBOARD ---
c1, c2 = st.columns([1, 2])

with c1:
    st.subheader("🛠️ Chart Settings")
    with st.container():
        all_cols = df.columns.tolist()
        num_cols = df.select_dtypes(include=['number']).columns.tolist()
        
        x_axis = st.selectbox("Select X-Axis (Categories)", all_cols)
        y_axis = st.selectbox("Select Y-Axis (Values)", num_cols)
        chart_type = st.radio("Chart Type", ["Bar", "Line", "Area"], horizontal=True)
        chart_color = st.color_picker("Brand Primary Color", "#004b95")
        slide_title = st.text_input("PPT Slide Title", value=f"{y_axis} by {x_axis}")

with c2:
    st.subheader("📊 Visual Preview")
    if chart_type == "Bar":
        fig = px.bar(df, x=x_axis, y=y_axis, template="plotly_white", color_discrete_sequence=[chart_color])
    elif chart_type == "Line":
        fig = px.line(df, x=x_axis, y=y_axis, template="plotly_white", color_discrete_sequence=[chart_color])
    else:
        fig = px.area(df, x=x_axis, y=y_axis, template="plotly_white", color_discrete_sequence=[chart_color])
    
    fig.update_layout(margin=dict(l=0, r=0, t=30, b=0))
    st.plotly_chart(fig, use_container_width=True)

# --- 5. POWERPOINT EXPORT LOGIC ---
def build_presentation(data, x, y, title, color):
    prs = Presentation()
    
    # Title Slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Executive Summary Report"
    title_slide.placeholders[1].text = "Generated Automatically via DataSlide Pro"

    # Content Slide with Chart
    chart_slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_slide.shapes.title.text = title

    # Generate stable image using Matplotlib (Bypasses Kaleido errors)
    plt.figure(figsize=(10, 6))
    plt.bar(data[x], data[y], color=color) if chart_type == "Bar" else plt.plot(data[x], data[y], color=color, marker='o')
    plt.title(title, fontsize=14)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', dpi=300, bbox_inches='tight')
    plt.close()
    img_buf.seek(0)

    # Insert Image to PPT
    chart_slide.shapes.add_picture(img_buf, Inches(1), Inches(1.5), width=Inches(8))

    # Save to Buffer
    ppt_buf = io.BytesIO()
    prs.save(ppt_buf)
    ppt_buf.seek(0)
    return ppt_buf

st.divider()

# Center the button
_, btn_col, _ = st.columns([1, 1, 1])
with btn_col:
    if st.button("🪄 Export to PowerPoint"):
        with st.spinner("Processing High-Res Graphics..."):
            ppt_out = build_presentation(df, x_axis, y_axis, slide_title, chart_color)
            st.download_button(
                label="📥 Download Presentation",
                data=ppt_out,
                file_name="Business_Report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            st.balloons()
