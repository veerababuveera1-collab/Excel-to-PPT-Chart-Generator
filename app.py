import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import io

# --- 1. PAGE CONFIG & CSS ---
st.set_page_config(page_title="DataSlide Pro", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    .stButton>button { width: 100%; border-radius: 20px; background-color: #004b95; color: white; }
    .header-box { background-color: #004b95; padding: 20px; border-radius: 10px; color: white; text-align: center; }
    </style>
    """, unsafe_content_usage=True)

st.markdown('<div class="header-box"><h1>📊 Excel to PPT Automator</h1><p>Professional Reports in Seconds</p></div>', unsafe_content_usage=True)

# --- 2. DATA HANDLING ---
@st.cache_data
def load_sample_data():
    return pd.DataFrame({
        'Category': ['Q1', 'Q2', 'Q3', 'Q4'],
        'Revenue': [4500, 5200, 4800, 6100],
        'Profit': [1200, 1400, 1100, 1800]
    })

with st.sidebar:
    st.header("Upload Center")
    uploaded_file = st.file_uploader("Choose Excel File", type=["xlsx"])
    st.divider()
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
    else:
        df = load_sample_data()
        st.info("Using sample data. Upload your own to change.")

# --- 3. DASHBOARD UI ---
col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("Chart Controls")
    x_axis = st.selectbox("X-Axis", df.columns)
    y_axis = st.selectbox("Y-Axis", [c for c in df.columns if df[c].dtype != 'O'])
    chart_title = st.text_input("Slide Title", value="Performance Analysis")
    chart_color = st.color_picker("Brand Color", "#004b95")

with col2:
    fig = px.bar(df, x=x_axis, y=y_axis, title=chart_title, template="plotly_white", color_discrete_sequence=[chart_color])
    st.plotly_chart(fig, use_container_width=True)

# --- 4. THE STABLE EXPORT LOGIC (No Kaleido Required) ---
def create_ppt_stable(data, x, y, title, color):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5]) # Title only layout
    slide.shapes.title.text = title

    # Create a Matplotlib version of the chart for the PPT (Ultra-stable)
    plt.figure(figsize=(10, 6))
    plt.bar(data[x], data[y], color=color)
    plt.title(title)
    plt.xlabel(x)
    plt.ylabel(y)
    
    # Save Matplotlib to buffer
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', dpi=300, bbox_inches='tight')
    plt.close() # Close plot to save memory
    img_buf.seek(0)

    # Add image to slide
    slide.shapes.add_picture(img_buf, Inches(1), Inches(1.5), width=Inches(8))

    # Save PPT to buffer
    ppt_buf = io.BytesIO()
    prs.save(ppt_buf)
    ppt_buf.seek(0)
    return ppt_buf

st.divider()

if st.button("🚀 Generate and Download PowerPoint"):
    with st.spinner("Building your presentation..."):
        final_ppt = create_ppt_stable(df, x_axis, y_axis, chart_title, chart_color)
        
        st.download_button(
            label="📥 Download Your PPTX File",
            data=final_ppt,
            file_name="Executive_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        st.balloons()
