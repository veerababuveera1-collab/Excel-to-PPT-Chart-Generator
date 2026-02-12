import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import os

# --- 1. SETTINGS & STYLING ---
st.set_page_config(page_title="Excel2PPT Pro", layout="wide", page_icon="📊")

st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #0078D4; color: white; }
    </style>
    """, unsafe_content_usage=True)

# --- 2. DATA ENGINE ---
def get_sample_data():
    """Generates a sample dataframe if no file is uploaded."""
    return pd.DataFrame({
        'Department': ['Sales', 'Marketing', 'IT', 'HR', 'Finance', 'Ops'],
        'Budget': [45000, 32000, 58000, 21000, 39000, 42000],
        'Actual_Spend': [42000, 35000, 60000, 19000, 38000, 44000],
        'Headcount': [15, 8, 22, 5, 10, 18]
    })

# --- 3. UI SIDEBAR ---
with st.sidebar:
    st.title("⚙️ Controls")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.success("File Uploaded!")
    else:
        df = get_sample_data()
        st.info("💡 Using sample data. Upload your own to replace.")
    
    st.divider()
    chart_type = st.selectbox("Chart Style", ["Bar", "Line", "Scatter", "Area"])
    color_theme = st.color_picker("Pick Chart Color", "#0078D4")

# --- 4. MAIN INTERFACE ---
st.title("📊 Excel to PowerPoint Automator")
st.subheader("Interactive Data Visualization")

cols = df.columns.tolist()
c1, c2, c3 = st.columns([2, 2, 2])

with c1:
    x_axis = st.selectbox("X-Axis (Categories)", cols, index=0)
with c2:
    y_axis = st.multiselect("Y-Axis (Numeric Values)", [c for c in cols if df[c].dtype in ['int64', 'float64']], default=[cols[1]])
with c3:
    slide_title = st.text_input("Slide Title", value="Monthly Performance Report")

# Create Chart
if chart_type == "Bar":
    fig = px.bar(df, x=x_axis, y=y_axis, barmode="group", template="plotly_white", color_discrete_sequence=[color_theme])
elif chart_type == "Line":
    fig = px.line(df, x=x_axis, y=y_axis, template="plotly_white", color_discrete_sequence=[color_theme])
elif chart_type == "Scatter":
    fig = px.scatter(df, x=x_axis, y=y_axis, template="plotly_white", color_discrete_sequence=[color_theme])
else:
    fig = px.area(df, x=x_axis, y=y_axis, template="plotly_white", color_discrete_sequence=[color_theme])

st.plotly_chart(fig, use_container_width=True)

# --- 5. POWERPOINT GENERATION ---
def create_ppt(figure, title):
    prs = Presentation()
    
    # Add a Title Slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    slide.shapes.title.text = "Automated Data Report"
    slide.placeholders[1].text = "Generated via Python Streamlit Tool"

    # Add Chart Slide
    chart_slide_layout = prs.slide_layouts[5] # Blank layout
    slide = prs.slides.add_slide(chart_slide_layout)
    
    # Set Slide Title
    title_shape = slide.shapes.title
    title_shape.text = title
    
    # Convert Plotly to Image (requires kaleido)
    img_bytes = figure.to_image(format="png", width=1200, height=700, scale=2)
    img_stream = io.BytesIO(img_bytes)
    
    # Add Image to Slide
    slide.shapes.add_picture(img_stream, Inches(0.5), Inches(1.5), width=Inches(9))
    
    # Save to memory
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

st.divider()

if st.button("🚀 Generate & Download Presentation"):
    with st.spinner("Converting chart to high-res image..."):
        ppt_file = create_ppt(fig, slide_title)
        
        st.download_button(
            label="Click here to download .PPTX",
            data=ppt_file,
            file_name="Business_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        st.balloons()
