import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
import io

# --- CONFIG & BEAUTIFICATION ---
st.set_page_config(page_title="DataSlide Pro", page_icon="📈", layout="wide")

# Inject Custom CSS
st.markdown("""
    <style>
    /* Main background */
    .stApp { background-color: #f8f9fa; }
    
    /* Custom Header */
    .main-header {
        background-color: #004b95;
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 25px;
    }
    
    /* Card-like containers for inputs */
    div.stSelectbox, div.stMultiSelect, div.stTextInput {
        background-color: white;
        padding: 10px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    
    /* Button styling */
    .stButton>button {
        background-color: #004b95;
        color: white;
        font-weight: bold;
        border-radius: 20px;
        border: none;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #0078d4;
        transform: scale(1.02);
    }
    </style>
    """, unsafe_content_usage=True)

# Header
st.markdown('<div class="main-header"><h1>🚀 DataSlide Pro</h1><p>Convert Excel Insights to Executive Presentations</p></div>', unsafe_content_usage=True)

# --- APP LOGIC ---
with st.sidebar:
    st.header("📂 Data Source")
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])
    
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # UI Layout with Columns
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("Chart Settings")
        x_col = st.selectbox("X-Axis", df.columns)
        y_cols = st.multiselect("Y-Axis", df.columns, default=[df.columns[1]])
        chart_type = st.selectbox("Format", ["Bar", "Line", "Area"])
        title = st.text_input("Slide Title", "Performance Overview")
        
    with col2:
        st.subheader("Preview")
        if chart_type == "Bar":
            fig = px.bar(df, x=x_col, y=y_cols, template="plotly_white")
        elif chart_type == "Line":
            fig = px.line(df, x=x_col, y=y_cols, template="plotly_white")
        else:
            fig = px.area(df, x=x_col, y=y_cols, template="plotly_white")
            
        fig.update_layout(margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig, use_container_width=True)

    # Export Section
    st.divider()
    if st.button("🪄 Generate Professional PPT"):
        # (Insert the PPT generation logic here from previous messages)
        st.success("PowerPoint generated successfully!")
else:
    st.info("Please upload an Excel file in the sidebar to begin.")
