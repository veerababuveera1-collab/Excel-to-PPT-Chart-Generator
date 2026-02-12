import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches
import io
from datetime import datetime

# --- 1. PROFESSIONAL UI STYLING (THE PRO SECRET) ---
st.set_page_config(page_title="DataSlide Enterprise", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f4f7f9; }
    .main-header {
        background: linear-gradient(90deg, #002e5d, #0056b3);
        color: white; padding: 2rem; border-radius: 15px;
        text-align: center; margin-bottom: 2rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    [data-testid="stMetricValue"] { font-size: 32px; font-weight: 800; color: #002e5d; }
    .stButton>button {
        width: 100%; border-radius: 10px; height: 3em;
        background-color: #002e5d; color: white; font-weight: bold;
        border: none; transition: 0.3s;
    }
    .stButton>button:hover { background-color: #0056b3; transform: translateY(-2px); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. SECURITY LAYER ---
if "auth" not in st.session_state:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.image("https://cdn-icons-png.flaticon.com/512/3064/3064155.png", width=80)
        st.title("Enterprise BI Portal")
        user = st.text_input("Authorized User Name")
        pw = st.text_input("Access Key", type="password")
        if st.button("Unlock Dashboard"):
            if pw == "Company2026" and user:
                st.session_state["auth"] = True
                st.session_state["user"] = user
                st.rerun()
            else:
                st.error("Invalid Credentials")
    st.stop()

# --- 3. MAIN DASHBOARD ---
user_name = st.session_state["user"]

with st.sidebar:
    st.markdown(f"### 👤 {user_name}")
    st.divider()
    # THE COLOR PICKER OPTION
    chart_theme = st.color_picker("🎨 Theme Color", "#002e5d")
    uploaded_file = st.file_uploader("Upload Monthly Data", type=["xlsx"])
    if st.button("🚪 Logout"):
        del st.session_state["auth"]
        st.rerun()

st.markdown(f'<div class="main-header"><h1>🚀 DataSlide BI Enterprise</h1><p>Welcome, {user_name} | Executive Intelligence System</p></div>', unsafe_allow_html=True)

# Data Loading
df = pd.read_excel(uploaded_file) if uploaded_file else pd.DataFrame({
    'Period': ['Q1', 'Q2', 'Q3', 'Q4'], 'Revenue': [120000, 155000, 140000, 190000]
})

# Metrics & Trends
df['Trend %'] = df.select_dtypes(include='number').iloc[:,0].pct_change() * 100
val_col = df.select_dtypes(include='number').columns[0]

m1, m2, m3 = st.columns(3)
m1.metric("Total Volume", f"{df[val_col].sum():,.0f}")
m2.metric("Avg Growth", f"{df['Trend %'].mean():.1f}%")
m3.metric("Peak Value", f"{df[val_col].max():,.0f}")

# Dynamic Chart with Selected Color
fig = px.area(df, x=df.columns[0], y=val_col, template="plotly_white")
fig.update_traces(line_color=chart_theme, fillcolor=chart_theme)
st.plotly_chart(fig, use_container_width=True)

# PPT Export
if st.button(f"📊 Generate Executive Deck for {user_name}"):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = f"Business Review: {val_col}"
    slide.placeholders[1].text = f"Prepared by: {user_name}\nDate: {datetime.now().strftime('%Y-%m-%d')}"
    
    ppt_buf = io.BytesIO()
    prs.save(ppt_buf)
    st.download_button("📥 Download Report", data=ppt_buf.getvalue(), file_name="Report.pptx")
