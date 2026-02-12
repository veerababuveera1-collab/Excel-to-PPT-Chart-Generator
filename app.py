import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import io
from datetime import datetime

# --- 1. THE SECURITY LAYER ---
def check_auth():
    """Polished Login Interface"""
    if "authenticated" not in st.session_state:
        st.markdown("""
            <style>
            .login-container {
                max-width: 450px;
                margin: 100px auto;
                padding: 30px;
                background-color: white;
                border-radius: 15px;
                box-shadow: 0 4px 15px rgba(0,0,0,0.1);
                text-align: center;
            }
            .stApp { background-color: #f0f2f6; }
            </style>
        """, unsafe_allow_html=True)

        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        st.image("https://cdn-icons-png.flaticon.com/512/3064/3064155.png", width=70)
        st.title("Enterprise BI Portal")
        
        user_name = st.text_input("Authorized User Name", placeholder="Enter your name")
        password = st.text_input("Access Key", type="password")
        
        if st.button("Unlock Dashboard"):
            if password == "Company2026" and user_name:
                st.session_state["authenticated"] = True
                st.session_state["user_name"] = user_name
                st.balloons()
                st.rerun()
            else:
                st.error("Invalid Credentials or Name Missing")
        st.markdown('</div>', unsafe_allow_html=True)
        return False
    return True

# --- 2. THE MAIN APPLICATION ---
def run_bi_app():
    user = st.session_state["user_name"]
    
    # --- PAGE CONFIG & BRANDING ---
    st.set_page_config(page_title=f"Dashboard | {user}", layout="wide", page_icon="📈")
    
    st.markdown(f"""
        <style>
        .main-header {{
            background: linear-gradient(90deg, #002e5d, #0056b3);
            color: white; padding: 25px; border-radius: 12px;
            text-align: center; margin-bottom: 30px;
        }}
        .stMetric {{ background-color: white; padding: 15px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }}
        </style>
        <div class="main-header">
            <h1>🚀 DataSlide BI Enterprise</h1>
            <p>Welcome, <b>{user}</b> | Strategy & Insights Dashboard</p>
        </div>
    """, unsafe_allow_html=True)

    # --- SIDEBAR & DATA ---
    with st.sidebar:
        st.title(f"👤 {user}")
        uploaded_file = st.file_uploader("Upload Corporate Excel", type=["xlsx"])
        st.divider()
        if st.button("Logout"):
            del st.session_state["authenticated"]
            st.rerun()

    # Data Loader
    @st.cache_data
    def load_data(file):
        if file: return pd.read_excel(file)
        return pd.DataFrame({
            'Period': ['Q1', 'Q2', 'Q3', 'Q4'],
            'Revenue': [120000, 155000, 140000, 190000],
            'Expenses': [80000, 85000, 90000, 95000]
        })

    df = load_data(uploaded_file)
    
    # --- BI CALCULATIONS (TRENDS) ---
    num_cols = df.select_dtypes(include=['number']).columns.tolist()
    cat_cols = df.select_dtypes(exclude=['number']).columns.tolist()
    
    y_axis = st.selectbox("Select Key Performance Indicator (KPI)", num_cols)
    x_axis = st.selectbox("Select Reporting Dimension", cat_cols if cat_cols else df.columns)
    
    df['Trend %'] = df[y_axis].pct_change() * 100
    avg_growth = df['Trend %'].mean()

    # --- EXECUTIVE SCORECARD ---
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Volume", f"{df[y_axis].sum():,.0f}")
    m2.metric("Growth Velocity", f"{avg_growth:.1f}%", delta=f"{avg_growth:.1f}%")
    m3.metric("Max Peak", f"{df[y_axis].max():,.0f}")

    # --- INTERACTIVE DASHBOARD ---
    c1, c2 = st.columns([2, 1])
    
    with c1:
        fig = px.area(df, x=x_axis, y=y_axis, template="plotly_white", 
                      color_discrete_sequence=["#002e5d"], title=f"{y_axis} Trend Analysis")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.write("📋 **Data Summary Table**")
        st.dataframe(df[[x_axis, y_axis, 'Trend %']].style.format({'Trend %': '{:.1f}%', y_axis: '{:,.0f}'}), 
                     hide_index=True, use_container_width=True)

    # --- POWERPOINT EXPORT ENGINE ---
    def generate_ppt(data, x, y, user_name):
        prs = Presentation()
        
        # Slide 1: Cover
        slide1 = prs.slides.add_slide(prs.slide_layouts[0])
        slide1.shapes.title.text = f"Executive Summary: {y}"
        slide1.placeholders[1].text = f"Author: {user_name}\nDate: {datetime.now().strftime('%Y-%m-%d')}\nConfidential"

        # Slide 2: Chart
        slide2 = prs.slides.add_slide(prs.slide_layouts[5])
        slide2.shapes.title.text = f"{y} Performance by {x}"
        
        plt.figure(figsize=(10, 5))
        plt.fill_between(data[x], data[y], color='#002e5d', alpha=0.3)
        plt.plot(data[x], data[y], color='#002e5d', marker='o', linewidth=3)
        
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', dpi=300, bbox_inches='tight')
        img_buf.seek(0)
        slide2.shapes.add_picture(img_buf, Inches(0.5), Inches(1.5), width=Inches(9))
        
        ppt_buf = io.BytesIO()
        prs.save(ppt_buf)
        ppt_buf.seek(0)
        return ppt_buf

    st.divider()
    if st.button(f"🚀 Prepare PowerPoint Deck for {user}"):
        with st.spinner("Automating Slide Creation..."):
            final_ppt = generate_ppt(df, x_axis, y_axis, user)
            st.download_button(f"📥 Download {user}'s Report", data=final_ppt, 
                               file_name=f"Executive_Report_{user}.pptx")

# --- 3. RUN LOGIC ---
if check_auth():
    run_bi_app()
