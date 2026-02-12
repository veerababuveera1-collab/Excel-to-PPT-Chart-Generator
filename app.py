import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
import io
from datetime import datetime

# --- 1. CONFIG & THEME ---
st.set_page_config(page_title="DataSlide BI Enterprise", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f4f7f9; }
    .main-header {
        background: linear-gradient(90deg, #002e5d, #0056b3);
        color: white; padding: 2rem; border-radius: 15px; text-align: center; margin-bottom: 2rem;
    }
    .insight-box {
        background-color: #ffffff; padding: 25px; border-radius: 15px;
        border-left: 8px solid #0056b3; margin: 15px 0; box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    }
    .future-box {
        background-color: #fff9db; padding: 15px; border-radius: 10px; border: 1px dashed #f59f00;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. CELEBRATION LOGIC ---
def trigger_celebration():
    st.balloons()
    sound_url = "https://www.soundjay.com/buttons/sounds/button-10.mp3"
    st.markdown(f'<iframe src="{sound_url}" allow="autoplay" style="display:none"></iframe>', unsafe_allow_html=True)

# --- 3. AUTHENTICATION ---
if "auth" not in st.session_state:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🔒 Strategic BI Portal")
        user = st.text_input("Username")
        pw = st.text_input("Access Key", type="password")
        if st.button("Unlock Dashboard"):
            if pw == "Company2026" and user:
                st.session_state["auth"], st.session_state["user"] = True, user
                trigger_celebration()
                st.rerun()
            else: st.error("Invalid Credentials")
    st.stop()

# --- 4. SIDEBAR CONTROLS ---
user_name = st.session_state["user"]
with st.sidebar:
    st.markdown(f"### 👤 Welcome, **{user_name}**")
    chart_color = st.color_picker("Brand Theme Color", "#002e5d")
    chart_type = st.selectbox("Visualization Engine", ["Area", "Bar", "Line", "Donut"])
    uploaded_file = st.file_uploader("📂 Data Source (Excel)", type=["xlsx"])
    
    x_axis, y_axis = None, None
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        # AUTO-CLEAN: Fix the "Y-axis disabled" issue
        for col in df.columns:
            if df[col].dtype == 'object':
                try:
                    df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[$, ]', '', regex=True))
                except: pass
        
        st.subheader("📐 Axis Mapping")
        x_axis = st.selectbox("X-Axis (Dimension)", df.columns)
        num_cols = df.select_dtypes(include=['number']).columns.tolist()
        y_axis = st.selectbox("Y-Axis (Metric)", num_cols) if num_cols else None

    if st.button("Logout"):
        del st.session_state["auth"]
        st.rerun()

# --- 5. MAIN INTERFACE ---
st.markdown(f'<div class="main-header"><h1>🚀 DataSlide BI Enterprise</h1><p>Strategic Performance Insights for {user_name}</p></div>', unsafe_allow_html=True)

if uploaded_file and x_axis and y_axis:
    # --- AI LOGIC ---
    total = df[y_axis].sum()
    top_performer = df.groupby(x_axis)[y_axis].sum().idxmax()
    avg_val = df[y_axis].mean()
    
    st.markdown(f"""
    <div class="insight-box">
        <h3>💡 Executive AI Summary</h3>
        <p>Current analysis of <b>{y_axis}</b> shows <b>{top_performer}</b> as the lead contributor.</p>
        <div class="future-box">
            <b>🔮 Prescriptive Insight:</b> {top_performer} is performing {((df.groupby(x_axis)[y_axis].sum().max()/avg_val)-1)*100:.1f}% 
            above average. We recommend reallocating Q3 resources to this segment to maximize ROI.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Visualization
    if chart_type == "Area": fig = px.area(df, x=x_axis, y=y_axis, template="plotly_white")
    elif chart_type == "Bar": fig = px.bar(df, x=x_axis, y=y_axis, template="plotly_white")
    elif chart_type == "Line": fig = px.line(df, x=x_axis, y=y_axis, template="plotly_white", markers=True)
    else: fig = px.pie(df, names=x_axis, values=y_axis, hole=0.5, template="plotly_white")
    
    fig.update_traces(marker_color=chart_color) if chart_type != "Donut" else fig.update_traces(marker=dict(colors=[chart_color]))
    st.plotly_chart(fig, use_container_width=True, key=f"viz_{chart_type}_{x_axis}")

    # Export
    if st.button("📊 Export Strategic PPT"):
        trigger_celebration()
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = f"Executive Review: {y_axis}"
        slide.placeholders[1].text = f"Top Performer: {top_performer}\nTotal Volume: {total:,.0f}"
        buf = io.BytesIO()
        prs.save(buf)
        st.download_button("📥 Download Report", buf.getvalue(), f"Report_{user_name}.pptx")
else:
    # --- Missing Functionality: The Welcome Screen ---
    st.image("https://img.freepik.com/free-vector/growth-analytics-concept-illustration_114360-5481.jpg", width=500)
    st.info("👋 Welcome to the Enterprise Portal. Please upload an Excel file in the sidebar to generate AI Insights.")
