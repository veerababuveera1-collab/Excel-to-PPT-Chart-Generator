import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
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
        border-left: 5px solid #0056b3; margin: 15px 0;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. SOUND FUNCTION ---
def play_sound():
    # A professional "Success" chime
    sound_url = "https://www.soundjay.com/buttons/sounds/button-10.mp3"
    st.markdown(f'<iframe src="{sound_url}" allow="autoplay" style="display:none"></iframe>', unsafe_allow_html=True)

# --- 3. AUTHENTICATION ---
if "auth" not in st.session_state:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.title("🔒 Enterprise BI Portal")
        user = st.text_input("Authorized User Name")
        pw = st.text_input("Access Key", type="password")
        if st.button("Unlock Dashboard"):
            if pw == "Company2026" and user:
                st.session_state["auth"], st.session_state["user"] = True, user
                st.balloons() # 🎉 Balloons on Login
                play_sound()  # 🔊 Chime on Login
                st.rerun()
            else: st.error("Invalid Credentials")
    st.stop()

# --- 4. SIDEBAR CONTROLS ---
user_name = st.session_state["user"]
with st.sidebar:
    st.markdown(f"### 👤 User: **{user_name}**")
    st.divider()
    # COLOR PICKER (Right here in the sidebar!)
    chart_color = st.color_picker("Corporate Theme Color", "#002e5d")
    chart_type = st.selectbox("Chart Style", ["Area", "Bar", "Line", "Donut"])
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])
    if st.button("Logout"):
        del st.session_state["auth"]
        st.rerun()

# --- 5. MAIN DASHBOARD ---
st.markdown(f'<div class="main-header"><h1>🚀 DataSlide BI Enterprise</h1><p>Executive Insights for {user_name}</p></div>', unsafe_allow_html=True)

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    num_col = df.select_dtypes(include='number').columns[0]
    cat_col = df.select_dtypes(exclude='number').columns[0]
    
    # AI Summary
    total = df[num_col].sum()
    top_performer = df.groupby(cat_col)[num_col].sum().idxmax()
    
    st.markdown(f"""
    <div class="insight-box">
        <h3>💡 AI Summary: {top_performer} is leading with {total:,.0f} total units.</h3>
    </div>
    """, unsafe_allow_html=True)

    # Dynamic Chart
    if chart_type == "Area": fig = px.area(df, x=cat_col, y=num_col)
    elif chart_type == "Bar": fig = px.bar(df, x=cat_col, y=num_col)
    elif chart_type == "Line": fig = px.line(df, x=cat_col, y=num_col, markers=True)
    else: fig = px.pie(df, names=cat_col, values=num_col, hole=0.5)
    
    fig.update_traces(marker_color=chart_color) if chart_type != "Donut" else fig.update_traces(marker=dict(colors=[chart_color]))
    st.plotly_chart(fig, use_container_width=True, key=f"v_{chart_type}")

    # PPT Export with Balloons
    if st.button("📊 Export Executive PPT"):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = f"Report for {user_name}"
        
        buf = io.BytesIO()
        prs.save(buf)
        
        st.balloons() # 🎉 Balloons on success
        play_sound()  # 🔊 Sound on success
        st.download_button("📥 Click to Download", buf.getvalue(), f"Report.pptx")
else:
    st.info("Upload your file to see the magic happen! ✨")
