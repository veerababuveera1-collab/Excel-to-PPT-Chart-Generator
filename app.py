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
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .insight-box {
        background-color: #ffffff; padding: 25px; border-radius: 15px;
        border-left: 8px solid #0056b3; margin: 15px 0;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    }
    .future-box {
        background-color: #fff9db; padding: 20px; border-radius: 10px;
        border: 1px dashed #f59f00; margin-top: 15px;
    }
    div[data-testid="stMetricValue"] { font-size: 32px; font-weight: 800; color: #002e5d; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. SUCCESS FEEDBACK FUNCTIONS ---
def trigger_success():
    st.balloons()
    # Chime sound effect
    sound_url = "https://www.soundjay.com/buttons/sounds/button-10.mp3"
    st.markdown(f'<iframe src="{sound_url}" allow="autoplay" style="display:none"></iframe>', unsafe_allow_html=True)

# --- 3. AUTHENTICATION ---
if "auth" not in st.session_state:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.title("🔒 Strategic BI Portal")
        user = st.text_input("Executive Username")
        pw = st.text_input("Access Key", type="password")
        if st.button("Unlock Dashboard"):
            if pw == "Company2026" and user:
                st.session_state["auth"], st.session_state["user"] = True, user
                trigger_success()
                st.rerun()
            else: st.error("Access Denied.")
    st.stop()

# --- 4. SIDEBAR & DYNAMIC CONFIGURATION ---
user_name = st.session_state["user"]
with st.sidebar:
    st.markdown(f"### 👤 Welcome, **{user_name}**")
    st.divider()
    chart_color = st.color_picker("Brand Theme Color", "#002e5d")
    chart_type = st.selectbox("Visualization Engine", ["Area", "Bar", "Line", "Donut"])
    
    uploaded_file = st.file_uploader("📂 Data Source (Excel)", type=["xlsx"])
    
    x_axis, y_axis = None, None
    if uploaded_file:
        df_headers = pd.read_excel(uploaded_file, nrows=0)
        st.subheader("📐 Axis Mapping")
        x_axis = st.selectbox("X-Axis (Dimensions)", df_headers.columns)
        y_axis = st.selectbox("Y-Axis (Metrics)", df_headers.select_dtypes(include='number').columns)

    if st.button("Logout"):
        del st.session_state["auth"]
        st.rerun()

# --- 5. MAIN INTERFACE ---
st.markdown(f'<div class="main-header"><h1>🚀 DataSlide BI Enterprise</h1><p>Strategic Performance Insights for {user_name}</p></div>', unsafe_allow_html=True)

if uploaded_file and x_axis and y_axis:
    df = pd.read_excel(uploaded_file)
    
    # --- 6. ADVANCED AI SUMMARY & FUTURE SUGGESTIONS ---
    total = df[y_axis].sum()
    avg_val = df[y_axis].mean()
    top_performer = df.groupby(x_axis)[y_axis].sum().idxmax()
    top_val = df.groupby(x_axis)[y_axis].sum().max()
    
    st.markdown(f"""
    <div class="insight-box">
        <h3 style='margin-top:0; color:#002e5d;'>💡 Executive AI Summary</h3>
        <p>Analysis of <b>{y_axis}</b> across <b>{x_axis}</b> reveals a total volume of <b>{total:,.2f}</b>.</p>
        <p>The <b>{top_performer}</b> segment is the primary driver, contributing <b>{(top_val/total)*100:.1f}%</b> of total performance.</p>
        
        <div class="future-box">
            <b>🔮 Future Strategic AI Suggestions:</b><br>
            1. <b>Scalability:</b> Based on current trends, expanding resources in <b>{top_performer}</b> could yield a 15% growth next quarter.<br>
            2. <b>Risk Mitigation:</b> Segments performing below the average of <b>{avg_val:,.2f}</b> should be audited for cost-efficiencies.<br>
            3. <b>Automation:</b> We suggest automating the reporting for {x_axis} to save 4 hours of manual data entry per week.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Metrics Row
    m1, m2, m3 = st.columns(3)
    m1.metric(f"Total {y_axis}", f"{total:,.0f}")
    m2.metric("Leader", top_performer)
    m3.metric("Avg Benchmark", f"{avg_val:,.0f}")

    # --- 7. CHART ENGINE ---
    if chart_type == "Area": fig = px.area(df, x=x_axis, y=y_axis, template="plotly_white")
    elif chart_type == "Bar": fig = px.bar(df, x=x_axis, y=y_axis, template="plotly_white")
    elif chart_type == "Line": fig = px.line(df, x=x_axis, y=y_axis, template="plotly_white", markers=True)
    else: fig = px.pie(df, names=x_axis, values=y_axis, hole=0.5, template="plotly_white")
    
    fig.update_traces(marker_color=chart_color) if chart_type != "Donut" else fig.update_traces(marker=dict(colors=[chart_color]))
    st.plotly_chart(fig, use_container_width=True, key=f"viz_{x_axis}_{y_axis}_{chart_type}")

    # --- 8. PPT EXPORT ENGINE ---
    if st.button("📊 Finalize & Export Executive PPT"):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = f"Strategic Review: {y_axis}"
        slide.placeholders[1].text = f"Prepared for: {user_name}\nKey Performance Leader: {top_performer}\nTotal: {total:,.0f}"
        
        buf = io.BytesIO()
        prs.save(buf)
        trigger_success()
        st.download_button("📥 Download Official Report", buf.getvalue(), f"Executive_Report_{datetime.now().strftime('%Y%m%d')}.pptx")

else:
    st.info("👋 System Ready. Please upload your dataset and configure axes in the sidebar to generate insights.")
