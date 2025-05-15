import streamlit as st
import os
import base64

# Set page config
st.set_page_config(
    page_title="📊 Data Validation Toolkit",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Import after setting page config
import std
import val
import mrg

# Custom CSS
st.markdown(
    """
    <style>
    [data-testid="stSidebar"] {
        background-color: #f8f9fa !important; /* off-white sidebar */
        box-shadow: none !important;
    }

    .sidebar-title {
        font-size: 26px !important;
        color: #28a745 !important;
        font-weight: bold !important;
        margin-bottom: 25px !important;
        text-align: center !important;
    }

    .main-title {
        font-size: 40px !important;
        color: #dc3545 !important;
        text-align: center !important;
        font-weight: bold !important;
        margin-bottom: 15px !important;
    }

    .tagline {
        font-size: 22px !important;
        font-weight: bold !important;
        color: #747f8b !important;
        text-align: center !important;
        margin-bottom: 20px !important;
    }

    .instruction {
        font-size: 16px !important;
        color: #747f8b !important;
        font-weight: bold !important;
        margin-bottom: 10px !important;
        text-align: center !important;
    }

    .tool-list {
        list-style-type: none !important;
        padding: 0 !important;
        margin: 20px auto !important;
        max-width: 600px;
    }

    .tool-list li {
        background-color: rgb(128 128 128 / 10%);
        border-left: 5px solid #007bff !important;
        padding: 15px !important;
        margin-bottom: 10px !important;
        border-radius: 5px !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
    }

    .tool-list b {
        color: #28a745 !important;
    }

    .stButton>button {
        background-color: #28a745 !important;
        color: white !important;
        border: none !important;
        padding: 12px 24px !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }

    .stButton>button:hover {
        background-color: #218838 !important;
    }

    .contact-info {
        text-align: center;
        font-size: 14px;
        color: #6c757d;
        margin-bottom: 5px;
        width: 100%;
    }

    .contact-info a {
        color: #007bff;
        text-decoration: none;
    }

    .contact-info a:hover {
        text-decoration: underline;
    }

    .sidebar-logo {
        max-width: 120px !important;
        height: auto !important;
        margin: 0 auto 10px auto;
        display: block;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Sidebar
st.sidebar.markdown("<p class='sidebar-title'>🛠️ Tools</p>", unsafe_allow_html=True)
selection = st.sidebar.radio(
    label="Explore",
    options=["🏠 Overview", "📐 Standardiser", "📊 Validation Report Generator", "🧩 Excel File Merger"]
)

# Sidebar logo and contact section with padding
image_path = "Sigmoid_Logo.jpg"
if os.path.exists(image_path):
    with open(image_path, "rb") as img_file:
        encoded_img = base64.b64encode(img_file.read()).decode()

    st.sidebar.markdown(f"""
        <div style="padding: 25px 15px; margin-top: 80px; border-radius: 10px; text-align: center;">
            <img src='data:image/png;base64,{encoded_img}' class='sidebar-logo'>
            <p class='contact-info'>📧 <a href='mailto:arkaprova@sigmoidanalytics.com'>Contact Us</a></p>
            <p class='contact-info'>🔗 <a href='https://github.com/sahaa63/validator-' target='_blank'>GitHub Repository</a></p>
        </div>
    """, unsafe_allow_html=True)
else:
    st.sidebar.markdown("""
        <div style="padding: 25px 15px; margin-top: 80px; border-radius: 10px; text-align: center;">
            <p class='contact-info'>📧 <a href='mailto:arkaprova@sigmoidanalytics.com'>Contact Us</a></p>
            <p class='contact-info'>🔗 <a href='https://github.com/sahaa63/validator-' target='_blank'>GitHub Repository</a></p>
        </div>
    """, unsafe_allow_html=True)

# Main area
if selection == "🏠 Overview":
    st.markdown("<h1 class='main-title'>📊 Data Validation Toolkit</h1>", unsafe_allow_html=True)
    st.markdown("<p class='tagline'><b>Your Central Hub for Excel Transformation and Validation</b></p>", unsafe_allow_html=True)
    st.markdown("<p class='instruction'><b>Navigate using the sidebar:</b></p>", unsafe_allow_html=True)
    st.markdown("""
        <ul class='tool-list'>
            <li><b>📐 Standardiser:</b> Effortlessly clean and align your Excel and PBI sheets for consistency.</li>
            <li><b>📊 Validation Report:</b> Quickly compare data sheets and identify key performance indicator differences.</li>
            <li><b>🧩 File Merger:</b> Seamlessly combine up to 10 Excel files into a single, unified document.</li>
        </ul>
    """, unsafe_allow_html=True)

elif selection == "📐 Standardiser":
    std.run()

elif selection == "📊 Validation Report Generator":
    val.run()

elif selection == "🧩 Excel File Merger":
    mrg.run()
