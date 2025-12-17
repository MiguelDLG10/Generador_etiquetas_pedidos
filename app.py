import streamlit as st
import tempfile
import os
from generate_labels import generate_labels_and_summary

# Page Configuration
st.set_page_config(
    page_title="Label Generator",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for Next.js / Vercel Dark Theme Aesthetic
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    /* Global Reset & Theme */
    .stApp {
        background-color: #000000;
        color: #ffffff;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* Hide Streamlit Elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Typography */
    h1 {
        font-weight: 800;
        font-size: 3.5rem;
        background: linear-gradient(180deg, #fff 0%, #888 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        letter-spacing: -0.04em;
        margin-bottom: 0.5rem;
        text-align: center;
    }

    .subtitle {
        color: #888;
        font-size: 1.2rem;
        text-align: center;
        margin-bottom: 3rem;
        font-weight: 400;
    }

    /* Card / Container */
    .css-1y4p8pa {
        padding: 0;
    }
    
    .card {
        background: #111;
        border: 1px solid #333;
        border-radius: 12px;
        padding: 2rem;
        margin: 0 auto;
        max-width: 800px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.5);
    }

    /* File Uploader */
    .stFileUploader {
        padding: 2rem;
        border: 1px dashed #444;
        border-radius: 8px;
        background: #0a0a0a;
        transition: border-color 0.2s;
    }
    
    .stFileUploader:hover {
        border-color: #666;
    }

    /* Buttons */
    .stButton button {
        background-color: #fff;
        color: #000;
        border: none;
        border-radius: 6px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.2s;
        width: 100%;
    }

    .stButton button:hover {
        background-color: #ccc;
        transform: translateY(-1px);
    }

    .stDownloadButton button {
        background-color: #0070f3; /* Vercel Blue */
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.2s;
        width: 100%;
    }

    .stDownloadButton button:hover {
        background-color: #0060df;
        box-shadow: 0 4px 14px 0 rgba(0,118,255,0.39);
        transform: translateY(-1px);
    }

    /* Success/Error Messages */
    .stSuccess, .stError, .stInfo {
        background-color: #111;
        color: #fff;
        border: 1px solid #333;
        border-radius: 8px;
    }

    /* Divider */
    hr {
        border-color: #333;
        margin: 3rem 0;
    }
    
    /* Stats */
    div[data-testid="stMetricValue"] {
        color: #fff;
        font-size: 1.5rem;
    }
    div[data-testid="stMetricLabel"] {
        color: #888;
    }

    </style>
""", unsafe_allow_html=True)

# Layout
col1, col2, col3 = st.columns([1, 2, 1])

with col2:
    st.markdown("<h1>Label Generator</h1>", unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Enterprise-grade PDF generation for your logistics.</p>', unsafe_allow_html=True)

    # Main Card
    st.markdown('<div class="card">', unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader(
        "Upload your Excel file",
        type=['xlsx', 'xls'],
        help="Drag and drop your orders file here"
    )

    if uploaded_file is not None:
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Processing
        with st.spinner('Processing orders...'):
            try:
                # Save uploaded file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_input:
                    tmp_input.write(uploaded_file.getbuffer())
                    tmp_input_path = tmp_input.name
                
                # Output file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_output:
                    tmp_output_path = tmp_output.name
                
                # Generate
                generate_labels_and_summary(tmp_input_path, tmp_output_path)
                
                # Read PDF
                with open(tmp_output_path, 'rb') as pdf_file:
                    pdf_data = pdf_file.read()
                
                # Clean up
                os.unlink(tmp_input_path)
                
                # Success State
                st.success(f"Successfully processed {uploaded_file.name}")
                
                # Stats Row
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("Status", "Ready")
                with c2:
                    st.metric("Format", "PDF")
                with c3:
                    st.metric("Size", f"{len(pdf_data) / 1024:.1f} KB")
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # Download Action
                st.download_button(
                    label="Download Labels",
                    data=pdf_data,
                    file_name="etiquetas_pedidos.pdf",
                    mime="application/pdf"
                )
                
                if os.path.exists(tmp_output_path):
                    os.unlink(tmp_output_path)
                    
            except Exception as e:
                st.error(f"Error: {str(e)}")
    
    st.markdown('</div>', unsafe_allow_html=True)

    # Footer
    st.markdown("""
        <div style="text-align: center; margin-top: 4rem; color: #444; font-size: 0.8rem;">
            POWERED BY ANTIGRAVITY
        </div>
    """, unsafe_allow_html=True)
