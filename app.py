import streamlit as st

# Page configuration
st.set_page_config(
    page_title="Voter Workspace",
    page_icon="🗳️",
    layout="wide"
)

# CSS - Match exact styling and consistent container size to prevent size/layout shifting
st.markdown(
    """
    <style>
    header[data-testid="stHeader"] {
        visibility: hidden !important;
        display: none !important;
    }
    .block-container {
        padding: 0.8rem 1rem !important;
        max-width: 100% !important;
    }
    div.row-widget.stVerticalBlock {
        gap: 0.05rem !important;
    }
    h1, h2, h3, h4 {
        padding-top: 0px !important;
        margin-top: 0px !important;
        margin-bottom: 0.1rem !important;
        font-size: 1.15rem !important;
    }
    p, span, label {
        margin-bottom: 0rem !important;
        margin-top: 0px !important;
    }
    hr {
        margin-top: 0.2rem !important;
        margin-bottom: 0.2rem !important;
    }
    .stButton > button {
        border-radius: 4px;
        font-weight: bold;
        font-size: 1rem !important;
        padding: 6px 12px !important;
        min-height: 40px !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Sidebar Navigation
st.sidebar.title("Voter Workspace")
page = st.sidebar.radio(
    "Navigation", 
    [
       "Dashboard", 
       "New Entry", 
       "PDF Converter", 
       "PDF to Excel", 
       "Reports", 
       "Reports 2",  # <-- Yahan Reports 2 add kar diya gaya hai
       "Search", 
       "Settings", 
       "Upload"
    ]
)

# Page Routing
if page == "Dashboard":
    try:
        from pages.dashboard import show
        show()
    except Exception as e:
        st.error(f"Error loading Dashboard: {e}")

elif page == "New Entry":
    try:
        from pages.new_entry import show
        show()
    except Exception as e:
        st.error(f"Error loading New Entry: {e}")

elif page == "PDF Converter":
    try:
        from pages.pdf_converter import show
        show()
    except Exception as e:
        st.error(f"Error loading PDF Converter: {e}")

elif page == "PDF to Excel":
    try:
        from pages.pdf_to_excel import show
        show()
    except Exception as e:
        st.error(f"Error loading PDF to Excel: {e}")

elif page == "Reports":
    try:
        from pages.reports import show
        show()
    except Exception as e:
        st.error(f"Error loading Reports: {e}")

elif page == "Reports 2":
    try:
        from pages.report2 import show
        show()
    except Exception as e:
        st.error(f"Error loading Reports 2: {e}")

elif page == "Search":
    try:
        from pages.search import show
        show()
    except Exception as e:
        st.error(f"Error loading Search: {e}")

elif page == "Settings":
    try:
        from pages.settings import show
        show()
    except Exception as e:
        st.error(f"Error loading Settings: {e}")

elif page == "Upload":
    try:
        from pages.upload import show
        show()
    except Exception as e:
        st.error(f"Error loading Upload: {e}")