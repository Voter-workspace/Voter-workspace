import streamlit as st

def show():

    st.title("📤 Upload PDF")

    st.file_uploader(
        "Upload PDF",
        type=["pdf"]
    )