import streamlit as st
import json

st.set_page_config(page_title="TPP Election Toolkit", layout="wide")

st.title("🗳️ TPP Election Toolkit")

uploaded_file = st.file_uploader("Upload your savefile", type=["json"])

if uploaded_file:
    try:
        # Read and parse the uploaded file
        raw_data = uploaded_file.read()
        data = json.loads(raw_data)

        st.success("Savefile loaded successfully.")
        st.json(data)  # TEMP: Show full JSON structure
    except Exception as e:
        st.error(f"Failed to load file: {e}")
else:
    st.info("Please upload a JSON savefile.")
