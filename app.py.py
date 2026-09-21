import streamlit as st
import tempfile
import os
from sync_format_tool import run

st.title("Transcript Sync Tool")

uploaded_file = st.file_uploader("Upload DOCX", type=["docx"])

media_name = st.text_input("Media Name")
offset = st.text_input("Offset (HH:MM:SS:FF)")
fps = st.number_input("FPS", value=25)

if st.button("Process"):
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uploaded_file.read())
            input_path = tmp.name

        output_path = input_path.replace(".docx", "_output.docx")

        run(input_path, output_path, media_name, offset, fps)

        with open(output_path, "rb") as f:
            st.download_button("Download Output", f, file_name="output.docx")