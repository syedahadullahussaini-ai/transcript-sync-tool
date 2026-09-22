import streamlit as st
import tempfile
import os
from sync_format_tool import run, validate_offset
from docx import Document

st.title("Transcript Sync Tool")

uploaded_file = st.file_uploader("Upload DOCX", type=["docx"])

media_name = st.text_input("Media Name")
offset = st.text_input("Offset (HH:MM:SS:FF)")
fps = st.number_input("FPS", value=25)

# ✅ Offset validation warning
if offset and not validate_offset(offset):
    st.warning("⚠️ Offset format should be HH:MM:SS:FF")

if st.button("Process"):
    if not uploaded_file:
        st.error("Please upload a file")
    else:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uploaded_file.read())
            input_path = tmp.name

        output_path = input_path.replace(".docx", "_output.docx")

        result = run(input_path, output_path, media_name, offset, fps)

        # ✅ Show QC errors
        if isinstance(result, dict) and "errors" in result:
            st.error("❌ Errors found:")
            for err in result["errors"]:
                st.write(f"- {err}")
            st.stop()

        st.success("✅ File processed successfully!")

        # ✅ Preview
        doc = Document(output_path)
        preview_text = [para.text for para in doc.paragraphs[:10]]

        st.subheader("📄 Preview")
        st.code("\n".join(preview_text))

        # ✅ Download
        with open(output_path, "rb") as f:
            st.download_button("Download Output", f, file_name="output.docx")
