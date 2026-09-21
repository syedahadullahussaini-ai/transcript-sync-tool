import streamlit as st
import tempfile
import os
import zipfile
from sync_format_tool import run

st.set_page_config(page_title="Transcript Sync Tool", layout="centered")

st.title("🎬 Transcript Sync Formatter (Batch Enabled)")
st.markdown("Upload one or multiple DOCX files and get synced formatted outputs.")

# Upload multiple files
uploaded_files = st.file_uploader("Upload DOCX files", type=["docx"], accept_multiple_files=True)

# Inputs
media_name = st.text_input("Media Name Prefix", placeholder="e.g. Conversation_proxy")
offset = st.text_input("Offset (HH:MM:SS:FF)", placeholder="e.g. 11:28:14:15")
fps = st.number_input("FPS", min_value=1, max_value=120, value=25)

speakers = st.text_input(
    "Speakers (optional, comma-separated)",
    placeholder="Crew,Q,Yasaswi ard"
)

show_frames = st.checkbox("Show Frames (HH:MM:SS:FF)", value=True)
merge_bundled = st.checkbox("Merge Bundled Lines", value=True)

if st.button("🚀 Process Files"):
    if not uploaded_files or not media_name or not offset:
        st.error("Please upload files and fill required fields.")
    else:
        try:
            temp_dir = tempfile.mkdtemp()
            output_files = []

            progress_bar = st.progress(0)

            for i, uploaded_file in enumerate(uploaded_files):
                input_path = os.path.join(temp_dir, uploaded_file.name)

                with open(input_path, "wb") as f:
                    f.write(uploaded_file.read())

                output_filename = uploaded_file.name.replace(".docx", "_SYNCED.docx")
                output_path = os.path.join(temp_dir, output_filename)

                # Run your script
                run(
                    input_path=input_path,
                    output_path=output_path,
                    media_name=f"{media_name}_{i+1}",
                    offset_str=offset,
                    fps=fps,
                    speakers=speakers if speakers else None,
                    merge_bundled=merge_bundled,
                    show_frames=show_frames
                )

                output_files.append(output_path)

                progress_bar.progress((i + 1) / len(uploaded_files))

            # Create ZIP
            zip_path = os.path.join(temp_dir, "processed_outputs.zip")
            with zipfile.ZipFile(zip_path, "w") as zipf:
                for file in output_files:
                    zipf.write(file, os.path.basename(file))

            # Download ZIP
            with open(zip_path, "rb") as f:
                st.success(f"✅ {len(output_files)} files processed successfully!")
                st.download_button(
                    label="📦 Download All Outputs (ZIP)",
                    data=f,
                    file_name="synced_outputs.zip",
                    mime="application/zip"
                )

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")