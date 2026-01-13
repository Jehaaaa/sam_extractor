import streamlit as st
import zipfile
import os
from openpyxl import Workbook
import tempfile
from io import BytesIO

st.set_page_config(page_title="ZIP Manifest Converter", layout="centered")

st.title("📂 ZIP to Excel Converter")
st.write("Upload your `.zip` file to extract `.txt` manifests and convert them to Excel.")

# File Uploader (Changed type to zip)
uploaded_file = st.file_uploader("Choose a ZIP file", type="zip")

if uploaded_file is not None:
    if st.button("Process File"):
        with st.spinner("Extracting and processing..."):
            # Create a temporary directory to handle extraction
            with tempfile.TemporaryDirectory() as temp_dir:
                try:
                    # 1. Open the ZIP file directly from memory
                    with zipfile.ZipFile(uploaded_file) as zf:
                        zf.extractall(temp_dir)

                    # 2. Prepare Excel Workbook
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Data"
                    ws.append(["Prefix", "Content"])

                    files_found = 0
                    
                    # 3. Walk through directories to find .txt files
                    for root, dirs, files in os.walk(temp_dir):
                        for filename in files:
                            # Filter: Ignore Mac system files and non-txt files
                            if filename.endswith(".txt") and not filename.startswith("._") and "__MACOSX" not in root:
                                file_path = os.path.join(root, filename)
                                
                                try:
                                    # Logic: Split prefix and read content
                                    prefix = filename.split("-")[0]
                                    
                                    with open(file_path, "r", encoding="utf-8") as f:
                                        content = f.read().strip()
                                    
                                    ws.append([prefix, content])
                                    files_found += 1
                                except Exception as e:
                                    st.warning(f"Skipped {filename}: {e}")

                    if files_found > 0:
                        # 4. Save Excel to memory buffer
                        excel_buffer = BytesIO()
                        wb.save(excel_buffer)
                        excel_buffer.seek(0)

                        st.success(f"Success! Processed {files_found} text files.")
                        
                        st.download_button(
                            label="📥 Download Output Excel",
                            data=excel_buffer,
                            file_name="output.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("No .txt files found in the archive.")

                except zipfile.BadZipFile:
                    st.error("Error: The file is corrupted or not a valid ZIP.")
                except Exception as e:
                    st.error(f"An unexpected error occurred: {e}")