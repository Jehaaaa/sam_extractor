import streamlit as st
import rarfile
import os
import pandas as pd
import tempfile
import shutil

st.title("Manifest & PCID Matcher")

# Upload RAR file
uploaded_file = st.file_uploader("Upload RAR File", type=["rar"])

def get_manifest_id(filename):
    name_without_ext = os.path.splitext(filename)[0]
    if "-" in name_without_ext:
        part_before_hyphen = name_without_ext.split("-")[0]
        if len(part_before_hyphen) >= 6:
            return part_before_hyphen[-6:]
    return None

def get_pcid_id(filename):
    name_without_ext = os.path.splitext(filename)[0]
    if len(name_without_ext) >= 6:
        return name_without_ext[-6:]
    return None

if uploaded_file:

    with tempfile.TemporaryDirectory() as tmp_dir:
        rar_path = os.path.join(tmp_dir, uploaded_file.name)

        # Save uploaded file
        with open(rar_path, "wb") as f:
            f.write(uploaded_file.read())

        # Extract RAR
        try:
            with rarfile.RarFile(rar_path) as rf:
                rf.extractall(tmp_dir)
            st.success("RAR Extracted Successfully!")
        except Exception as e:
            st.error(f"Extraction failed: {e}")
            st.stop()

        # Locate folders
        base_path = os.path.join(tmp_dir, "File Manifest & PCID")
        manifest_path = os.path.join(base_path, "Manifest")
        pcid_path = os.path.join(base_path, "PCID")

        if not os.path.exists(manifest_path) or not os.path.exists(pcid_path):
            st.error("Manifest or PCID folder not found!")
            st.stop()

        # Build PCID map
        pcid_map = {}
        for file in os.listdir(pcid_path):
            if file.lower().endswith(".txt"):
                key = get_pcid_id(file)
                if key:
                    pcid_map[key] = file

        combined_data = []

        # Match with Manifest
        for file in os.listdir(manifest_path):
            if file.lower().endswith(".txt"):
                manifest_key = get_manifest_id(file)
                if manifest_key and manifest_key in pcid_map:

                    manifest_file_path = os.path.join(manifest_path, file)

                    with open(manifest_file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        manifest_content = f.read().strip()

                    combined_data.append({
                        'Match ID': manifest_key,
                        'Manifest Filename': file,
                        'PCID Filename': pcid_map[manifest_key],
                        'Manifest Content': manifest_content,
                        'PCID Content': pcid_map[manifest_key].replace(".txt","")
                    })

        if combined_data:
            df = pd.DataFrame(combined_data)
            st.success(f"{len(df)} matches found!")
            st.dataframe(df)

            output_file = "Combined_Manifest_PCID.xlsx"
            df.to_excel(output_file, index=False)

            with open(output_file, "rb") as f:
                st.download_button(
                    "Download Excel File",
                    f,
                    file_name=output_file
                )
        else:
            st.warning("No matches found.")
