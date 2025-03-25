#!/usr/bin/env python3

"""
Python Editor, Compiler and Interpreter for the Python Course at "I.I.S. G. Cena di Ivrea".

Author: Francesco Giuseppe Gillio
Date: 25.02.2025
"""

import io
import contextlib
import streamlit as st
from streamlit_ace import st_ace

st.set_page_config(
    page_title="Python IDE", 
    layout="wide"
)
col_img, col_title = st.columns([1, 5])

with col_img:
    st.image(".python/docs/python.png", width=150)

with col_title:
    st.title("Python IDE")

st.markdown("Editor, Compiler and Interpreter to learn Python.")

# --- SIDEBAR: Open File from Disk ---
with st.sidebar:
    st.header("Open File from Disk 💾")
    uploaded_file = st.file_uploader(
        "📤 Open a .py file from Local Disk", 
        type=["py"]
    )
    if uploaded_file is not None:
        file_content = uploaded_file.getvalue().decode("utf-8")
    else:
        file_content = "# TODO\n"

# --- BOX 1: Code Editor ---
with st.container():
    code = st_ace(
        value=file_content,
        language="python",
        theme="monokai",
        keybinding="vscode",
        show_gutter=True,
        tab_size=4,
        wrap=True,
        font_size=14
    )

# --- BOX 2: Operations ---
# Columns: ▶ Run Code and ⬇ Download Code (.py)
col_run, col_download = st.columns(2)

with col_run:
    run_code = st.button("🚀 Run")

with col_download:
    download_code_placeholder = st.empty()

# --- Code Execution ---
if run_code:
    buffer = io.StringIO()
    try:
        with contextlib.redirect_stdout(buffer), contextlib.redirect_stderr(buffer):
            exec(code, {})
    except Exception as e:
        buffer.write(f"❌ Error: {e}")
    output = buffer.getvalue()

    st.markdown("### Output:")
    st.code(output, language="python")
    st.download_button(
        label="📥 Download Output",
        data=output,
        file_name="output.txt",
        mime="text/plain"
    )
    st.text_area("📑 Copy to Clipboard", value=output, height=150)

# --- Download Code Button ---
download_code_placeholder.download_button(
    label="📥 Download Code (.py)",
    data=code,
    file_name="script.py",
    mime="text/x-python"
)
