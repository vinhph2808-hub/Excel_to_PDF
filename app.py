# -*- coding: utf-8 -*-
"""
Main entry point for Streamlit Community Cloud and Web deployment.
Delegates to excel_to_pdf_streamlit.py.
"""
import runpy
from pathlib import Path

target_file = Path(__file__).resolve().parent / "excel_to_pdf_streamlit.py"
runpy.run_path(str(target_file), run_name="__main__")
