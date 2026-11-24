# -*- coding: utf-8 -*-
"""
Entry point for Streamlit Web App.
Delegates to excel_to_pdf_streamlit.py.
"""
import runpy
from pathlib import Path

target_file = Path(__file__).resolve().parent / "excel_to_pdf_streamlit.py"
runpy.run_path(str(target_file), run_name="__main__")
