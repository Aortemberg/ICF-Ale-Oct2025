# streamlit_app.py
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
import openpyxl
import io, zipfile, re

# -----------------------------
# Configuración inicial
# -----------------------------
st.set_page_config(page_title="Generador de Consentimientos", layout="wide")
st.title("🩺 Generador automático de Consentimientos Informados")

st.markdown("""
Subí tu **modelo .docx** con placeholders (por ejemplo `{{INVESTIGADOR}}`)  
y el **Excel .xlsx** con los datos.  
Solo se procesarán las **filas visibles** del Excel (las no ocultas).
""")

# -----------------------------
# Carga de archivos
# -----------------------------
uploaded_docx = st.file_uploader("📄 Subí el modelo (.docx)", type=["docx"])
uploaded_xlsx = st.file_uploader("📊 Subí el Excel (.xlsx)", type=["xlsx"])

# -----------------------------
# Funciones auxiliares
# -----------------------------
def remove_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None

def replace_text_in_runs(paragraph, old, new):
    for run in paragraph.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)

def replace_text_in_doc(doc, replacements):
    # Reemplazo en párrafos
    for p in doc.paragraphs:
        for old, new in replacements.items():
            replace_text_in_runs(p, old, new)
        fulltext = p.text
        for old, new in replacements.items():
            if old in fulltext:
                for r in p.runs:
                    r.text = ""
                p.add_run(fulltext.replace(old, new))
    # Reemplazo en tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for old, new in replacements.items():
                        replace_text_in_runs(p, old, new)
                    fulltext = p.text
                    for old, new in replacements.items():
                        if old in fulltext:
                            fo
