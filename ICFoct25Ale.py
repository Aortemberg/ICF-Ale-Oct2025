# ICFAle.py
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from datetime import datetime
import io
import zipfile
import re

# -----------------------------
# Configuración de Streamlit
# -----------------------------
st.set_page_config(page_title="Generador DOCX Consentimientos", layout="wide")

st.title("🩺 Generador automático de Consentimientos (Excel → Word)")

st.markdown("""
Subí tu **modelo.docx** (plantilla con placeholders `<<...>>`) y el **datos.xlsx** con la información de cada investigador.  
El nombre del archivo final se construirá con el Investigador, el Nro. de Centro y el Número de Protocolo.
""")

# -----------------------------
# Funciones auxiliares
# -----------------------------
def remove_paragraph(paragraph):
    """Elimina completamente un párrafo del documento."""
    p = paragraph._element
    p.getparent().remove(p)

def replace_text_in_runs(paragraph, old, new, font_name="Arial", font_size=11, font_color=RGBColor(0, 0, 0)):
    """Reemplaza texto dentro de runs sin modificar formato del resto."""
    for run in paragraph.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.color.rgb = font_color

def replace_text_in_doc(doc, replacements):
    """Reemplaza texto en todo el documento, incluyendo tablas."""
    def process_paragraphs(paragraphs):
        for p in paragraphs:
            for old, new in replacements.items():
                if old in p.text:
                    replace_text_in_runs(p, old, new)
            # Fallback si el placeholder está dividido en varios runs
            for old, new in replacements.items():
                if old in p.text:
                    fulltext = p.text
                    for r in p.runs:
                        r.text = ""
                    new_run = p.add_run(fulltext.replace(old, new))
                    new_run.font.name = "Arial"
                    new_run.font.size = Pt(11)
                    new_run.font.color.rgb = RGBColor(0, 0, 0)

    process_paragraphs(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_paragraphs(cell.paragraphs)

def find_paragraphs_containing(doc, snippet):
    """Busca párrafos que contengan un texto determinado."""
    res = []
    for p in doc.paragraphs:
        if snippet.lower() in p.text.lower():
            res.append(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if snippet.lower() in p.text.lower():
                        res.append(p)
    return res

# -----------------------------
# Generación de documento
# -----------------------------
def process_row_and_generate_doc(template_bytes, row):
    """Genera un documento Word a partir de una fila del Excel."""
    doc = Document(io.BytesIO(template_bytes))

    # Función para limpiar valores vacíos o NaN
    def safe_value(val):
        if pd.isna(val) or str(val).strip().lower() in ("nan", "none"):
            return ""
        return str(val).strip()

    replacements = {
        "<<NUMERO_PROTOCOLO>>": safe_value(row.get("Numero de protocolo", "")),
        "<<TITULO_ESTUDIO>>": safe_value(row.get("Titulo del Estudio", "")),
        "<<PATROCINADOR>>": safe_value(row.get("Patrocinador", "")),
        "<<INVESTIGADOR>>": safe_value(row.get("Investigador", "")),
        "<<INSTITUCION>>": safe_value(row.get("Institucion", "")),
        "<<DIRECCION>>": safe_value(row.get("Direccion", "")),
        "<<CARGO_INVESTIGADOR>>": safe_value(row.get("Cargo del Investigador en la Institucion", "")),
        "<<Centro_Nro.>>": safe_value(row.get("Nro. de Centro", "")),
        "<<COMITE>>": safe_value(row.get("COMITE", "")),
        "<<SUBINVESTIGADOR>>": safe_value(row.get("Subinvestigador", "")),
        "<<TELEFONO_24HS>>": safe_value(row.get("TELEFONO 24HS", "")),
        "<<TELEFONO_24HS_SUBINV>>": safe_value(row.get("TELEFONO 24HS subinvestigador", "")),
    }

    # Lógica Subinvestigador vacío → eliminar secciones
    sub_val = replacements.get("<<SUBINVESTIGADOR>>", "")
    if not sub_val:
        for key in ["<<SUBINVESTIGADOR>>", "<<TELEFONO_24HS_SUBINV>>"]:
            replacements.pop(key, None)
            for p in find_paragraphs_containing(doc, key):
                remove_paragraph(p)

    # Reemplazos con formato Arial 11 negro solo en texto nuevo
    replace_text_in_doc(doc, replacements)

    out_io = io.BytesIO()
    doc.save(out_io)
    out_io.seek(0)
    return out_io

# -----------------------------
# Ejecución principal
# -----------------------------
uploaded_docx = st.file_uploader("📄 Subí el documento modelo (.docx)", type=["docx"])
uploaded_xlsx = st.file_uploader("📊 Subí el Excel (.xlsx)", type=["xlsx"])

if uploaded_docx and uploaded_xlsx:
    try:
        # Leemos el Excel manteniendo ceros iniciales
        df = pd.read_excel(uploaded_xlsx, engine="openpyxl", dtype=str)
        if df.empty:
            st.error("⚠️ El archivo Excel está vacío.")
            st.stop()
    except Exception as e:
        st.error(f"Error leyendo el Excel: {e}")
        st.stop()

    uploaded_docx.seek(0)
    template_bytes = uploaded_docx.read()
    zip_io = io.BytesIO()

    with st.spinner("⏳ Generando documentos..."):
        with zipfile.ZipFile(zip_io, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx, row in df.iterrows():
                try:
                    doc_io = process_row_and_generate_doc(template_bytes, row.to_dict())
                except Exception as e:
                    st.error(f"Error procesando fila {idx + 2}: {e}")
                    continue

                inv = str(row.get("Investigador", "")).strip()
                centro = str(row.get("Nro. de Centro", "")).strip()
                protocolo = str(row.get("Numero de protocolo", "")).strip()

                # Limpieza de caracteres no válidos
                safe_inv = re.sub(r'[\\/*?:"<>|]', "_", inv)[:100]
                safe_centro = re.sub(r'[\\/*?:"<>|]', "_", centro)[:50]
                safe_prot = re.sub(r'[\\/*?:"<>|]', "_", protocolo)[:50]

                filename = f"{safe_inv} - Centro {safe_centro} - {safe_prot}.docx"
                if not safe_inv and not safe_centro:
                    filename = f"documento_generado_{idx + 1}.docx"

                zf.writestr(filename, doc_io.getvalue())

    zip_io.seek(0)
    st.success(f"✅ ¡Documentos generados correctamente! ({len(df)} archivos)")
    st.download_button(
        "📥 Descargar ZIP",
        data=zip_io.getvalue(),
        file_name="consentimientos_generados.zip",
        mime="application/zip"
    )
else:
    st.info("👆 Subí el modelo .docx y el .xlsx para comenzar.")
