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
                            for r in p.runs:
                                r.text = ""
                            p.add_run(fulltext.replace(old, new))

def find_paragraphs_containing(doc, snippet):
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

def set_font_style(doc, font_name="Arial", font_size=11, font_color=RGBColor(0, 0, 0)):
    """Aplica formato Arial 11 negro a todo el documento."""
    for p in doc.paragraphs:
        for run in p.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.color.rgb = font_color
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.name = font_name
                        run.font.size = Pt(font_size)
                        run.font.color.rgb = font_color

def copy_footer(template_doc, target_doc):
    """Copia el pie de página del documento modelo al documento generado."""
    try:
        for section_index, section in enumerate(template_doc.sections):
            footer = section.footer
            target_footer = target_doc.sections[section_index].footer
            for p in list(target_footer.paragraphs):
                remove_paragraph(p)
            for p in footer.paragraphs:
                new_p = target_footer.add_paragraph(p.text)
                for run in new_p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(11)
                    run.font.color.rgb = RGBColor(0, 0, 0)
    except Exception as e:
        print(f"No se pudo copiar el pie de página: {e}")

# -----------------------------
# Procesamiento de cada fila
# -----------------------------
def process_row_and_generate_doc(template_bytes, row):
    template_doc = Document(io.BytesIO(template_bytes))
    doc = Document(io.BytesIO(template_bytes))

    replacements = {
        "{{NUM_PROTOCOLO}}": str(row.get("Numero de protocolo", "")).strip(),
        "{{TITULO_ESTUDIO}}": str(row.get("Titulo del Estudio", "")).strip(),
        "{{PATROCINADOR}}": str(row.get("Patrocinador", "")).strip(),
        "{{INVESTIGADOR}}": str(row.get("Investigador", "")).strip(),
        "{{INSTITUCION}}": str(row.get("Institucion", "")).strip(),
        "{{DIRECCION}}": str(row.get("Direccion", "")).strip(),
        "{{CARGO}}": str(row.get("Cargo", "")).strip(),
        "{{PROVINCIA}}": str(row.get("provincia", "")).strip(),
        "{{COMITE}}": str(row.get("comite", "")).strip(),
        "{{SUBINVESTIGADOR}}": str(row.get("subinvestigador", "")).strip(),
        "{{TELEFONO_24HS}}": str(row.get("TELEFONO_24HS", "")).strip(),
    }

    # Si no hay subinvestigador → eliminar placeholders y párrafos relacionados
    if not replacements["{{SUBINVESTIGADOR}}"]:
        replacements.pop("{{SUBINVESTIGADOR}}", None)
        for p in find_paragraphs_containing(doc, "{{SUBINVESTIGADOR}}"):
            remove_paragraph(p)

    # Reemplazo de texto
    replace_text_in_doc(doc, replacements)

    # Aplicar formato y pie de página
    set_font_style(doc)
    copy_footer(template_doc, doc)

    # Guardar a memoria
    out_io = io.BytesIO()
    doc.save(out_io)
    out_io.seek(0)
    return out_io

# -----------------------------
# Ejecución principal
# -----------------------------
if uploaded_docx and uploaded_xlsx:
    try:
        # Leer el Excel completo
        wb = openpyxl.load_workbook(uploaded_xlsx, data_only=True)
        sheet = wb.active
        df_all = pd.read_excel(uploaded_xlsx, engine="openpyxl")

        # Crear lista de filas visibles (no ocultas)
        visible_rows = []
        for i, row in enumerate(sheet.iter_rows(min_row=2), start=0):
            row_idx = i + 2
            if not sheet.row_dimensions[row_idx].hidden:
                visible_rows.append(i)

        # Si no hay filas ocultas → usar todas
        if len(visible_rows) == 0:
            df = df_all
        else:
            df = df_all.iloc[visible_rows]
    except Exception as e:
        st.error(f"Error leyendo el Excel: {e}")
        st.stop()

    template_bytes = uploaded_docx.read()

    # Crear archivo ZIP
    zip_io = io.BytesIO()
    with zipfile.ZipFile(zip_io, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, row in df.iterrows():
            try:
                doc_io = process_row_and_generate_doc(template_bytes, row)
            except Exception as e:
                st.error(f"Error procesando fila {idx}: {e}")
                continue

            inv = str(row.get("Investigador", "")).strip()
            centro = str(row.get("Nro. de Centro", "")).strip()
            safe_inv = re.sub(r'[\\/*?:"<>|]', "_", inv)[:100]
            safe_centro = re.sub(r'[\\/*?:"<>|]', "_", centro)[:50]
            filename = f"{safe_inv} - Centro {safe_centro}.docx" if safe_inv or safe_centro else f"doc_{idx}.docx"
            zf.writestr(filename, doc_io.getvalue())

    zip_io.seek(0)
    st.success("✅ Documentos generados correctamente en Arial 11 negro.")
    st.download_button(
        "📥 Descargar ZIP",
        data=zip_io.getvalue(),
        file_name="consentimientos_generados.zip",
        mime="application/zip"
    )
else:
    st.info("👆 Subí el modelo .docx y el .xlsx para comenzar.")
