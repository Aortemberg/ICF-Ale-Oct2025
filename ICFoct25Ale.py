import io
import re
import zipfile
import pandas as pd
import openpyxl
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor

# -----------------------------
# Texto de reemplazo específico
# -----------------------------
texto_ba_reemplazo = (
    "El médico del estudio discutirá con usted qué métodos anticonceptivos se consideran adecuados. "
    "El Patrocinador y/o el médico del estudio garantizará su acceso a este método anticonceptivo "
    "acordado y necesario para su participación en el ensayo. El costo de los métodos anticonceptivos "
    "seleccionados correrá a cargo del Patrocinador."
)

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
    for p in doc.paragraphs:
        for old, new in replacements.items():
            replace_text_in_runs(p, old, new)
        fulltext = p.text
        for old, new in replacements.items():
            if old in fulltext:
                for r in p.runs:
                    r.text = ""
                p.add_run(fulltext.replace(old, new))
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
    """Aplica formato a todos los runs del documento."""
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
        "<<NUMERO_PROTOCOLO>>": str(row.get("Numero de protocolo", "")).strip(),
        "<<TITULO_ESTUDIO>>": str(row.get("Titulo del Estudio", "")).strip(),
        "<<PATROCINADOR>>": str(row.get("Patrocinador", "")).strip(),
        "<<INVESTIGADOR>>": str(row.get("Investigador", "")).strip(),
        "<<INSTITUCION>>": str(row.get("Institucion", "")).strip(),
        "<<DIRECCION>>": str(row.get("Direccion", "")).strip(),
        "<<CARGO_INVESTIGADOR>>": str(row.get("Cargo del Investigador en la Institucion", "")).strip(),
        "<<Centro_Nro.>>": str(row.get("Nro. de Centro", "")).strip(),
        "<<COMITE>>": str(row.get("COMITE", "")).strip(),
        "<<SUBINVESTIGADOR>>": str(row.get("Subinvestigador", "")).strip(),
        "<<TELEFONO_24HS>>": str(row.get("TELEFONO 24HS", "")).strip(),
        "<<TELEFONO_24HS_SUBINV>>": str(row.get("TELEFONO 24HS subinvestigador", "")).strip(),
    }

    # Si Subinvestigador está vacío, eliminar placeholders y párrafos relacionados
    if not replacements["<<SUBINVESTIGADOR>>"]:
        replacements.pop("<<SUBINVESTIGADOR>>", None)
        replacements.pop("<<TELEFONO_24HS_SUBINV>>", None)

        for p in find_paragraphs_containing(doc, "<<SUBINVESTIGADOR>>"):
            remove_paragraph(p)
        for p in find_paragraphs_containing(doc, "<<TELEFONO_24HS_SUBINV>>"):
            remove_paragraph(p)

    # Reemplazar texto
    replace_text_in_doc(doc, replacements)

    # Reglas por provincia
    prov = str(row.get("provincia", "")).strip().lower()
    texto_anticonceptivo_original = "El médico del estudio discutirá con usted qué métodos anticonceptivos"

    if prov == "cordoba":
        paras = find_paragraphs_containing(doc, texto_anticonceptivo_original)
        for p in paras:
            try:
                remove_paragraph(p)
            except Exception:
                pass
        paras_ba = find_paragraphs_containing(doc, "Requerido para centros de la provincia de Buenos Aires")
        for p in paras_ba:
            try:
                remove_paragraph(p)
            except Exception:
                pass

    elif prov.replace(" ", "") in ("buenosaires",):
        paras = find_paragraphs_containing(doc, texto_anticonceptivo_original)
        if paras:
            for p in paras:
                for r in p.runs:
                    r.text = ""
                p.add_run(texto_ba_reemplazo)
        else:
            paras_ba = find_paragraphs_containing(doc, "Requerido para centros de la provincia de Buenos Aires")
            for p in paras_ba:
                for r in p.runs:
                    r.text = ""
                p.add_run(texto_ba_reemplazo)

    # Aplicar formato general Arial 11 negro
    set_font_style(doc)

    # Copiar pie de página del modelo
    copy_footer(template_doc, doc)

    # Guardar en memoria
    out_io = io.BytesIO()
    doc.save(out_io)
    out_io.seek(0)
    return out_io

# -----------------------------
# Ejecución principal
# -----------------------------
if uploaded_docx and uploaded_xlsx:
    try:
        # Detectar filas visibles (filtradas) en el Excel
        wb = openpyxl.load_workbook(uploaded_xlsx, data_only=True)
        sheet = wb.active
        visible_rows = [i for i, row_dim in sheet.row_dimensions.items() if not row_dim.hidden]

        if not visible_rows:
            df = pd.read_excel(uploaded_xlsx, engine="openpyxl")
        else:
            df_all = pd.read_excel(uploaded_xlsx, engine="openpyxl")
            df = df_all.iloc[[i - 2 for i in visible_rows if i > 1]]  # Ajuste por encabezado
    except Exception as e:
        st.error(f"Error leyendo el Excel: {e}")
        st.stop()

    template_bytes = uploaded_docx.read()

    zip_io = io.BytesIO()
    with zipfile.ZipFile(zip_io, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
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
    st.success("✅ Documentos generados correctamente con formato Arial 11 negro y pie de página.")
    st.download_button(
        "📥 Descargar ZIP",
        data=zip_io.getvalue(),
        file_name="consentimientos_generados.zip",
        mime="application/zip"
    )
else:
    st.info("Subí el modelo .docx y el .xlsx para comenzar.")
