import io
import re
import zipfile
import pandas as pd
import openpyxl
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from datetime import datetime
from docx.oxml import OxmlElement

# -----------------------------
# Configuración básica de la app
# -----------------------------
st.set_page_config(page_title="Generador de Consentimientos", layout="centered")
st.title("🩺 Generador automatizado de Consentimientos Informados")

st.write("Subí el modelo (.docx) y el Excel (.xlsx) con los datos filtrados para generar los documentos personalizados.")

# Cargadores de archivos
uploaded_docx = st.file_uploader("📄 Subí el documento modelo (.docx)", type="docx")
uploaded_xlsx = st.file_uploader("📊 Subí el archivo Excel con los datos", type="xlsx")

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
# Funciones auxiliares (Sin cambios en esta sección ya que no afecta el error)
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

def get_docx_creation_date(file):
    """Intenta leer la fecha de creación o modificación del modelo."""
    try:
        from zipfile import ZipFile
        from xml.etree import ElementTree as ET
        with ZipFile(file) as docx:
            core = docx.read("docProps/core.xml")
            tree = ET.fromstring(core)
            ns = {"dc": "http://purl.org/dc/elements/1.1/", "dcterms": "http://purl.org/dc/terms/"}
            modified = tree.find("dcterms:modified", ns)
            if modified is not None and modified.text:
                dt = datetime.fromisoformat(modified.text.replace("Z", "+00:00"))
                return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    # Si no puede leerla, usa la fecha actual
    return datetime.now().strftime("%d/%m/%Y")

# -----------------------------
# Procesamiento de cada fila
# -----------------------------
def process_row_and_generate_doc(template_bytes, row, fecha_modelo):
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

    if not replacements["<<SUBINVESTIGADOR>>"]:
        replacements.pop("<<SUBINVESTIGADOR>>", None)
        replacements.pop("<<TELEFONO_24HS_SUBINV>>", None)
        for p in find_paragraphs_containing(doc, "<<SUBINVESTIGADOR>>"):
            remove_paragraph(p)
        for p in find_paragraphs_containing(doc, "<<TELEFONO_24HS_SUBINV>>"):
            remove_paragraph(p)

    replace_text_in_doc(doc, replacements)

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

    # Formato y pie de página
    set_font_style(doc)
    copy_footer(template_doc, doc)

    # Agregar fecha del modelo al final
    doc.add_paragraph(f"Fecha del documento modelo: {fecha_modelo}")
    set_font_style(doc)

    out_io = io.BytesIO()
    doc.save(out_io)
    out_io.seek(0)
    return out_io

# -----------------------------
# Ejecución principal - Sección Modificada para manejar filas visibles
# -----------------------------
if uploaded_docx is not None and uploaded_xlsx is not None:
    fecha_modelo = get_docx_creation_date(uploaded_docx)

    try:
        wb = openpyxl.load_workbook(uploaded_xlsx, data_only=True)
        sheet = wb.active
        
        # 1. Identificar encabezados (primera fila de datos) y filas visibles.
        # Obtener el encabezado (asumiendo que es la primera fila)
        headers = [cell.value for cell in sheet[1] if cell.value is not None]
        
        # Obtener solo los valores de las filas visibles y que no son la fila de encabezado
        data = []
        for i, row in enumerate(sheet.rows):
            if i == 0: # Saltar la fila del encabezado que ya fue guardada
                continue
            
            # openpyxl usa indices basados en 1, así que row_num es i + 1
            row_num = i + 1
            if not sheet.row_dimensions.get(row_num, openpyxl.worksheet.dimensions.RowDimension()).hidden:
                # Extraer los valores de la fila
                row_values = [cell.value for cell in row]
                if any(row_values): # Asegurarse de que no sea una fila completamente vacía
                    data.append(row_values)
        
        # 2. Crear el DataFrame de pandas solo con las filas visibles
        if not headers:
            st.error("El archivo Excel no tiene encabezados válidos.")
            st.stop()
        
        # Asegurar que los datos y los encabezados tengan el mismo largo
        max_cols = len(headers)
        df = pd.DataFrame([row[:max_cols] for row in data], columns=headers)
        
        if df.empty:
            st.warning("No se encontraron filas de datos visibles y no vacías en el Excel.")
            st.stop()
        
    except Exception as e:
        st.error(f"Error leyendo el Excel: {e}")
        st.stop()

    uploaded_docx.seek(0)
    template_bytes = uploaded_docx.read()

    zip_io = io.BytesIO()
    with zipfile.ZipFile(zip_io, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for idx, row in df.iterrows():
            try:
                # Convierte la fila de pandas a un diccionario para usar en process_row_and_generate_doc
                doc_io = process_row_and_generate_doc(template_bytes, row.to_dict(), fecha_modelo)
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
    st.success(f"✅ Documentos generados correctamente (modelo del {fecha_modelo}).")
    st.download_button(
        "📥 Descargar ZIP",
        data=zip_io.getvalue(),
        file_name="consentimientos_generados.zip",
        mime="application/zip"
    )
else:
    st.info("👆 Subí el modelo .docx y el .xlsx para comenzar.")
