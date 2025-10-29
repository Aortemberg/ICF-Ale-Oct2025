# ICFAle.py
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from datetime import datetime
import io
import zipfile
import re
from docx.oxml import OxmlElement

# -----------------------------
# Configuración de la aplicación Streamlit
# -----------------------------
st.set_page_config(page_title="Generador DOCX Consentimientos", layout="wide")

st.title("🩺 Generador automático de Consentimientos (Excel → Word)")

st.markdown("""
Subí tu **modelo.docx** (plantilla con placeholders `<<...>>`) y el **datos.xlsx** con la información de cada investigador. 
El nombre del archivo final se construirá con el Investigador, el Nro. de Centro y la fecha del modelo Word.
""")

# Cargadores de archivos
uploaded_docx = st.file_uploader("📄 Subí el documento modelo (.docx)", type=["docx"])
uploaded_xlsx = st.file_uploader("📊 Subí el Excel (.xlsx)", type=["xlsx"])

# Variables globales para lógica de reemplazo
# Textos para la lógica de provincia
texto_anticonceptivo_original = (
    "El médico del estudio discutirá con usted qué método anticonceptivo se considera adecuado. "
    "El patrocinador y/o el investigador del estudio garantizarán su acceso al método anticonceptivo "
    "acordado y necesario para su participación en este estudio"
)

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
    """Elimina un párrafo completamente del documento."""
    p = paragraph._element
    p.getparent().remove(p)
    # paragraph._p = paragraph._element = None # No es necesario para el uso actual

def replace_text_in_runs(paragraph, old, new):
    """Reemplaza texto en fragmentos de párrafo (runs) sin romper el formato original."""
    for run in paragraph.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)

def replace_text_in_doc(doc, replacements):
    """Aplica reemplazos en todos los párrafos y tablas del documento."""
    
    # Función interna para procesar una lista de párrafos (para reutilizar en tablas)
    def process_paragraphs(paragraphs):
        for p in paragraphs:
            # 1. Intento rápido de reemplazo sin romper runs (mantiene formato)
            for old, new in replacements.items():
                replace_text_in_runs(p, old, new)
            
            # 2. Fallback: Si el texto está dividido en runs (e.g., por formato), 
            #    borra runs y crea uno nuevo con el texto completo reemplazado.
            fulltext = p.text
            for old, new in replacements.items():
                if old in fulltext:
                    # Borrar todos los runs existentes
                    for r in p.runs:
                        r.text = ""
                    # Agregar un run con el texto corregido (perderá el formato intermedio)
                    p.add_run(fulltext.replace(old, new))

    # Proceso en párrafos principales
    process_paragraphs(doc.paragraphs)

    # Proceso en tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_paragraphs(cell.paragraphs)

def find_paragraphs_containing(doc, snippet):
    """Busca y devuelve todos los párrafos que contienen el fragmento de texto dado."""
    res = []
    # Buscar en párrafos principales
    for p in doc.paragraphs:
        if snippet.lower() in p.text.lower():
            res.append(p)
    # Buscar dentro de las tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if snippet.lower() in p.text.lower():
                        res.append(p)
    return res

def get_docx_creation_date(file):
    """Intenta leer la fecha de modificación del modelo Word desde los metadatos."""
    try:
        from zipfile import ZipFile
        from xml.etree import ElementTree as ET
        
        # Volver al inicio del archivo subido
        file.seek(0)
        
        with ZipFile(file) as docx:
            # Archivo de metadatos principal
            core = docx.read("docProps/core.xml")
            tree = ET.fromstring(core)
            # Definición de namespaces para buscar los elementos
            ns = {"dc": "http://purl.org/dc/elements/1.1/", "dcterms": "http://purl.org/dc/terms/"}
            
            # Buscar la fecha de modificación
            modified = tree.find("dcterms:modified", ns)
            if modified is not None and modified.text:
                # Convertir formato ISO a datetime y luego a DD/MM/YYYY
                dt = datetime.fromisoformat(modified.text.replace("Z", "+00:00"))
                return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    # Si falla, usar la fecha actual del sistema
    return datetime.now().strftime("%d/%m/%Y")

def set_global_font_style(doc, font_name="Arial", font_size=11, font_color=RGBColor(0, 0, 0)):
    """Aplica formato de fuente consistente a todos los runs en el documento, incluyendo tablas."""
    font_size_pt = Pt(font_size)

    def apply_style(p):
        for run in p.runs:
            run.font.name = font_name
            run.font.size = font_size_pt
            run.font.color.rgb = font_color

    for p in doc.paragraphs:
        apply_style(p)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    apply_style(p)

# -----------------------------
# Procesamiento de cada fila
# -----------------------------
def process_row_and_generate_doc(template_bytes, row, fecha_modelo):
    # Cargar el documento de plantilla para este ciclo
    doc = Document(io.BytesIO(template_bytes))

    # -----------------------------
    # Mapeo de placeholders <<...>> con columnas del Excel
    # -----------------------------
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

    # -----------------------------
    # Lógica Condicional: Subinvestigador
    # -----------------------------
    subinvestigador_valido = replacements.get("<<SUBINVESTIGADOR>>", "")
    
    if not subinvestigador_valido:
        # Si no hay Subinvestigador, eliminar los placeholders y sus párrafos si existen.
        placeholders_to_remove = ["<<SUBINVESTIGADOR>>", "<<TELEFONO_24HS_SUBINV>>"]
        
        # Eliminar las entradas del diccionario para que no intenten reemplazar con ""
        for key in placeholders_to_remove:
             replacements.pop(key, None) 
        
        # Buscar y eliminar los párrafos que contengan estos placeholders
        for p_key in placeholders_to_remove:
            # Usar la clave como snippet, ya que el texto aún no ha sido reemplazado
            paras = find_paragraphs_containing(doc, p_key)
            for p in paras:
                try:
                    remove_paragraph(p)
                except Exception:
                    pass
    
    # -----------------------------
    # Aplicar todos los reemplazos
    # -----------------------------
    replace_text_in_doc(doc, replacements)

    # -----------------------------
    # Lógica de provincia (después de los reemplazos generales)
    # -----------------------------
    # Se usa row.get("provincia", ...) ya que el dataframe de pandas nos da acceso directo.
    prov = str(row.get("provincia", "")).strip().lower().replace(" ", "")
    
    # 1. Logica Cordoba: Eliminar ambos textos (original y BA)
    if prov == "cordoba":
        # Eliminar texto anticonceptivo original
        paras_orig = find_paragraphs_containing(doc, texto_anticonceptivo_original)
        for p in paras_orig:
            try:
                remove_paragraph(p)
            except Exception:
                pass
        
        # Eliminar referencia a Buenos Aires
        paras_ba_ref = find_paragraphs_containing(doc, "Requerido para centros de la provincia de Buenos Aires")
        for p in paras_ba_ref:
            try:
                remove_paragraph(p)
            except Exception:
                pass
    
    # 2. Logica Buenos Aires: Reemplazar texto anticonceptivo original por el texto BA
    elif prov in ("buenosaires",):
        # Encontrar y reemplazar el texto anticonceptivo original
        paras = find_paragraphs_containing(doc, texto_anticonceptivo_original)
        if paras:
            for p in paras:
                # Reemplazar el contenido del párrafo con el nuevo texto de BA
                for r in p.runs:
                    r.text = ""
                p.add_run(texto_ba_reemplazo)
        else:
            # Fallback: Si no encuentra el texto original, busca la referencia a BA para reemplazarla.
            paras_ba = find_paragraphs_containing(doc, "Requerido para centros de la provincia de Buenos Aires")
            for p in paras_ba:
                for r in p.runs:
                    r.text = ""
                p.add_run(texto_ba_reemplazo)
    
    # -----------------------------
    # Formato y fecha del documento
    # -----------------------------
    
    # Agregar la fecha del modelo al final (para usar en el nombre de archivo)
    # Buscamos el último párrafo y añadimos un separador (opcional)
    # Luego, agregamos la fecha del modelo
    doc.add_paragraph()
    doc.add_paragraph(f"Documento basado en modelo de fecha: {fecha_modelo}")

    # Aplicar formato de fuente a todo el documento
    set_global_font_style(doc)

    out_io = io.BytesIO()
    doc.save(out_io)
    out_io.seek(0)
    return out_io

# -----------------------------
# Ejecución principal
# -----------------------------
if uploaded_docx and uploaded_xlsx:
    
    # Obtener la fecha del modelo primero
    uploaded_docx.seek(0)
    fecha_modelo = get_docx_creation_date(uploaded_docx)

    try:
        # Se asume que no hay filas ocultas y se lee el Excel completo
        # La mejor práctica para evitar el error de índice es no manipular los índices manualmente.
        df = pd.read_excel(uploaded_xlsx, engine="openpyxl")
        
        if df.empty:
            st.error("El archivo Excel está vacío.")
            st.stop()
            
    except Exception as e:
        st.error(f"Error leyendo el Excel: {e}")
        st.stop()

    # Volver a cargar los bytes del modelo .docx después de leer la fecha
    uploaded_docx.seek(0)
    template_bytes = uploaded_docx.read()

    zip_io = io.BytesIO()
    
    with st.spinner('Generando documentos...'):
        with zipfile.ZipFile(zip_io, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for idx, row in df.iterrows():
                try:
                    # Usamos .to_dict() para asegurar compatibilidad con la función de procesamiento
                    doc_io = process_row_and_generate_doc(template_bytes, row.to_dict(), fecha_modelo)
                except Exception as e:
                    # Mostrar error específico para la fila fallida y continuar
                    st.error(f"Error procesando la fila {idx + 2} (registro #{idx + 1}): {e}")
                    continue

                # -----------------------------
                # Construcción del Nombre de Archivo
                # Formato: Investigador - Centro Nº - FechaModelo.docx
                # -----------------------------
                inv = str(row.get("Investigador", "")).strip()
                centro = str(row.get("Nro. de Centro", "")).strip()
                
                # Saneamiento de nombres para evitar caracteres no válidos en archivos
                safe_inv = re.sub(r'[\\/*?:"<>|]', "_", inv)[:100]
                safe_centro = re.sub(r'[\\/*?:"<>|]', "_", centro)[:50]
                safe_fecha = re.sub(r'[\\/]', "-", fecha_modelo)
                
                filename = f"{safe_inv} - Centro {safe_centro} - {safe_fecha}.docx"
                
                # Fallback si las columnas importantes están vacías
                if not safe_inv and not safe_centro:
                    filename = f"documento_generado_{idx + 1}.docx"

                zf.writestr(filename, doc_io.getvalue())

    zip_io.seek(0)
    st.success(f"✅ ¡Documentos generados correctamente! Se crearon {len(df)} archivos.")
    st.download_button(
        "📥 Descargar ZIP", 
        data=zip_io.getvalue(),
        file_name="consentimientos_generados.zip", 
        mime="application/zip"
    )
else:
    st.info("Subí el modelo .docx y el .xlsx para comenzar la generación.")
