import hashlib
import os
import json
import urllib.request
import urllib.error
import base64
import sqlite3
import tempfile
import re
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
try:
    from streamlit_quill import st_quill
    QUILL_DISPONIBLE = True
except Exception:
    st_quill = None
    QUILL_DISPONIBLE = False

from docx import Document

# ==================================================
# CONFIGURACIÓN TESSERACT OCR
# ==================================================
TESSERACT_CMD = ""
TESSERACT_TESSDATA = ""

try:
    import pytesseract

    TESSERACT_RUTAS = [
        r"C:\Users\William Tobar\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]

    for ruta_tesseract in TESSERACT_RUTAS:
        if os.path.exists(ruta_tesseract):
            TESSERACT_CMD = ruta_tesseract
            pytesseract.pytesseract.tesseract_cmd = ruta_tesseract

            carpeta_base = os.path.dirname(ruta_tesseract)
            tessdata_path = os.path.join(carpeta_base, "tessdata")

            if os.path.isdir(tessdata_path):
                TESSERACT_TESSDATA = tessdata_path
                # En Windows funciona mejor apuntando directamente a tessdata
                os.environ["TESSDATA_PREFIX"] = tessdata_path

            break
except Exception:
    pass


def idioma_ocr_disponible():
    """
    Detecta idioma OCR disponible.
    Prioridad:
    1. Español (spa)
    2. Inglés (eng)
    3. Sin idioma explícito
    """
    try:
        import pytesseract

        # Si no existe spa.traineddata, NO pedir lang='spa'
        if TESSERACT_TESSDATA:
            spa_file = os.path.join(TESSERACT_TESSDATA, "spa.traineddata")
            eng_file = os.path.join(TESSERACT_TESSDATA, "eng.traineddata")

            if os.path.exists(spa_file):
                return "spa"
            if os.path.exists(eng_file):
                return "eng"

        # Intento alternativo por API
        try:
            idiomas = pytesseract.get_languages(config="")
            if "spa" in idiomas:
                return "spa"
            if "eng" in idiomas:
                return "eng"
        except Exception:
            pass

    except Exception:
        pass

    return None


def image_to_string_seguro(imagen):
    """
    Ejecuta OCR sin caerse si falta spa.traineddata.
    """
    import pytesseract

    lang = idioma_ocr_disponible()

    if lang:
        return pytesseract.image_to_string(imagen, lang=lang)

    # Último recurso: sin lang explícito
    return pytesseract.image_to_string(imagen)


def diagnostico_tesseract_ocr():
    """
    Diagnóstico simple para mostrar estado OCR.
    """
    spa = os.path.join(TESSERACT_TESSDATA, "spa.traineddata") if TESSERACT_TESSDATA else ""
    eng = os.path.join(TESSERACT_TESSDATA, "eng.traineddata") if TESSERACT_TESSDATA else ""
    return {
        "ejecutable": TESSERACT_CMD or "No detectado",
        "tessdata": TESSERACT_TESSDATA or "No detectada",
        "spa": "Disponible" if spa and os.path.exists(spa) else "No disponible",
        "eng": "Disponible" if eng and os.path.exists(eng) else "No disponible",
        "idioma_usado": idioma_ocr_disponible() or "Automático / sin idioma explícito",
    }




try:
    from PIL import Image
except Exception:
    Image = None


# ==================================================
# APP CONVIVENCIA PUKARAY - FASE 1
# MÓDULO HOGAR / APODERADO ESTABLE
# ==================================================

APP_DIR = Path(__file__).parent
DATOS_DIR = APP_DIR / "datos"
PLANTILLAS_DIR = APP_DIR / "plantillas"
EXPORTADOS_DIR = APP_DIR / "exportados"
ASSETS_DIR = APP_DIR / "assets"

for carpeta in [DATOS_DIR, PLANTILLAS_DIR, EXPORTADOS_DIR, ASSETS_DIR]:
    carpeta.mkdir(exist_ok=True)

DB_PATH = APP_DIR / "pukaray_entrevistas.db"

BASE_EXCEL = DATOS_DIR / "base_datos_pukaray.xlsx"
# Desde esta versión se usa UNA sola base definitiva: base_datos_pukaray.xlsx
TIPO_MOTIVOS_EXCEL = BASE_EXCEL
REGISTRO_EXCEL = DATOS_DIR / "registro_entrevistas.xlsx"
PLANTILLA_APODERADO = PLANTILLAS_DIR / "Ficha_Entrevista_APODERADO.docx"
PLANTILLA_ESTUDIANTE = PLANTILLAS_DIR / "Ficha_Entrevista_ESTUDIANTE.docx"
NOTEBOOKLM_URL = "https://notebooklm.google.com/notebook/95aeeaf6-1f2a-4b4e-b7b1-b9d2ab3e10d9"
def buscar_logo_institucional():
    """
    Busca automáticamente logo institucional en carpeta assets.
    Nombres recomendados:
    - logo_pukaray.png
    - logo.png
    - insignia.png
    - escudo.png
    También acepta jpg/jpeg/webp.
    """
    candidatos = [
        "logo_pukaray.png",
        "logo.png",
        "insignia.png",
        "escudo.png",
        "pukaray.png",
        "Logo_Pukaray.png",
        "LOGO.png",
    ]

    for nombre in candidatos:
        ruta = ASSETS_DIR / nombre
        if ruta.exists():
            return ruta

    for ext in ["*.png", "*.jpg", "*.jpeg", "*.webp"]:
        encontrados = list(ASSETS_DIR.glob(ext))
        if encontrados:
            return encontrados[0]

    return ASSETS_DIR / "logo_pukaray.png"


LOGO_PATH = buscar_logo_institucional()


def logo_html_base64(width=96):
    """
    Logo embebido en HTML para evitar recortes de st.image.
    """
    try:
        if LOGO_PATH.exists():
            ext = LOGO_PATH.suffix.lower().replace(".", "")
            if ext == "jpg":
                ext = "jpeg"
            data = base64.b64encode(LOGO_PATH.read_bytes()).decode("utf-8")
            return f'<img src="data:image/{ext};base64,{data}" style="width:{width}px; height:auto; object-fit:contain; display:block;">'
    except Exception:
        pass
    return '<div style="font-size:60px;">🌳</div>'



# ==================================================
# CONFIGURACIÓN VISUAL
# ==================================================

st.set_page_config(
    page_title="App Convivencia Pukaray - Fase 1",
    page_icon="🌳",
    layout="wide"
)

COLOR_CREMA = "#f5f3eb"
COLOR_VERDE_OSCURO = "#1a542a"
COLOR_VERDE_MEDIO = "#34a853"
COLOR_BURDEO = "#6b1e11"
COLOR_ROJO = "#aa3522"
COLOR_NEGRO_VERDE = "#0d2a15"

st.markdown(f"""
<style>
/* =====================================================
   FIX VISUAL LISTAS DESPLEGABLES LARGAS
===================================================== */

div[data-baseweb="select"] {{
    width: 100% !important;
    min-width: 100% !important;
}}

div[data-baseweb="select"] > div {{
    min-height: 46px !important;
    width: 100% !important;
}}

div[data-baseweb="popover"] {{
    min-width: 1280px !important;
    max-width: 98vw !important;
}}

div[role="listbox"] {{
    min-width: 1280px !important;
    max-width: 98vw !important;
}}

div[role="option"] {{
    white-space: normal !important;
    line-height: 1.35 !important;
    min-height: 42px !important;
    padding-top: 8px !important;
    padding-bottom: 8px !important;
}}

div[data-baseweb="select"] span {{
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    line-height: 1.3 !important;
}}

.stApp {{
    background:
        radial-gradient(circle at top left, rgba(52,168,83,0.15), transparent 28%),
        linear-gradient(180deg, {COLOR_CREMA} 0%, #ffffff 55%, #f5f3eb 100%);
    color: {COLOR_NEGRO_VERDE};
}}

.block-container {{
    padding-top: 1rem;
    max-width: 1680px;
}}

h1, h2, h3 {{
    color: {COLOR_VERDE_OSCURO};
    font-weight: 850;
}}

section[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, {COLOR_NEGRO_VERDE} 0%, {COLOR_VERDE_OSCURO} 60%, {COLOR_BURDEO} 100%);
    border-right: 5px solid {COLOR_ROJO};
}}

section[data-testid="stSidebar"] * {{
    color: white !important;
}}

.pukaray-hero {{
    width: 100%;
    min-height: 190px;
    background: linear-gradient(115deg, {COLOR_VERDE_OSCURO} 0%, {COLOR_NEGRO_VERDE} 62%, {COLOR_BURDEO} 100%);
    border-radius: 26px;
    border-left: 10px solid {COLOR_ROJO};
    box-shadow: 0 14px 34px rgba(13,42,21,0.24);
    padding: 28px 34px;
    margin: 8px 0 26px 0;
    display: flex;
    align-items: center;
    gap: 28px;
    overflow: hidden;
    position: relative;
}}

.pukaray-logo-box {{
    width: 138px;
    height: 138px;
    min-width: 138px;
    background: rgba(255,255,255,0.94);
    border-radius: 26px;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 14px;
    box-shadow: 0 10px 26px rgba(0,0,0,0.28);
    position: relative;
    z-index: 2;
}}

.pukaray-logo-box img {{
    max-width: 108px;
    max-height: 108px;
    object-fit: contain;
}}

.pukaray-hero-text {{
    position: relative;
    z-index: 2;
}}

.pukaray-small-title {{
    color: #f5f3eb;
    font-size: 1.15rem;
    font-weight: 750;
    margin-bottom: 10px;
}}

.pukaray-main-title {{
    color: white;
    font-size: 2.35rem;
    font-weight: 900;
    line-height: 1.1;
    letter-spacing: -0.03em;
    text-shadow: 0 3px 10px rgba(0,0,0,0.26);
}}

.pukaray-subtitle {{
    color: #f5f3eb;
    font-size: 1.08rem;
    margin-top: 18px;
}}

.pukaray-watermark {{
    position: absolute;
    right: 36px;
    top: 22px;
    font-size: 8rem;
    opacity: 0.10;
    transform: rotate(-12deg);
}}

.pukaray-card {{
    background: rgba(255,255,255,0.94);
    border-radius: 22px;
    padding: 24px;
    box-shadow: 0 10px 26px rgba(26,84,42,0.12);
    border-left: 7px solid {COLOR_ROJO};
    border: 1px solid rgba(26,84,42,0.12);
}}

.pukaray-badge {{
    display: inline-block;
    background: linear-gradient(90deg, {COLOR_ROJO}, {COLOR_BURDEO});
    color: white;
    padding: 6px 14px;
    border-radius: 999px;
    font-size: 0.85rem;
    font-weight: 850;
    margin-bottom: 10px;
}}

div[data-testid="stMetric"] {{
    background: rgba(255,255,255,0.96);
    border-radius: 18px;
    padding: 16px;
    border-left: 6px solid {COLOR_ROJO};
    box-shadow: 0 8px 18px rgba(26,84,42,0.10);
}}

.stButton button, .stDownloadButton button {{
    background: linear-gradient(90deg, {COLOR_VERDE_OSCURO}, {COLOR_VERDE_MEDIO});
    color: white;
    border-radius: 14px;
    border: none;
    font-weight: 800;
    min-height: 42px;
    box-shadow: 0 7px 16px rgba(26,84,42,0.20);
}}

.stButton button:hover, .stDownloadButton button:hover {{
    background: linear-gradient(90deg, {COLOR_ROJO}, {COLOR_BURDEO});
    color: white;
}}

div[data-baseweb="select"] > div, textarea, input {{
    border-radius: 12px !important;
    border: 1.5px solid rgba(26,84,42,0.20) !important;
    background: rgba(255,255,255,0.98) !important;
}}

label {{
    font-weight: 750 !important;
    color: {COLOR_VERDE_OSCURO} !important;
}}

section[data-testid="stSidebar"] .pukaray-sidebar-logo {{
    text-align: center;
    margin: 14px 0 18px 0;
}}

section[data-testid="stSidebar"] .pukaray-sidebar-logo-inner {{
    background: rgba(255,255,255,0.94);
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 104px;
    height: 104px;
    border-radius: 24px;
    box-shadow: 0 8px 20px rgba(0,0,0,0.25);
}}

section[data-testid="stSidebar"] .pukaray-sidebar-logo-inner img {{
    max-width: 82px;
    max-height: 82px;
    object-fit: contain;
}}

[data-testid="stDataFrame"] {{
    border-radius: 16px;
    overflow: hidden;
    box-shadow: 0 8px 22px rgba(26,84,42,0.10);
}}




/* =====================================================
   FIX DEFINITIVO STREAMLIT-QUILL - EDITORES DE TEXTO
   Mantiene visible toolbar y cuerpo del editor.
===================================================== */

/* Contenedor general de componentes Streamlit */
div[data-testid="stIFrame"],
iframe {{
    min-height: 285px !important;
    height: 285px !important;
    visibility: visible !important;
    opacity: 1 !important;
    display: block !important;
    border: none !important;
}}

/* Iframe específico de componentes Quill cuando Streamlit lo etiqueta */
iframe[title*="streamlit_quill"],
iframe[title*="st_quill"],
iframe[title*="quill"] {{
    min-height: 285px !important;
    height: 285px !important;
    display: block !important;
    visibility: visible !important;
    opacity: 1 !important;
}}

/* Refuerzo visual si Quill se renderiza fuera del iframe en alguna versión */
.ql-toolbar,
.ql-toolbar.ql-snow {{
    display: block !important;
    visibility: visible !important;
    opacity: 1 !important;
    position: relative !important;
    z-index: 9999 !important;
    background: #ffffff !important;
    border-radius: 12px 12px 0 0 !important;
    border: 1px solid #d0d7de !important;
}}

.ql-container,
.ql-container.ql-snow {{
    display: block !important;
    visibility: visible !important;
    min-height: 190px !important;
    border-radius: 0 0 12px 12px !important;
    border: 1px solid #d0d7de !important;
    background: #ffffff !important;
}}

.ql-editor {{
    min-height: 170px !important;
    font-size: 15px !important;
    line-height: 1.55 !important;
}}




</style>
""", unsafe_allow_html=True)




def boton_notebooklm_normativo():
    """
    Acceso directo a NotebookLM normativo.
    No envía datos sensibles automáticamente.
    """
    st.link_button("📚 Abrir NotebookLM normativo", NOTEBOOKLM_URL)

def encabezado(titulo: str, subtitulo: str = "Formando en Respeto y Responsabilidad"):
    logo = logo_html_base64(width=104)

    st.markdown(
        f"""
        <div class="pukaray-hero">
            <div class="pukaray-logo-box">
                {logo}
            </div>
            <div class="pukaray-hero-text">
                <div class="pukaray-small-title">Colegio Pukaray</div>
                <div class="pukaray-main-title">{titulo}</div>
                <div class="pukaray-subtitle">{subtitulo}</div>
            </div>
            <div class="pukaray-watermark">🌿</div>
        </div>
        """,
        unsafe_allow_html=True
    )



# ==================================================
# LEVANTAMIENTO DE INFORMACIÓN DESDE ESCANEOS
# ==================================================

def ocr_disponible():
    """
    Verifica si pytesseract está disponible.
    Requiere instalar en el PC servidor:
    pip install pytesseract pillow pymupdf
    Además, instalar Tesseract OCR para Windows.
    """
    try:
        import pytesseract  # noqa
        return True
    except Exception:
        return False


def leer_texto_imagen_ocr(archivo_bytes, nombre_archivo="archivo"):
    """
    Lee texto desde imagen usando OCR local.
    """
    if Image is None:
        return "", "Falta instalar Pillow: pip install pillow"

    try:
        import pytesseract
    except Exception:
        return "", "Falta instalar pytesseract: pip install pytesseract"

    try:
        image = Image.open(BytesIO(archivo_bytes))
        texto = image_to_string_seguro(image)
        return texto.strip(), ""
    except Exception as e:
        return "", f"No fue posible leer la imagen con OCR: {e}. Verifica instalación de Tesseract OCR."


def leer_texto_pdf_ocr(archivo_bytes):
    """
    Intenta leer PDF escaneado usando PyMuPDF + pytesseract.
    """
    try:
        import fitz  # PyMuPDF
    except Exception:
        return "", "Para leer PDF escaneado instala PyMuPDF: pip install pymupdf"

    try:
        import pytesseract
    except Exception:
        return "", "Falta instalar pytesseract: pip install pytesseract"

    if Image is None:
        return "", "Falta instalar Pillow: pip install pillow"

    textos = []
    try:
        doc = fitz.open(stream=archivo_bytes, filetype="pdf")
        for i, page in enumerate(doc):
            # Render a resolución moderada
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            txt = image_to_string_seguro(img)
            if txt.strip():
                textos.append(txt.strip())
        return "\n\n".join(textos).strip(), ""
    except Exception as e:
        return "", f"No fue posible leer el PDF escaneado: {e}. Verifica instalación de Tesseract OCR."


def limpiar_texto_ocr(texto):
    """
    Limpia texto proveniente de OCR.
    Blindado: importa re localmente para evitar NameError.
    """
    import re as _re

    texto = str(texto or "")
    texto = texto.replace("\r", "\n")
    texto = _re.sub(r"[ \t]+", " ", texto)
    texto = _re.sub(r"\n{3,}", "\n\n", texto)
    texto = texto.replace("MOTIVO DELA", "MOTIVO DE LA")
    texto = texto.replace("ACUERDOS Y CONCLUSION ES", "ACUERDOS Y CONCLUSIONES")
    texto = texto.replace("CONCLUCIONES", "CONCLUSIONES")
    texto = texto.replace("CONCLUCION", "CONCLUSIÓN")
    return texto.strip()


def extraer_bloque_por_titulos(texto, titulos_inicio, titulos_fin=None):
    """
    Extrae un bloque buscando títulos aproximados.
    """
    if titulos_fin is None:
        titulos_fin = []

    original = str(texto or "")
    mayus = original.upper()

    inicio = -1
    titulo_encontrado = ""
    for titulo in titulos_inicio:
        pos = mayus.find(titulo.upper())
        if pos != -1 and (inicio == -1 or pos < inicio):
            inicio = pos
            titulo_encontrado = titulo

    if inicio == -1:
        return ""

    inicio_contenido = inicio + len(titulo_encontrado)

    # Saltar separadores comunes
    while inicio_contenido < len(original) and original[inicio_contenido] in [":", "-", "\n", " "]:
        inicio_contenido += 1

    fin = len(original)
    for titulo in titulos_fin:
        pos = mayus.find(titulo.upper(), inicio_contenido)
        if pos != -1 and pos < fin:
            fin = pos

    return original[inicio_contenido:fin].strip()


def separar_motivo_acuerdos_desde_ocr(texto):
    """
    Separa MOTIVO DE LA ENTREVISTA y ACUERDOS / CONCLUSIONES desde texto OCR.
    Si no detecta títulos, entrega propuesta conservadora.
    """
    texto = limpiar_texto_ocr(texto)

    titulos_motivo = [
        "MOTIVO DE LA ENTREVISTA",
        "MOTIVO ENTREVISTA",
        "MOTIVO",
        "ANTECEDENTES",
    ]

    titulos_acuerdos = [
        "ACUERDOS Y CONCLUSIONES",
        "ACUERDOS / CONCLUSIONES",
        "ACUERDOS",
        "CONCLUSIONES",
        "ACUERDO",
        "CONCLUSIÓN",
        "CONCLUSION",
    ]

    otros_titulos = [
        "OBSERVACIONES",
        "FIRMA",
        "NOMBRE",
        "REGISTRO",
        "ENTREVISTADOR",
        "APODERADO",
        "ESTUDIANTE",
    ]

    motivo = extraer_bloque_por_titulos(
        texto,
        titulos_motivo,
        titulos_acuerdos + otros_titulos
    )

    acuerdos = extraer_bloque_por_titulos(
        texto,
        titulos_acuerdos,
        otros_titulos
    )

    # Si el OCR no encuentra títulos, usar separación aproximada por mitad solo como respaldo.
    if not motivo and not acuerdos:
        lineas = [l.strip() for l in texto.splitlines() if l.strip()]
        if len(lineas) >= 4:
            mitad = max(1, len(lineas) // 2)
            motivo = "\n".join(lineas[:mitad])
            acuerdos = "\n".join(lineas[mitad:])
        else:
            motivo = texto
            acuerdos = ""

    motivo = mejorar_redaccion_conservadora(motivo, contexto="antecedentes_hp", tarea="antecedentes", sensible=True) if motivo else ""
    acuerdos = mejorar_redaccion_conservadora(acuerdos, contexto="hp", tarea="acuerdos", sensible=True) if acuerdos else ""

    return motivo, acuerdos, texto


def modulo_levantamiento_informacion():
    """
    Módulo para subir escaneos y separar motivo/acuerdos.
    """
    encabezado("Levantamiento de información", "Transcripción inicial desde documentos escaneados")

    st.markdown("### Subir archivo escaneado")
    st.caption("Formatos recomendados: PNG, JPG, JPEG o PDF escaneado. OCR para texto impreso; para manuscrita use transcripción manual asistida.")

    with st.expander("Requisitos del PC servidor", expanded=False):
        st.markdown("""
        Para leer escaneos se requiere instalar:
        ```powershell
        pip install pytesseract pillow pymupdf
        ```
        Además, instalar **Tesseract OCR para Windows**. La app busca automáticamente tu instalación local.

        Si el archivo es una foto, se recomienda que esté derecha, clara y con buena iluminación.
        """)

        st.markdown("**Diagnóstico OCR actual:**")
        st.write(f"Tesseract ejecutable: `{TESSERACT_CMD or 'No detectado'}`")
        st.write(f"Carpeta tessdata: `{TESSERACT_TESSDATA or 'No detectada'}`")
        st.write(f"Idioma OCR usado: `{idioma_ocr_disponible() or 'Automático / sin idioma explícito'}`")

    archivo = st.file_uploader(
        "Cargar archivo escaneado",
        type=["png", "jpg", "jpeg", "pdf"],
        key="levantamiento_archivo"
    )

    if not archivo:
        st.info("Sube un archivo escaneado para iniciar el levantamiento.")
        return

    bytes_archivo = archivo.read()
    nombre = archivo.name.lower()

    if st.button("🔎 Leer escaneo y levantar información", key="btn_levantar_info_ocr"):
        with st.spinner("Leyendo escaneo con OCR local..."):
            if nombre.endswith(".pdf"):
                texto_ocr, error = leer_texto_pdf_ocr(bytes_archivo)
            else:
                texto_ocr, error = leer_texto_imagen_ocr(bytes_archivo, archivo.name)

            if error:
                st.error(error)
                st.stop()

            if not texto_ocr.strip():
                st.warning("No se detectó texto legible. Intenta con una imagen más clara o escaneo de mejor resolución.")
                st.stop()

            motivo, acuerdos, texto_limpio = separar_motivo_acuerdos_desde_ocr(texto_ocr)

            st.session_state["levantamiento_texto_ocr"] = texto_limpio
            st.session_state["levantamiento_motivo"] = motivo
            st.session_state["levantamiento_acuerdos"] = acuerdos

    if st.session_state.get("levantamiento_texto_ocr"):
        st.markdown("### Texto OCR detectado")
        st.text_area(
            "Transcripción completa",
            value=st.session_state["levantamiento_texto_ocr"],
            height=220,
            key="levantamiento_texto_completo"
        )

        col1, col2 = st.columns(2)

        with col1:
            motivo_editado = editor_texto_word(
                "Motivo de la entrevista",
                key="levantamiento_motivo_editor",
                value=st.session_state.get("levantamiento_motivo", ""),
                height=260,
                help_text="Texto separado automáticamente. Revísalo antes de usarlo."
            )

        with col2:
            acuerdos_editados = editor_texto_word(
                "Acuerdos y conclusiones",
                key="levantamiento_acuerdos_editor",
                value=st.session_state.get("levantamiento_acuerdos", ""),
                height=260,
                help_text="Texto separado automáticamente. Revísalo antes de usarlo."
            )

        st.session_state["levantamiento_motivo"] = motivo_editado
        st.session_state["levantamiento_acuerdos"] = acuerdos_editados

        st.markdown("### Enviar al formulario")
        st.caption("Esto deja el texto preparado en la sesión para copiarlo al formulario hp/e. No guarda automáticamente.")

        c1, c2 = st.columns(2)

        with c1:
            if st.button("Enviar a Hogar/Apoderado", key="btn_enviar_ocr_hp"):
                st.session_state["acuerdos_hogar_mejorado"] = acuerdos_editados
                st.session_state["antecedentes_ordenados_hogar"] = motivo_editado
                st.success("Texto preparado para Hogar/Apoderado. Entra al formulario hp y revisa antecedentes/acuerdos.")

        with c2:
            if st.button("Enviar a Entrevista Estudiante", key="btn_enviar_ocr_e"):
                st.session_state["acuerdos_estudiante_mejorado"] = acuerdos_editados
                st.session_state["antecedentes_ordenados_estudiante"] = motivo_editado
                st.success("Texto preparado para Entrevista Estudiante. Entra al formulario e y revisa antecedentes/acuerdos.")




def previsualizar_archivo_escaneado(archivo, bytes_archivo):
    """
    Muestra una previsualización básica del archivo subido.
    Para PDF muestra aviso y permite usar OCR; para imágenes muestra la imagen.
    """
    nombre = archivo.name.lower()

    if nombre.endswith((".png", ".jpg", ".jpeg")):
        try:
            st.image(bytes_archivo, caption="Vista del escaneo cargado", use_container_width=True)
        except Exception:
            st.info("Archivo de imagen cargado. No fue posible previsualizarlo, pero puede intentar OCR o transcripción manual.")
    elif nombre.endswith(".pdf"):
        st.info("PDF cargado. Si es manuscrito, se recomienda abrirlo en otra ventana o visualizarlo desde el archivo original mientras transcribe manualmente.")
    else:
        st.info("Archivo cargado.")


def bloque_transcripcion_manual_asistida(prefix, destino="hp"):
    """
    Alternativa confiable para letra manuscrita:
    el usuario lee el escaneo y transcribe manualmente Motivo / Acuerdos.
    Luego la app mejora redacción de forma conservadora.
    """
    st.markdown("#### ✍️ Transcripción manual asistida")
    st.caption("Recomendado para letra manuscrita. La IA no inventa: solo mejora redacción, ortografía y puntuación.")

    motivo_manual = editor_texto_word(
        "Motivo de la entrevista transcrito",
        key=f"{prefix}_motivo_manual_editor",
        value=st.session_state.get(f"{prefix}_motivo_manual_mejorado", ""),
        height=240,
        help_text="Transcriba aquí el motivo tal como aparece en el documento."
    )

    acuerdos_manual = editor_texto_word(
        "Acuerdos y conclusiones transcritos",
        key=f"{prefix}_acuerdos_manual_editor",
        value=st.session_state.get(f"{prefix}_acuerdos_manual_mejorado", ""),
        height=260,
        help_text="Transcriba cada acuerdo en una línea o viñeta independiente."
    )

    c1, c2 = st.columns(2)

    with c1:
        if st.button("✏️ Mejorar motivo transcrito", key=f"{prefix}_btn_mejorar_motivo_manual"):
            st.session_state[f"{prefix}_motivo_manual_mejorado"] = mejorar_texto_mixto(
                motivo_manual,
                contexto="antecedentes_hp" if destino == "hp" else "antecedentes_e",
                tarea="antecedentes",
                sensible=True
            )
            st.rerun()

    with c2:
        if st.button("✏️ Mejorar acuerdos transcritos", key=f"{prefix}_btn_mejorar_acuerdos_manual"):
            st.session_state[f"{prefix}_acuerdos_manual_mejorado"] = mejorar_texto_mixto(
                acuerdos_manual,
                contexto="hp" if destino == "hp" else "e",
                tarea="acuerdos",
                sensible=True
            )
            st.rerun()

    if st.button("Enviar transcripción manual al formulario", key=f"{prefix}_btn_enviar_manual"):
        motivo_final = st.session_state.get(f"{prefix}_motivo_manual_mejorado", motivo_manual)
        acuerdos_final = st.session_state.get(f"{prefix}_acuerdos_manual_mejorado", acuerdos_manual)

        if destino == "hp":
            st.session_state["antecedentes_ordenados_hogar"] = motivo_final
            st.session_state["acuerdos_hogar_mejorado"] = acuerdos_final
        else:
            st.session_state["antecedentes_ordenados_estudiante"] = motivo_final
            st.session_state["acuerdos_estudiante_mejorado"] = acuerdos_final

        st.success("Transcripción enviada al formulario. Revise antecedentes y acuerdos antes de guardar.")


def bloque_levantamiento_ocr_formulario(prefix, destino="hp"):
    """
    Bloque compacto para formularios HP y E.
    Incluye:
    - OCR automático para texto impreso.
    - Transcripción manual asistida para letra manuscrita.
    """
    with st.expander("📄 Levantamiento de información desde escaneo", expanded=False):
        st.caption("OCR recomendado para texto impreso. Para letra manuscrita, use transcripción manual asistida.")

        archivo = st.file_uploader(
            "Cargar escaneo",
            type=["png", "jpg", "jpeg", "pdf"],
            key=f"{prefix}_ocr_archivo"
        )

        if archivo:
            bytes_archivo = archivo.read()
            nombre = archivo.name.lower()

            previsualizar_archivo_escaneado(archivo, bytes_archivo)

            st.markdown("#### 🔎 OCR automático")
            st.caption("Úselo principalmente si el documento tiene texto impreso o letra muy clara.")

            if st.button("Intentar lectura OCR", key=f"{prefix}_btn_ocr_leer"):
                with st.spinner("Leyendo escaneo..."):
                    if nombre.endswith(".pdf"):
                        texto_ocr, error = leer_texto_pdf_ocr(bytes_archivo)
                    else:
                        texto_ocr, error = leer_texto_imagen_ocr(bytes_archivo, archivo.name)

                if error:
                    st.warning(error)
                    st.info("Si el documento es manuscrito, use la transcripción manual asistida más abajo.")
                elif not texto_ocr.strip():
                    st.warning("No se detectó texto legible. Para manuscrita, use la transcripción manual asistida.")
                else:
                    motivo, acuerdos, texto_limpio = separar_motivo_acuerdos_desde_ocr(texto_ocr)

                    st.session_state[f"{prefix}_ocr_texto"] = texto_limpio
                    st.session_state[f"{prefix}_ocr_motivo"] = motivo
                    st.session_state[f"{prefix}_ocr_acuerdos"] = acuerdos

                    if destino == "hp":
                        st.session_state["antecedentes_ordenados_hogar"] = motivo
                        st.session_state["acuerdos_hogar_mejorado"] = acuerdos
                    else:
                        st.session_state["antecedentes_ordenados_estudiante"] = motivo
                        st.session_state["acuerdos_estudiante_mejorado"] = acuerdos

                    st.success("Texto OCR preparado para el formulario. Revise antes de guardar.")

            if st.session_state.get(f"{prefix}_ocr_texto"):
                st.markdown("**Texto detectado por OCR**")
                st.text_area(
                    "Transcripción OCR completa",
                    value=st.session_state.get(f"{prefix}_ocr_texto", ""),
                    height=160,
                    key=f"{prefix}_ocr_texto_visible"
                )

                c1, c2 = st.columns(2)
                with c1:
                    st.text_area(
                        "Motivo detectado",
                        value=st.session_state.get(f"{prefix}_ocr_motivo", ""),
                        height=140,
                        key=f"{prefix}_ocr_motivo_visible"
                    )
                with c2:
                    st.text_area(
                        "Acuerdos detectados",
                        value=st.session_state.get(f"{prefix}_ocr_acuerdos", ""),
                        height=140,
                        key=f"{prefix}_ocr_acuerdos_visible"
                    )

            st.divider()
            bloque_transcripcion_manual_asistida(prefix, destino=destino)
        else:
            st.info("Suba un escaneo para usar OCR o transcripción manual asistida.")




def limpiar_html_editor(texto):
    import re as _re
    texto = str(texto or "")
    for a, b in [
        ("</p>", "\n"), ("<br>", "\n"), ("<br/>", "\n"), ("<br />", "\n"),
        ("</li>", "\n"), ("<li>", "• "), ("&nbsp;", " "), ("&amp;", "&"),
        ("&lt;", "<"), ("&gt;", ">")
    ]:
        texto = texto.replace(a, b)
    texto = _re.sub(r"<[^>]+>", "", texto)
    texto = _re.sub(r"\n{3,}", "\n\n", texto)
    texto = _re.sub(r"[ \t]+", " ", texto)
    return texto.strip()


def reemplazar_campo_diccionario(datos, campo, *keys_posibles):
    for key in keys_posibles:
        val = st.session_state.get(key, "")
        if val is not None and str(val).strip():
            datos[campo] = limpiar_html_editor(val)
            return datos
    datos[campo] = limpiar_html_editor(datos.get(campo, ""))
    return datos


def selector_ver_y_borrar_registros(prefix="global"):
    st.markdown("### Registros ingresados")
    df = leer_registro_entrevistas()
    if df.empty:
        st.info("Aún no existen registros guardados.")
        return

    opciones = {}
    for idx, fila in df.iterrows():
        fecha = str(fila.get("Fecha_entrevista", fila.get("Fecha", ""))).strip()
        estudiante = str(fila.get("Estudiante", fila.get("Estudiante(s)", ""))).strip()
        curso = str(fila.get("Curso", "")).strip()
        cc = str(fila.get("CC", "")).strip()
        motivo = str(fila.get("Motivo", "")).strip()
        etiqueta = f"Fila {idx + 2} | {fecha} | {estudiante} | {curso} | {cc} | {motivo}"
        opciones[etiqueta] = idx

    seleccion = st.selectbox(
        "Seleccionar registro",
        ["Seleccione registro..."] + list(opciones.keys()),
        key=f"{prefix}_selector_registro_guardado"
    )

    with st.expander("Ver tabla completa de registros", expanded=False):
        st.dataframe(df, use_container_width=True)

    if seleccion == "Seleccione registro...":
        return

    idx = opciones[seleccion]
    fila = df.loc[idx].to_dict()

    with st.expander("Detalle completo del registro seleccionado", expanded=True):
        for k, v in fila.items():
            st.write(f"**{k}:** {v}")

    confirmar = st.checkbox(
        "Confirmo que deseo eliminar este registro",
        key=f"{prefix}_confirmar_borrar_registro"
    )

    if st.button("🗑️ Borrar registro seleccionado", key=f"{prefix}_btn_borrar_registro"):
        if not confirmar:
            st.error("Debe marcar la confirmación antes de borrar.")
            return
        try:
            df_nuevo = df.drop(index=idx).reset_index(drop=True)
            REGISTRO_EXCEL.parent.mkdir(exist_ok=True)
            with pd.ExcelWriter(REGISTRO_EXCEL, engine="openpyxl", mode="w") as writer:
                df_nuevo.to_excel(writer, index=False, sheet_name="Registro_Entrevistas")
            st.success("Registro eliminado correctamente.")
            st.cache_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"No fue posible eliminar el registro: {e}")


def sincronizar_textos_formulario_word(datos_previos, tipo="hogar"):
    if tipo == "hogar":
        datos_previos = reemplazar_campo_diccionario(datos_previos, "Antecedentes", "antecedentes_hogar_editor", "antecedentes_hogar", "antecedentes_ordenados_hogar")
        datos_previos = reemplazar_campo_diccionario(datos_previos, "Acuerdos", "acuerdos_hogar_editor", "acuerdos_hogar", "acuerdos_hogar_mejorado")
        datos_previos = reemplazar_campo_diccionario(datos_previos, "Compromisos", "compromisos_hogar_editor", "compromisos_hogar")
        datos_previos = reemplazar_campo_diccionario(datos_previos, "Resumen_libro", "resumen_hogar", "resumen_hogar_editor")
    else:
        datos_previos = reemplazar_campo_diccionario(datos_previos, "Antecedentes", "antecedentes_estudiante_editor", "antecedentes_estudiante", "antecedentes_ordenados_estudiante")
        datos_previos = reemplazar_campo_diccionario(datos_previos, "Acuerdos", "acuerdos_estudiante_editor", "acuerdos_estudiante", "acuerdos_estudiante_mejorado")
        datos_previos = reemplazar_campo_diccionario(datos_previos, "Compromisos", "compromisos_estudiante_editor", "compromisos_estudiante")
        datos_previos = reemplazar_campo_diccionario(datos_previos, "Resumen_libro", "resumen_estudiante", "resumen_estudiante_editor")
    return datos_previos


def panel_registros_en_formulario(prefix):
    with st.expander("🔎 Ver / borrar registros ingresados", expanded=False):
        selector_ver_y_borrar_registros(prefix)

# ==================================================
# LOGIN BÁSICO
# ==================================================

def conectar():
    return sqlite3.connect(DB_PATH)


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def init_db():
    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE,
            password_hash TEXT,
            nombre TEXT,
            rol TEXT,
            activo INTEGER DEFAULT 1
        )
    """)

    cur.execute("SELECT COUNT(*) FROM usuarios")
    if cur.fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO usuarios (usuario, password_hash, nombre, rol, activo) VALUES (?, ?, ?, ?, 1)",
            ("admin", hash_password("admin123"), "Administrador", "Administrador")
        )

    # Usuario estándar para uso institucional limitado.
    cur.execute(
        "INSERT OR IGNORE INTO usuarios (usuario, password_hash, nombre, rol, activo) VALUES (?, ?, ?, ?, 1)",
        ("usuario", hash_password("usuario123"), "Usuario", "Usuario")
    )

    conn.commit()
    conn.close()


def autenticar(usuario, password):
    conn = conectar()
    cur = conn.cursor()

    cur.execute(
        "SELECT nombre, rol FROM usuarios WHERE usuario=? AND password_hash=? AND activo=1",
        (usuario, hash_password(password))
    )

    data = cur.fetchone()
    conn.close()
    return data


init_db()


# ==================================================
# LECTURA DE EXCEL
# ==================================================

@st.cache_data(show_spinner=False)
def hojas_excel(ruta: str):
    try:
        if not Path(ruta).exists():
            return []
        return pd.ExcelFile(ruta).sheet_names
    except Exception:
        return []


@st.cache_data(show_spinner=False)
def leer_hoja(ruta: str, hoja: str) -> pd.DataFrame:
    try:
        if not Path(ruta).exists():
            return pd.DataFrame()

        df = pd.read_excel(ruta, sheet_name=hoja)
        df.columns = [str(c).strip() for c in df.columns]
        return df.fillna("")

    except Exception:
        return pd.DataFrame()


def buscar_columna(df: pd.DataFrame, nombres_posibles):
    if df.empty:
        return None

    mapa = {str(c).strip().lower(): c for c in df.columns}

    for nombre in nombres_posibles:
        normalizado = str(nombre).strip().lower()
        if normalizado in mapa:
            return mapa[normalizado]

    return None


def lista_columna(df: pd.DataFrame, nombres_posibles):
    columna = buscar_columna(df, nombres_posibles)

    if columna is None:
        return []

    valores = []
    for valor in df[columna].dropna().unique().tolist():
        texto = str(valor).strip()
        if texto and texto.lower() != "nan":
            valores.append(texto)

    return sorted(set(valores))


def valor_fila(fila: dict, nombres_posibles, defecto=""):
    for nombre in nombres_posibles:
        for clave, valor in fila.items():
            if str(clave).strip().lower() == str(nombre).strip().lower():
                texto = str(valor).strip()
                if texto and texto.lower() != "nan":
                    return texto

    return defecto



def opciones_con_seleccione(opciones):
    """
    Agrega opción visual 'Seleccione...' a listas desplegables.
    Evita duplicados.
    """
    opciones = [str(x).strip() for x in opciones if str(x).strip()]
    opciones = [x for x in opciones if x != "Seleccione..."]
    return ["Seleccione..."] + opciones


def limpiar_seleccione(valor):
    """
    Limpia valores 'Seleccione...' para no guardarlos en Word/registro.
    """
    if isinstance(valor, list):
        return [x for x in valor if str(x).strip() != "Seleccione..."]
    if str(valor).strip() == "Seleccione...":
        return ""
    return valor



def texto_lista(valor):
    valor = limpiar_seleccione(valor)
    if isinstance(valor, list):
        return ", ".join([str(x).strip() for x in valor if str(x).strip()])
    return str(valor).strip()


# Bases principales
df_estudiantes = leer_hoja(str(BASE_EXCEL), "Estudiantes")
df_entrevistadores = leer_hoja(str(BASE_EXCEL), "Entrevistadores")
df_departamentos = leer_hoja(str(BASE_EXCEL), "departamento_cita")
if df_departamentos.empty:
    df_departamentos = leer_hoja(str(BASE_EXCEL), "Departamento_cita")
df_lista_faltas = leer_hoja(str(BASE_EXCEL), "lista_faltas")

# Motivos desde base definitiva.
# Hoja esperada: lista_motivo_entrevista.
# Respaldo: Hoja1, según archivo entregado.
df_motivos = leer_hoja(str(BASE_EXCEL), "lista_motivo_entrevista")
if df_motivos.empty:
    df_motivos = leer_hoja(str(BASE_EXCEL), "Hoja1")

# Tipo registro RICE/falta desde hoja lista_faltas.
df_tipo_registro = df_lista_faltas.copy()


# ==================================================
# LISTAS DE FORMULARIO
# ==================================================



def construir_opciones_agrupadas(df, col_categoria, col_valor, extras=None, formato="valor_categoria"):
    """
    Construye opciones agrupadas visualmente:
    ─── CATEGORÍA ───
    opción
    opción

    Streamlit no soporta optgroup nativo, por lo que se usan separadores visuales.
    """
    if df.empty or col_valor is None:
        return extras or []

    resultado = []

    categorias = {}

    for _, fila in df.iterrows():
        categoria = str(fila[col_categoria]).strip() if col_categoria and str(fila[col_categoria]).strip() else "OTROS"
        valor = str(fila[col_valor]).strip()

        if not valor or valor.lower() == "nan":
            continue

        categorias.setdefault(categoria, []).append(valor)

    for categoria in sorted(categorias.keys()):
        resultado.append(f"────────── {categoria.upper()} ──────────")

        for valor in sorted(set(categorias[categoria])):
            resultado.append(valor)

    if extras:
        resultado.extend(extras)

    return resultado


def construir_opciones_faltas_agrupadas():
    """
    Agrupa faltas RICE por categoría.
    Visual:
    ─── CATEGORÍA ───
    tipo | entorno | categoría
    """
    if df_lista_faltas.empty:
        return ["No aplica"]

    col_tipo = buscar_columna(df_lista_faltas, ["tipo", "Tipo", "tipo_falta", "Tipo falta", "Tipo Falta RICE"])
    col_entorno = buscar_columna(df_lista_faltas, ["entorno", "Entorno"])
    col_categoria = buscar_columna(df_lista_faltas, ["categoria", "Categoría", "Categoria", "categoría"])

    categorias = {}
    resultado = ["No aplica"]

    for _, fila in df_lista_faltas.iterrows():
        categoria = str(fila[col_categoria]).strip() if col_categoria else "OTROS"
        tipo = str(fila[col_tipo]).strip() if col_tipo else ""
        entorno = str(fila[col_entorno]).strip() if col_entorno else ""

        partes = [x for x in [tipo, entorno, categoria] if x and x.lower() != "nan"]
        texto = " | ".join(partes)

        if texto:
            categorias.setdefault(categoria, []).append(texto)

    for categoria in sorted(categorias.keys()):
        resultado.append(f"────────── {categoria.upper()} ──────────")

        for item in sorted(set(categorias[categoria])):
            resultado.append(item)

    return resultado



def obtener_cursos():
    col_curso = buscar_columna(df_estudiantes, ["Curso", "curso", "Nivel"])
    if col_curso is None:
        return []

    return sorted(set(
        str(x).strip()
        for x in df_estudiantes[col_curso].dropna().tolist()
        if str(x).strip()
    ))


def obtener_estudiantes_por_curso(curso):
    col_curso = buscar_columna(df_estudiantes, ["Curso", "curso", "Nivel"])

    if col_curso is None or df_estudiantes.empty:
        return pd.DataFrame()

    return df_estudiantes[
        df_estudiantes[col_curso].astype(str).str.strip() == str(curso).strip()
    ].copy()


def seleccionar_estudiante(prefix="hogar"):
    cursos = obtener_cursos()

    if not cursos:
        st.warning("No se encontró listado de cursos. Revisa datos/base_datos_pukaray.xlsx, hoja Estudiantes.")
        nombre = st.text_input("Nombre estudiante", key=f"{prefix}_est_manual")
        curso = st.text_input("Curso", key=f"{prefix}_curso_manual")
        apoderado = st.text_input("Nombre apoderado", key=f"{prefix}_apod_manual")
        run = ""
        return nombre, curso, apoderado, run

    curso = st.selectbox(
        "Curso",
        ["Seleccione curso..."] + cursos,
        key=f"{prefix}_curso"
    )

    if curso == "Seleccione curso...":
        return "", "", "", ""

    filtrados = obtener_estudiantes_por_curso(curso)

    col_nombre = buscar_columna(filtrados, ["Nombre Estudiante", "Nombre", "Estudiante", "Nombre_Completo"])
    col_apod = buscar_columna(filtrados, ["Nombre Apoderado", "Apoderado", "Apoderado Titular"])
    col_run = buscar_columna(filtrados, ["RUN", "Run", "RUT", "Rut"])

    if col_nombre is None or filtrados.empty:
        st.warning("No se encontraron estudiantes asociados al curso seleccionado.")
        return "", curso, "", ""

    nombres = [
        str(x).strip()
        for x in filtrados[col_nombre].dropna().tolist()
        if str(x).strip()
    ]

    estudiante = st.selectbox(
        "Estudiante",
        ["Seleccione estudiante..."] + nombres,
        key=f"{prefix}_estudiante"
    )

    if estudiante == "Seleccione estudiante...":
        return "", curso, "", ""

    fila = filtrados[filtrados[col_nombre].astype(str).str.strip() == estudiante].iloc[0].to_dict()

    apoderado = valor_fila(fila, ["Nombre Apoderado", "Apoderado", "Apoderado Titular"])
    run = valor_fila(fila, ["RUN", "Run", "RUT", "Rut"])

    return estudiante, curso, apoderado, run


def obtener_departamentos():
    deps = lista_columna(df_departamentos, ["Departamentos", "Departamento", "Departamento_cita", "Departamento cita", "Nombre", "Área", "Area"])
    return deps or ["Dirección", "Convivencia Escolar", "Inspectoría General", "UTP", "PIE", "Orientación"]


def obtener_entrevistadores():
    return lista_columna(df_entrevistadores, ["Nombre Entrevistador", "Nombre", "Funcionario", "Entrevistador"])


def cargo_por_entrevistadores(nombres):
    """
    Obtiene cargo(s) desde base_datos_pukaray.xlsx / hoja Entrevistadores.
    Requiere columnas equivalentes a:
    - Nombre Entrevistador / Nombre / Entrevistador / Funcionario
    - Cargo / Función / Rol
    """
    if df_entrevistadores.empty:
        return ""

    col_nombre = buscar_columna(
        df_entrevistadores,
        ["Nombre Entrevistador", "Nombre", "Entrevistador", "Funcionario", "Nombre funcionario"]
    )
    col_cargo = buscar_columna(
        df_entrevistadores,
        ["Cargo", "cargo", "Función", "Funcion", "Rol", "Cargo Entrevistador"]
    )

    if col_nombre is None or col_cargo is None:
        return ""

    if isinstance(nombres, str):
        nombres_lista = [x.strip() for x in nombres.split(",") if x.strip()]
    else:
        nombres_lista = [str(x).strip() for x in nombres if str(x).strip()]

    cargos = []
    for nombre in nombres_lista:
        nombre_norm = nombre.strip().lower()
        for _, fila in df_entrevistadores.iterrows():
            nombre_base = str(fila[col_nombre]).strip()
            if nombre_base.lower() == nombre_norm:
                cargo = str(fila[col_cargo]).strip()
                if cargo and cargo.lower() != "nan" and cargo not in cargos:
                    cargos.append(cargo)

    return ", ".join(cargos)


def seleccionar_participantes(prefix="hogar"):
    entrevistadores = obtener_entrevistadores()

    if not entrevistadores:
        participantes_manual = st.text_input("Participantes / entrevistadores", key=f"{prefix}_participantes_manual")
        return participantes_manual, ""

    seleccion = st.multiselect(
        "Participantes / entrevistadores",
        entrevistadores,
        key=f"{prefix}_participantes",
        placeholder="Seleccione..."
    )

    cargo_auto = cargo_por_entrevistadores(seleccion)

    return texto_lista(seleccion), cargo_auto

def obtener_tipos_registro():
    """
    Tipo registro RICE desde base_datos_pukaray.xlsx / hoja lista_faltas.
    Orden visible: tipo + entorno + categoría.
    """
    return obtener_tipo_falta_categoria_entorno_tipo()

def obtener_motivos():
    """
    Motivos agrupados visualmente por categoría.
    """
    if df_motivos.empty:
        return [
            "────────── CONVIVENCIA ESCOLAR ──────────",
            "Convivencia escolar",
            "Seguimiento de caso",
            "Orientación"
        ]

    col_categoria = buscar_columna(df_motivos, ["CATEGORÍA", "Categoría", "Categoria", "categoria"])
    col_motivo = buscar_columna(df_motivos, ["MOTIVO", "Motivo", "motivo"])

    categorias = {}
    resultado = []

    for _, fila in df_motivos.iterrows():
        categoria = str(fila[col_categoria]).strip() if col_categoria else "OTROS"
        motivo = str(fila[col_motivo]).strip() if col_motivo else ""

        if motivo and motivo.lower() != "nan":
            categorias.setdefault(categoria, []).append(motivo)

    categorias.setdefault("Convivencia Escolar", []).extend([
        "Conciliación / Mediación",
        "Seguimiento de caso",
        "Orientación"
    ])

    for categoria in sorted(categorias.keys()):
        resultado.append(f"────────── {categoria.upper()} ──────────")

        for motivo in sorted(set(categorias[categoria])):
            resultado.append(motivo)

    return resultado

def obtener_tipo_falta_categoria_entorno_tipo():
    """
    Faltas RICE agrupadas por categoría.
    """
    return construir_opciones_faltas_agrupadas()



def obtener_estados_caso():
    """
    Lee estados desde base_datos_pukaray.xlsx / hoja lista_estado_casos.
    Si la hoja no existe, usa valores base.
    """
    df_estados = leer_hoja(str(BASE_EXCEL), "lista_estado_casos")

    if df_estados.empty:
        return ["Abierto", "En proceso", "Cerrado", "No aplica"]

    col_estado = buscar_columna(df_estados, ["Estado", "estado", "Estado caso", "estado_caso"])

    if col_estado is None:
        return ["Abierto", "En proceso", "Cerrado", "No aplica"]

    estados = [
        str(x).strip()
        for x in df_estados[col_estado].dropna().tolist()
        if str(x).strip()
    ]

    return sorted(set(estados)) or ["Abierto", "En proceso", "Cerrado", "No aplica"]




def obtener_checklist_cierre_df():
    """
    Lee checklist desde hoja lista_checklist_cierre_caso.
    Columnas esperadas: Paso, Obligatorio, Descripción.
    """
    df_check = leer_hoja(str(BASE_EXCEL), "lista_checklist_cierre_caso")

    if df_check.empty:
        return pd.DataFrame([
            {"Paso": "Entrevista realizada", "Obligatorio": "Sí", "Descripción": "Existe entrevista formal registrada."},
            {"Paso": "Apoderado informado", "Obligatorio": "Sí", "Descripción": "La familia fue informada formalmente, cuando corresponde."},
            {"Paso": "Registro en libro", "Obligatorio": "Sí", "Descripción": "Se registró resumen en libro de clases o registro institucional."},
            {"Paso": "Antecedentes recopilados", "Obligatorio": "Sí", "Descripción": "Se incorporaron antecedentes mínimos del caso."},
            {"Paso": "Descargos realizados", "Obligatorio": "No", "Descripción": "Descargos realizados, si corresponde."},
            {"Paso": "Medidas formativas aplicadas", "Obligatorio": "No", "Descripción": "Medidas formativas aplicadas, si corresponde."},
            {"Paso": "Seguimiento realizado", "Obligatorio": "Sí", "Descripción": "Seguimiento posterior realizado, cuando corresponde."},
            {"Paso": "Equipo informado", "Obligatorio": "No", "Descripción": "Equipo informado, si corresponde."},
            {"Paso": "Resolución comunicada", "Obligatorio": "Sí", "Descripción": "Resolución o cierre comunicado, cuando corresponde."},
        ])

    col_paso = buscar_columna(df_check, ["Paso", "paso"])
    col_obligatorio = buscar_columna(df_check, ["Obligatorio", "obligatorio"])
    col_desc = buscar_columna(df_check, ["Descripción", "Descripcion", "descripcion", "descripción"])

    if col_paso is None:
        return pd.DataFrame()

    salida = pd.DataFrame()
    salida["Paso"] = df_check[col_paso].astype(str)

    if col_obligatorio:
        salida["Obligatorio"] = df_check[col_obligatorio].astype(str)
    else:
        salida["Obligatorio"] = "No"

    if col_desc:
        salida["Descripción"] = df_check[col_desc].astype(str)
    else:
        salida["Descripción"] = ""

    salida = salida[salida["Paso"].str.strip() != ""]
    return salida.reset_index(drop=True)


def checklist_cierre_interactivo(prefix):
    """
    Muestra checklist con estado por paso:
    Cumplido / Pendiente / No aplica.
    Retorna resumen textual y estado sugerido.
    """
    df_check = obtener_checklist_cierre_df()

    if df_check.empty:
        st.warning("No hay checklist de cierre configurado.")
        return "", "No aplica", []

    st.markdown("#### Checklist seguimiento / cierre de caso")
    st.caption("Un caso solo debería cerrarse si no quedan pasos pendientes. Usa 'No aplica' cuando el paso no corresponda al caso.")

    estados_validos = ["Pendiente", "Cumplido", "No aplica"]
    pendientes = []
    resumen = []

    for idx, fila in df_check.iterrows():
        paso = str(fila.get("Paso", "")).strip()
        obligatorio = str(fila.get("Obligatorio", "No")).strip()
        descripcion = str(fila.get("Descripción", "")).strip()

        col_a, col_b = st.columns([2, 1])

        with col_a:
            etiqueta = f"{paso}"
            if obligatorio.lower() in ["sí", "si", "s"]:
                etiqueta += " *"
            st.write(f"**{etiqueta}**")
            if descripcion:
                st.caption(descripcion)

        with col_b:
            estado = st.selectbox(
                "Estado",
                estados_validos,
                key=f"{prefix}_check_{idx}",
                label_visibility="collapsed"
            )

        resumen.append(f"{paso}: {estado}")

        if estado == "Pendiente":
            pendientes.append(paso)

    if pendientes:
        estado_sugerido = "En proceso"
        st.warning("Caso con acciones pendientes. Estado sugerido: En proceso.")
        with st.expander("Ver pendientes"):
            for pendiente in pendientes:
                st.write(f"• {pendiente}")
    else:
        estado_sugerido = "Cerrado"
        st.success("Checklist sin pendientes. Estado sugerido: Cerrado.")

    return "\n".join(resumen), estado_sugerido, pendientes



def obtener_estado_detalle_caso_df():
    """
    Lee lista de estado/detalle desde base_datos_pukaray.xlsx / hoja lista_estado_detalle_caso.
    Columnas esperadas: ESTADO, DETALLES.
    """
    df = leer_hoja(str(BASE_EXCEL), "lista_estado_detalle_caso")

    if df.empty:
        return pd.DataFrame([
            {"ESTADO": "En Proceso", "DETALLES": "Recopilación de antecedentes y entrevistas."},
            {"ESTADO": "En Proceso", "DETALLES": "Protocolo activado / Citación a apoderados."},
            {"ESTADO": "En Seguimiento", "DETALLES": "Monitoreo quincenal de acuerdos vigentes."},
            {"ESTADO": "Cerrado", "DETALLES": "Caso resuelto con compromisos firmados."},
            {"ESTADO": "No aplica", "DETALLES": "Registro informativo sin apertura de caso."},
        ])

    col_estado = buscar_columna(df, ["ESTADO", "Estado", "estado"])
    col_detalles = buscar_columna(df, ["DETALLES", "Detalles", "detalles", "Detalle", "detalle"])

    if col_estado is None or col_detalles is None:
        return pd.DataFrame([
            {"ESTADO": "En Proceso", "DETALLES": "Recopilación de antecedentes y entrevistas."},
            {"ESTADO": "En Seguimiento", "DETALLES": "Monitoreo quincenal de acuerdos vigentes."},
            {"ESTADO": "Cerrado", "DETALLES": "Caso resuelto con compromisos firmados."},
            {"ESTADO": "No aplica", "DETALLES": "Registro informativo sin apertura de caso."},
        ])

    salida = pd.DataFrame()
    salida["ESTADO"] = df[col_estado].astype(str)
    salida["DETALLES"] = df[col_detalles].astype(str)
    salida = salida[(salida["ESTADO"].str.strip() != "") & (salida["DETALLES"].str.strip() != "")]
    return salida.reset_index(drop=True)


def selector_estado_detalle_caso(prefix):
    """
    Selector encadenado Estado -> Detalle para conclusión de caso.
    """
    df = obtener_estado_detalle_caso_df()

    estados = sorted(set(df["ESTADO"].astype(str).str.strip().tolist()))
    estado = st.selectbox(
        "Estado institucional del caso",
        opciones_con_seleccione(estados),
        key=f"{prefix}_estado_institucional"
    )
    estado = limpiar_seleccione(estado)

    detalles = df[df["ESTADO"].astype(str).str.strip() == estado]["DETALLES"].astype(str).tolist() if estado else []

    detalle = st.selectbox(
        "Detalle del estado",
        opciones_con_seleccione(detalles),
        key=f"{prefix}_detalle_estado"
    )
    detalle = limpiar_seleccione(detalle)

    if estado and detalle:
        frase = f"Conclusión del estado del caso: {estado}. {detalle}"
        st.info(frase)
    else:
        frase = ""
        st.caption("Seleccione estado institucional y detalle para generar conclusión del caso.")

    return estado, detalle, frase




def leer_registro_entrevistas():
    """
    Lee el registro operativo de entrevistas desde datos/registro_entrevistas.xlsx.
    """
    if not REGISTRO_EXCEL.exists():
        return pd.DataFrame()

    try:
        return pd.read_excel(REGISTRO_EXCEL, sheet_name="Registro_Entrevistas").fillna("")
    except Exception:
        return pd.DataFrame()



def obtener_correlativos_existentes(columna):
    df = leer_registro_entrevistas()
    if df.empty or columna not in df.columns:
        return []
    return [
        str(x).strip()
        for x in df[columna].dropna().unique().tolist()
        if str(x).strip() and str(x).strip().lower() != "sin asignar"
    ]


def generar_correlativo(prefijo, columna):
    anio = datetime.now().year
    existentes = obtener_correlativos_existentes(columna)
    max_num = 0
    for valor in existentes:
        valor = str(valor).strip()
        if valor.startswith(f"{prefijo}-{anio}-"):
            try:
                max_num = max(max_num, int(valor.split("-")[-1]))
            except Exception:
                pass
    return f"{prefijo}-{anio}-{max_num + 1:04d}"


def generar_numero_entrevista_por_cc(cc):
    df = leer_registro_entrevistas()
    cc = str(cc).strip()
    if not cc:
        return "NE-SIN-CC-001"
    if df.empty or "CC" not in df.columns or "Numero_entrevista" not in df.columns:
        return f"NE-{cc}-001"
    df_cc = df[df["CC"].astype(str).str.strip() == cc].copy()
    max_num = 0
    for valor in df_cc["Numero_entrevista"].dropna().tolist():
        valor = str(valor).strip()
        if valor.lower() == "sin asignar":
            continue
        try:
            max_num = max(max_num, int(valor.split("-")[-1]))
        except Exception:
            pass
    return f"NE-{cc}-{max_num + 1:03d}"


def selector_correlativo(prefix, etiqueta, prefijo, columna, cc=None):
    modo = st.radio(
        etiqueta,
        ["Sin asignar", "Asignar"],
        horizontal=True,
        key=f"{prefix}_{columna}_modo"
    )
    if modo == "Asignar":
        if columna == "Numero_entrevista":
            valor = generar_numero_entrevista_por_cc(cc)
        else:
            valor = generar_correlativo(prefijo, columna)
        st.success(f"{etiqueta}: {valor}")
        return valor
    return "Sin asignar"


def obtener_cc_existentes():
    """
    Devuelve lista simple de códigos de caso existentes.
    """
    df = leer_registro_entrevistas()

    if df.empty or "CC" not in df.columns:
        return []

    return sorted([
        str(x).strip()
        for x in df["CC"].dropna().unique().tolist()
        if str(x).strip()
    ])


def generar_nuevo_cc():
    """
    Genera un código correlativo automático:
    CC-AAAA-0001
    """
    anio = datetime.now().year
    existentes = obtener_cc_existentes()
    max_num = 0

    for cc in existentes:
        cc = str(cc).strip()
        if cc.startswith(f"CC-{anio}-"):
            try:
                max_num = max(max_num, int(cc.split("-")[-1]))
            except Exception:
                pass

    return f"CC-{anio}-{max_num + 1:04d}"


def construir_resumen_caso(df_caso):
    """
    Construye un resumen humano para reconocer el caso.
    """
    if df_caso.empty:
        return ""

    primera = df_caso.iloc[0]
    ultima = df_caso.iloc[-1]

    estudiante = str(primera.get("Estudiante", primera.get("Estudiante(s)", ""))).strip()
    curso = str(primera.get("Curso", "")).strip()
    motivo = str(primera.get("Motivo", "")).strip()
    estado = str(ultima.get("Estado_institucional", ultima.get("Estado_caso", ""))).strip()
    fecha = str(primera.get("Fecha_entrevista", "")).strip()

    partes = []
    if estudiante:
        partes.append(estudiante)
    if curso:
        partes.append(curso)
    if motivo:
        partes.append(motivo)
    if estado:
        partes.append(f"Estado: {estado}")
    if fecha:
        partes.append(f"Inicio: {fecha}")

    return " | ".join(partes)


def opciones_cc_con_resumen():
    """
    Retorna opciones visuales:
    CC-2026-0001 | Estudiante | Curso | Motivo | Estado | Inicio
    """
    df = leer_registro_entrevistas()

    if df.empty or "CC" not in df.columns:
        return {}

    opciones = {}

    for cc in obtener_cc_existentes():
        df_caso = df[df["CC"].astype(str).str.strip() == cc].copy()
        resumen = construir_resumen_caso(df_caso)
        etiqueta = f"{cc} | {resumen}" if resumen else cc
        opciones[etiqueta] = cc

    return opciones


def selector_cc(prefix):
    """
    Selector seguro de Código de Caso:
    - Crear nuevo CC automático
    - Usar CC existente con resumen humano
    """
    st.markdown("#### Código de caso (cc)")
    st.caption("El cc permite asociar entrevistas, acciones, checklist, protocolos, evidencias y cierre del caso.")

    modo = st.radio(
        "Tipo de registro de caso",
        ["Crear nuevo CC automático", "Usar CC existente"],
        horizontal=True,
        key=f"{prefix}_cc_tipo"
    )

    if modo == "Crear nuevo CC automático":
        cc = generar_nuevo_cc()
        st.success(f"Nuevo código de caso asignado: {cc}")
        return cc

    opciones = opciones_cc_con_resumen()

    if not opciones:
        st.warning("No existen casos previos registrados. Se generará un nuevo CC automático.")
        cc = generar_nuevo_cc()
        st.success(f"Nuevo código de caso asignado: {cc}")
        return cc

    st.info("Usa un CC existente solo si esta entrevista pertenece al mismo caso ya iniciado.")

    etiqueta = st.selectbox(
        "Seleccionar caso existente",
        list(opciones.keys()),
        key=f"{prefix}_cc_existente"
    )

    cc = opciones[etiqueta]
    st.success(f"Registro asociado al caso: {cc}")
    return cc



def obtener_protocolos_df():
    df = leer_hoja(str(BASE_EXCEL), "protocolos")
    return df

def obtener_protocolos_formateados():
    """
    Protocolos agrupados por categoría.
    Visual:
    ─── CATEGORÍA ───
    detalle | protocolo | categoría
    """
    df = obtener_protocolos_df()

    if df.empty:
        return [
            "────────── CONVIVENCIA ESCOLAR ──────────",
            "Entrevistas y seguimiento inicial | Activación convivencia | Convivencia Escolar"
        ]

    col_categoria = buscar_columna(df, ["categoria", "Categoría", "Categoria"])
    col_protocolo = buscar_columna(df, ["protocolo", "Protocolo"])
    col_detalle = buscar_columna(df, ["detalle", "detalles", "Detalles"])

    categorias = {}
    resultado = []

    for _, fila in df.iterrows():
        categoria = str(fila[col_categoria]).strip() if col_categoria else "OTROS"

        partes = []

        if col_detalle and str(fila[col_detalle]).strip():
            partes.append(str(fila[col_detalle]).strip())

        if col_protocolo and str(fila[col_protocolo]).strip():
            partes.append(str(fila[col_protocolo]).strip())

        if col_categoria and str(fila[col_categoria]).strip():
            partes.append(str(fila[col_categoria]).strip())

        texto = " | ".join([x for x in partes if x and x.lower() != "nan"])

        if texto:
            categorias.setdefault(categoria, []).append(texto)

    for categoria in sorted(categorias.keys()):
        resultado.append(f"────────── {categoria.upper()} ──────────")

        for item in sorted(set(categorias[categoria])):
            resultado.append(item)

    return resultado


def corregir_ortografia_basica(texto):
    correcciones = {
        "hogara": "hogar", "revision": "revisión", "revison": "revisión",
        "autolesion": "autolesión", "situacion": "situación",
        "resolucion": "resolución", "mantine": "mantiene",
        "direccion": "Dirección", "psicologo": "psicólogo",
        "academico": "académico", "pedagogico": "pedagógico",
        "intervencion": "intervención", "protoclo": "protocolo",
        "segumiento": "seguimiento"
    }
    palabras = str(texto).split()
    salida = []
    for palabra in palabras:
        limpia = palabra.strip(".,;:()[]{}¡!¿?")
        prefijo = palabra[:len(palabra) - len(palabra.lstrip(".,;:()[]{}¡!¿?"))]
        sufijo = palabra[len(palabra.rstrip(".,;:()[]{}¡!¿?")):]
        nueva = correcciones.get(limpia.lower(), limpia)
        if limpia[:1].isupper():
            nueva = nueva[:1].upper() + nueva[1:]
        salida.append(prefijo + nueva + sufijo)
    return " ".join(salida)


def separar_acuerdos(texto):
    """
    Separa acuerdos/compromisos. Respeta viñetas, guiones, numeración y saltos de línea.
    """
    texto = corregir_ortografia_basica(str(texto).strip())
    if not texto:
        return []

    lineas = []
    for raw in texto.split("\n"):
        item = raw.strip()
        if not item:
            continue
        item = item.strip("•-–—* ").strip()
        if len(item) > 2 and item[0].isdigit() and item[1] in [".", ")"]:
            item = item[2:].strip()
        if item:
            lineas.append(item)

    if len(lineas) > 1:
        return lineas

    frase = lineas[0] if lineas else texto
    for sep in [". Se ", ". El ", ". La ", ". Los ", ". Las ", ";"]:
        if sep in frase:
            partes = []
            trozos = frase.split(sep)
            for i, t in enumerate(trozos):
                t = t.strip(". ; ")
                if not t:
                    continue
                if i > 0 and sep.startswith(". "):
                    t = sep.replace(".", "").strip() + " " + t
                partes.append(t)
            return partes

    return [frase]



def transformar_acuerdos_tecnicos(texto_usuario, contexto="general"):
    """
    Redacta acuerdos/compromisos con tono institucional más humano.
    Cada viñeta o línea se mantiene como punto independiente.
    """
    lineas = separar_acuerdos(texto_usuario)
    if not lineas:
        return ""

    resultado = []
    usados = set()

    for linea in lineas:
        original = corregir_ortografia_basica(linea).strip()
        bajo = original.lower()
        redaccion = None

        if any(x in bajo for x in ["toma conocimiento", "ya lo sabe"]):
            redaccion = "El apoderado toma conocimiento formal de los antecedentes informados respecto de su estudiante y de las acciones de seguimiento acordadas."
        elif any(x in bajo for x in ["informar al apoderado", "se informará al apoderado", "avisar al apoderado"]):
            redaccion = "Se acuerda informar formalmente al apoderado respecto de los antecedentes abordados y de las medidas institucionales adoptadas."
        elif any(x in bajo for x in ["autolesión", "autoagresión", "riesgo"]):
            redaccion = "Se activa el procedimiento institucional correspondiente, priorizando el resguardo, acompañamiento y seguimiento del estudiante."
        elif any(x in bajo for x in ["alta", "mejoría", "mejoria"]):
            redaccion = "Se deja constancia de la evolución favorable observada en el estudiante, considerando los antecedentes registrados durante el proceso de acompañamiento."
        elif any(x in bajo for x in ["observación preventiva", "observacion preventiva", "mantiene en observación"]):
            redaccion = "Se mantiene observación preventiva del estudiante, resguardando continuidad de apoyo y seguimiento institucional."
        elif "protocolo" in bajo:
            redaccion = "Se activa el protocolo correspondiente según los antecedentes recabados y la normativa interna aplicable."
        elif any(x in bajo for x in ["seguimiento", "monitoreo"]):
            redaccion = "El establecimiento realizará seguimiento institucional de los acuerdos adoptados, dejando registro de las acciones relevantes."
        elif any(x in bajo for x in ["evidencia", "evidencias", "recopilar", "juntar"]):
            redaccion = "Se acuerda recopilar y resguardar antecedentes pertinentes para complementar el análisis institucional del caso."
        elif any(x in bajo for x in ["dirección", "equipo", "informar al equipo"]):
            redaccion = "Se informará al equipo correspondiente y a Dirección, cuando proceda, para coordinar las acciones institucionales pertinentes."
        elif any(x in bajo for x in ["reflexión", "reflexionar", "se conversa", "conversa"]):
            redaccion = "Se realiza una instancia de reflexión formativa orientada a favorecer la comprensión de los hechos y sus efectos en la convivencia escolar."
        elif any(x in bajo for x in ["académico", "academico", "rendimiento", "notas", "profesor jefe"]):
            redaccion = "Se acuerda realizar seguimiento académico en coordinación con profesor/a jefe y equipos pertinentes, estableciendo apoyos y metas de mejora."
        elif any(x in bajo for x in ["puntualidad", "atrasos"]):
            redaccion = "Se establece compromiso de puntualidad y cumplimiento de horarios institucionales."
        elif any(x in bajo for x in ["disculpas", "pedir disculpas"]):
            redaccion = "Se promoverá una instancia de reparación relacional mediante disculpas formales, según corresponda."
        elif any(x in bajo for x in ["reparar", "reparación", "arreglar"]):
            redaccion = "Se acuerda una acción reparatoria proporcional a la situación abordada, con finalidad formativa."

        if not redaccion:
            if contexto == "compromiso_hp":
                redaccion = f"La familia se compromete a {original[0].lower() + original[1:] if len(original) > 1 else original}, manteniendo comunicación con el establecimiento para favorecer el seguimiento del estudiante."
            elif contexto == "compromiso_e":
                redaccion = f"El estudiante se compromete a {original[0].lower() + original[1:] if len(original) > 1 else original}, en coherencia con los acuerdos formativos establecidos en la entrevista."
            elif contexto == "hp":
                redaccion = f"Se acuerda con el apoderado abordar el siguiente punto: {original}, resguardando el acompañamiento y seguimiento del estudiante."
            elif contexto == "e":
                redaccion = f"Se acuerda con el estudiante abordar el siguiente punto: {original}, promoviendo responsabilidad, reflexión y mejora progresiva."
            else:
                redaccion = f"Se deja establecido el siguiente acuerdo: {original}."

        if redaccion in usados:
            redaccion = f"Se complementa lo anterior con el siguiente antecedente o acción: {original}."
        usados.add(redaccion)
        resultado.append(f"• {redaccion}")

    return "\n".join(resultado)




VINCULOS = [
    "Apoderados",
    "Padres",
    "Madre",
    "Padre",
    "Apoderado titular",
    "Apoderado suplente",
    "Abuela/o",
    "Tía/o",
    "Tutor legal",
    "Otro"
]


# ==================================================
# REDACCIÓN CONTROLADA SIN IA
# ==================================================

def generar_acuerdos_base(datos):
    acuerdos = [
        "1. El/la apoderado/a toma conocimiento de los antecedentes informados por el establecimiento.",
        "2. Se acuerda mantener comunicación permanente entre familia y establecimiento, según corresponda.",
        "3. La familia se compromete a reforzar normas, responsabilidades y acuerdos asociados al motivo de la entrevista.",
        "4. El establecimiento realizará seguimiento desde el/los departamento(s) que cita(n).",
        "5. Se deja constancia de que los acuerdos se basan exclusivamente en los antecedentes registrados en esta entrevista."
    ]

    tipo_falta = datos.get("tipo_falta", "")
    if tipo_falta and tipo_falta != "No aplica":
        acuerdos.insert(
            3,
            f"4. Para efectos de seguimiento interno, se considera la clasificación seleccionada: {tipo_falta}."
        )

    return "\n".join(acuerdos)


def generar_resumen_libro(datos):
    estudiante = datos.get("estudiante", "estudiante")
    curso = datos.get("curso", "")
    motivo = datos.get("motivo", "motivo no informado")
    departamento = datos.get("departamento", "departamento no informado")

    return (
        f"Se realiza entrevista de apoderado respecto de {estudiante}, curso {curso}. "
        f"Motivo registrado: {motivo}. "
        f"Departamento que cita: {departamento}. "
        f"Se informa situación, se establecen acuerdos y se deja registro para seguimiento institucional."
    )


# ==================================================
# WORD: FICHA APODERADO EXACTA
# ==================================================

def limpiar_nombre_archivo(texto):
    permitido = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-.áéíóúÁÉÍÓÚñÑ"
    texto = texto.replace(" ", "_").replace("/", "-")
    return "".join(c for c in texto if c in permitido)


def poner_texto(celda, texto):
    celda.text = "" if texto is None else str(texto)


def rellenar_docx_por_tablas(datos):
    """
    Rellena la ficha real Ficha_Entrevista_APODERADO.docx según su estructura:
    Tabla 0: estudiante/apoderado
    Tabla 1: curso/fecha/hora
    Tabla 2: entrevistador/cargo
    Tabla 3: departamento/número/asistencia
    Tabla 4: motivo
    Tabla 5: acuerdos
    """
    if PLANTILLA_APODERADO.exists():
        doc = Document(str(PLANTILLA_APODERADO))
    else:
        doc = Document()
        doc.add_heading("FICHA ENTREVISTA APODERADO", 0)

    # Si la plantilla es la oficial, trae 7 tablas.
    if len(doc.tables) >= 6:
        try:
            # Tabla 0
            poner_texto(doc.tables[0].rows[0].cells[1], datos.get("estudiante", ""))
            poner_texto(doc.tables[0].rows[1].cells[1], datos.get("apoderado", ""))

            # Tabla 1
            poner_texto(doc.tables[1].rows[0].cells[1], datos.get("curso", ""))
            poner_texto(doc.tables[1].rows[0].cells[3], datos.get("fecha", ""))
            poner_texto(doc.tables[1].rows[0].cells[5], datos.get("hora", ""))

            # Tabla 2
            entrevistadores_txt = datos.get("entrevistador", "")
            cargos_txt = datos.get("cargo_entrevistador", "")

            poner_texto(doc.tables[2].rows[0].cells[1], entrevistadores_txt)
            poner_texto(doc.tables[2].rows[0].cells[3], cargos_txt if cargos_txt else "No informado")

            # Tabla 3
            poner_texto(doc.tables[3].rows[0].cells[1], datos.get("departamento", ""))
            poner_texto(doc.tables[3].rows[0].cells[5], "")  # Número entrevista eliminado
            # Asistencia: mantener visibles las opciones Sí/No y marcar con X la opción seleccionada.
            poner_texto(doc.tables[3].rows[1].cells[1], "Sí  X" if datos.get("asiste_apoderado") == "Sí" else "Sí")
            poner_texto(doc.tables[3].rows[1].cells[2], "No  X" if datos.get("asiste_apoderado") == "No" else "No")
            poner_texto(doc.tables[3].rows[1].cells[5], "Sí  X" if datos.get("asiste_estudiante") == "Sí" else "Sí")
            poner_texto(doc.tables[3].rows[1].cells[6], "No  X" if datos.get("asiste_estudiante") == "No" else "No")

            # Tabla 4
            poner_texto(doc.tables[4].rows[0].cells[1], datos.get("motivo_word", ""))

            # Tabla 5
            texto_final = datos.get("acuerdos", "").strip()
            if datos.get("compromisos"):
                texto_final += "\n\nCOMPROMISOS:\n" + datos.get("compromisos", "").strip()
            if datos.get("conclusion_estado"):
                texto_final += "\n\n" + datos.get("conclusion_estado", "").strip()
            poner_texto(doc.tables[5].rows[0].cells[1], texto_final.strip())

        except Exception:
            doc.add_paragraph("No fue posible rellenar la plantilla oficial por estructura de tablas.")

    else:
        doc.add_paragraph(f"Nombre Estudiante: {datos.get('estudiante', '')}")
        doc.add_paragraph(f"Nombre Apoderado: {datos.get('apoderado', '')}")
        doc.add_paragraph(f"Curso: {datos.get('curso', '')}")
        doc.add_paragraph(f"Fecha: {datos.get('fecha', '')}")
        doc.add_paragraph(f"Hora: {datos.get('hora', '')}")
        doc.add_paragraph(f"Entrevistador: {datos.get('entrevistador', '')}")
        doc.add_heading("Motivo de la Entrevista", level=1)
        doc.add_paragraph(datos.get("motivo_word", ""))
        doc.add_heading("Acuerdos o Conclusiones", level=1)
        doc.add_paragraph((datos.get("acuerdos", "") + "\n\n" + datos.get("conclusion_estado", "")).strip())

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer



# ==================================================
# WORD: FICHA ESTUDIANTE EXACTA
# ==================================================

def generar_acuerdos_base_estudiante(datos):
    acuerdos = [
        "1. El/la estudiante toma conocimiento de los antecedentes abordados en la entrevista.",
        "2. Se acuerda mantener una actitud de respeto, responsabilidad y colaboración en el contexto escolar.",
        "3. El/la estudiante se compromete a cumplir los acuerdos establecidos durante la entrevista.",
        "4. El establecimiento realizará seguimiento de los acuerdos desde el/los departamento(s) que cita(n).",
        "5. Se deja constancia de que los acuerdos se basan exclusivamente en los antecedentes registrados en esta entrevista."
    ]

    tipo_falta = datos.get("tipo_falta", "")
    if tipo_falta and tipo_falta != "No aplica":
        acuerdos.insert(
            3,
            f"4. Para efectos de seguimiento interno, se considera la clasificación seleccionada: {tipo_falta}."
        )

    return "\n".join(acuerdos)


def generar_resumen_libro_estudiante(datos):
    estudiante = datos.get("estudiante", "estudiante")
    curso = datos.get("curso", "")
    motivo = datos.get("motivo", "motivo no informado")
    departamento = datos.get("departamento", "departamento no informado")

    return (
        f"Se realiza entrevista a estudiante {estudiante}, curso {curso}. "
        f"Motivo registrado: {motivo}. "
        f"Departamento que cita: {departamento}. "
        f"Se abordan antecedentes, se establecen acuerdos y se deja registro para seguimiento institucional."
    )



def ordenar_antecedentes_estudiante(datos):
    """
    Ordena los antecedentes sin inventar información.
    No usa IA externa; solo estructura el texto ingresado.
    """
    antecedentes = str(datos.get("antecedentes", "")).strip()
    motivo = str(datos.get("motivo", "")).strip()
    tipo_falta = str(datos.get("tipo_falta", "")).strip()
    tipo_relato = str(datos.get("tipo_relato", "")).strip()

    partes = []

    if tipo_relato:
        partes.append(f"Tipo de relato: {tipo_relato}.")

    if motivo:
        partes.append(f"Motivo registrado: {motivo}.")

    if tipo_falta:
        partes.append(f"Tipo Falta RICE: {tipo_falta}.")

    if antecedentes:
        partes.append(f"Antecedentes declarados: {antecedentes}")
    else:
        partes.append("Antecedentes declarados: No informado.")

    return "\n".join(partes)


def rellenar_docx_estudiante(datos):
    """
    Rellena la ficha real Ficha_Entrevista_ESTUDIANTE.docx según su estructura:
    Tabla 0: estudiante
    Tabla 1: curso/fecha/hora
    Tabla 2: entrevistador/cargo
    Tabla 3: motivo
    Tabla 4: acuerdos
    """
    if PLANTILLA_ESTUDIANTE.exists():
        doc = Document(str(PLANTILLA_ESTUDIANTE))
    else:
        doc = Document()
        doc.add_heading("FICHA ENTREVISTA ESTUDIANTE", 0)

    if len(doc.tables) >= 5:
        try:
            # Tabla 0
            poner_texto(doc.tables[0].rows[0].cells[1], datos.get("estudiante", ""))

            # Tabla 1
            poner_texto(doc.tables[1].rows[0].cells[1], datos.get("curso", ""))
            poner_texto(doc.tables[1].rows[0].cells[3], datos.get("fecha", ""))
            poner_texto(doc.tables[1].rows[0].cells[5], datos.get("hora", ""))

            # Tabla 2
            entrevistadores_txt = datos.get("entrevistador", "")
            departamento_txt = datos.get("departamento", "")
            cargos_txt = datos.get("cargo_entrevistador", "")

            if departamento_txt:
                entrevistadores_txt = f"{entrevistadores_txt}\nDepartamento que cita: {departamento_txt}"

            poner_texto(doc.tables[2].rows[0].cells[1], entrevistadores_txt)
            poner_texto(doc.tables[2].rows[0].cells[3], cargos_txt if cargos_txt else "No informado")

            # Tabla 3
            poner_texto(doc.tables[3].rows[0].cells[1], datos.get("motivo_word", ""))

            # Tabla 4
            acuerdos_finales = str(datos.get("acuerdos", "")).strip()

            if datos.get("compromisos"):
                acuerdos_finales += "\n\nCOMPROMISOS:\n" + str(datos.get("compromisos", "")).strip()

            if datos.get("conclusion_estado"):
                acuerdos_finales += "\n\n" + str(datos.get("conclusion_estado", "")).strip()

            poner_texto(
                doc.tables[4].rows[0].cells[1],
                acuerdos_finales if acuerdos_finales.strip() else "Sin acuerdos registrados."
            )

        except Exception:
            doc.add_paragraph("No fue posible rellenar la plantilla oficial por estructura de tablas.")
    else:
        doc.add_paragraph(f"Nombre Estudiante: {datos.get('estudiante', '')}")
        doc.add_paragraph(f"Curso: {datos.get('curso', '')}")
        doc.add_paragraph(f"Fecha: {datos.get('fecha', '')}")
        doc.add_paragraph(f"Hora: {datos.get('hora', '')}")
        doc.add_paragraph(f"Entrevistador: {datos.get('entrevistador', '')}")
        doc.add_paragraph(f"Cargo: {datos.get('cargo_entrevistador', '')}")
        doc.add_paragraph(f"Departamento que cita: {datos.get('departamento', '')}")
        doc.add_heading("Motivo de la Entrevista", level=1)
        doc.add_paragraph(datos.get("motivo_word", ""))
        doc.add_heading("Acuerdos o Conclusiones", level=1)
        doc.add_paragraph((datos.get("acuerdos", "") + "\n\n" + datos.get("conclusion_estado", "")).strip())

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ==================================================
# REGISTRO EXCEL SEPARADO
# ==================================================

def guardar_registro(datos):
    """
    Guarda datos mínimos en archivo separado para evitar bloquear/corromper base_datos_pukaray.xlsx.
    """
    registro = {
        "CC": datos.get("cc", ""),
        "Folio": datos.get("folio", ""),
        "Numero_entrevista": datos.get("numero_entrevista", ""),
        "Resumen_caso": datos.get("resumen_caso", ""),
        "Fecha_registro_sistema": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
        "Fecha_entrevista": datos.get("fecha", ""),
        "Hora": datos.get("hora", ""),
        "Tipo_registro": datos.get("tipo_registro", "Entrevista Hogar / Apoderado"),
        "Departamento_que_cita": datos.get("departamento", ""),
        "Curso": datos.get("curso", ""),
        "Estudiante": datos.get("estudiante", ""),
        "RUN": datos.get("run", ""),
        "Apoderado": datos.get("apoderado", ""),
        "Vínculo": datos.get("vinculo", ""),
        "Motivo": datos.get("motivo", ""),
        "Tipo_registro_RICE": datos.get("tipo_registro_rice", ""),
        "Tipo_falta_categoria_entorno_tipo": datos.get("tipo_falta", ""),
        "Estado_caso": datos.get("estado_caso", ""),
        "Condicion_caso": datos.get("condicion_caso", ""),
        "Estado_sugerido_checklist": datos.get("estado_sugerido_checklist", ""),
        "Estado_institucional": datos.get("estado_institucional", ""),
        "Detalle_estado": datos.get("detalle_estado", ""),
        "Protocolos_aplicados": datos.get("protocolos_aplicados", ""),
        "Compromisos": datos.get("compromisos", ""),
        "Checklist_cierre": datos.get("checklist_cierre", ""),
        "Pendientes_checklist": datos.get("pendientes_checklist", ""),
        "Usuario_registro": st.session_state.get("usuario", "")
    }

    try:
        if REGISTRO_EXCEL.exists():
            df_registro = pd.read_excel(REGISTRO_EXCEL, sheet_name="Registro_Entrevistas")
        else:
            df_registro = pd.DataFrame()

        df_nuevo = pd.concat([df_registro, pd.DataFrame([registro])], ignore_index=True)

        with pd.ExcelWriter(REGISTRO_EXCEL, engine="openpyxl", mode="w") as writer:
            df_nuevo.to_excel(writer, sheet_name="Registro_Entrevistas", index=False)

        st.success("Registro guardado en datos/registro_entrevistas.xlsx")

    except PermissionError:
        st.error("No se pudo guardar porque registro_entrevistas.xlsx está abierto. Cierra Excel y vuelve a intentar.")
    except Exception as e:
        st.error(f"No fue posible guardar el registro: {e}")



# ==================================================
# GESTIÓN DE CASOS POR CC
# ==================================================

def normalizar_estado_checklist(valor):
    valor = str(valor).strip()
    bajo = valor.lower()
    if "cumplido" in bajo:
        return "Cumplido"
    if "no aplica" in bajo:
        return "No aplica"
    if "pendiente" in bajo:
        return "Pendiente"
    return valor or "Pendiente"


def parsear_checklist_texto(texto):
    resultado = {}
    for linea in str(texto).split("\n"):
        linea = linea.strip()
        if not linea or ":" not in linea:
            continue
        paso, estado = linea.rsplit(":", 1)
        paso = paso.strip(" •-*0123456789.").strip()
        estado = normalizar_estado_checklist(estado)
        if paso:
            resultado[paso] = estado
    return resultado


def checklist_base_pasos():
    df_check = obtener_checklist_cierre_df()
    if df_check.empty or "Paso" not in df_check.columns:
        return []
    return [str(x).strip() for x in df_check["Paso"].tolist() if str(x).strip()]


def cargar_registros():
    if not REGISTRO_EXCEL.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(REGISTRO_EXCEL, sheet_name="Registro_Entrevistas").fillna("")
    except Exception:
        return pd.DataFrame()


def consolidar_checklist_por_cc(df_cc):
    pasos = checklist_base_pasos()
    estados = {paso: "Pendiente" for paso in pasos}
    prioridad = {"Pendiente": 1, "No aplica": 2, "Cumplido": 3}

    if df_cc.empty or "Checklist_cierre" not in df_cc.columns:
        return estados

    for _, fila in df_cc.iterrows():
        parsed = parsear_checklist_texto(fila.get("Checklist_cierre", ""))
        for paso, estado in parsed.items():
            if paso not in estados:
                estados[paso] = estado
            else:
                actual = estados.get(paso, "Pendiente")
                if prioridad.get(estado, 0) >= prioridad.get(actual, 0):
                    estados[paso] = estado
    return estados


def calcular_avance_checklist(estados):
    if not estados:
        return 0, 0, 0, 0
    total = len(estados)
    cumplidos = sum(1 for e in estados.values() if e == "Cumplido")
    no_aplica = sum(1 for e in estados.values() if e == "No aplica")
    pendientes = sum(1 for e in estados.values() if e == "Pendiente")
    avance = round(((cumplidos + no_aplica) / total) * 100) if total else 0
    return avance, cumplidos, no_aplica, pendientes


def estado_visual_por_avance(avance, pendientes, estado_institucional):
    estado_inst = str(estado_institucional).lower()
    if "cerrado" in estado_inst or (avance == 100 and pendientes == 0):
        return "Cerrado"
    if pendientes > 0 and avance >= 50:
        return "En seguimiento"
    if avance == 0:
        return "Abierto"
    return "En proceso"


def construir_resumen_gestion_casos(df_reg):
    if df_reg.empty or "CC" not in df_reg.columns:
        return pd.DataFrame()

    filas = []
    for cc, df_cc in df_reg.groupby("CC"):
        if not str(cc).strip():
            continue
        df_cc = df_cc.copy()
        primera = df_cc.iloc[0]
        ultima = df_cc.iloc[-1]
        estados = consolidar_checklist_por_cc(df_cc)
        avance, cumplidos, no_aplica, pendientes = calcular_avance_checklist(estados)
        estado_institucional = str(ultima.get("Estado_institucional", ultima.get("Estado_caso", ""))).strip()
        estado_visual = estado_visual_por_avance(avance, pendientes, estado_institucional)
        filas.append({
            "CC": cc,
            "Estudiante": str(primera.get("Estudiante", "")).strip(),
            "Curso": str(primera.get("Curso", "")).strip(),
            "Estado": estado_visual,
            "Estado institucional": estado_institucional,
            "Avance %": avance,
            "Cumplidos": cumplidos,
            "No aplica": no_aplica,
            "Pendientes": pendientes,
            "Última acción": str(ultima.get("Tipo_registro", "Registro")).strip(),
            "Fecha última": str(ultima.get("Fecha_entrevista", "")).strip(),
            "Motivo": str(primera.get("Motivo", "")).strip(),
        })
    return pd.DataFrame(filas)


def render_checklist_consolidado(estados):
    for paso, estado in estados.items():
        if estado == "Cumplido":
            st.success(f"✅ {paso}: Cumplido")
        elif estado == "No aplica":
            st.info(f"🚫 {paso}: No aplica")
        else:
            st.warning(f"⏳ {paso}: Pendiente")


def generar_resumen_estado_actual(df_cc, estados):
    if df_cc.empty:
        return "No existen registros asociados al caso."
    ultima = df_cc.iloc[-1]
    avance, cumplidos, no_aplica, pendientes = calcular_avance_checklist(estados)
    estudiante = str(ultima.get("Estudiante", "")).strip()
    curso = str(ultima.get("Curso", "")).strip()
    estado = str(ultima.get("Estado_institucional", ultima.get("Estado_caso", ""))).strip()
    protocolos = str(ultima.get("Protocolos_aplicados", "")).strip()
    texto = (
        f"El caso asociado al estudiante {estudiante}, curso {curso}, presenta un avance institucional de {avance}%. "
        f"Actualmente se encuentra en estado {estado if estado else 'en proceso de revisión'}. "
    )
    if protocolos:
        texto += f"Se registran protocolos o acciones asociadas: {protocolos}. "
    if pendientes:
        texto += f"Permanecen {pendientes} acción(es) pendiente(s) de seguimiento o cierre."
    else:
        texto += "No se observan acciones pendientes en el checklist consolidado del caso."
    return texto




# ==================================================
# IA MIXTA: LOCAL PARA SENSIBLE / EXTERNA PARA GENERAL
# ==================================================

def obtener_config_ia():
    """
    Configuración de IA.
    - OLLAMA_MODEL: por defecto llama3.1:8b
    - GEMINI_API_KEY: desde variable de entorno o st.secrets
    """
    ollama_model = os.environ.get("OLLAMA_MODEL", "llama3:8b")
    gemini_model = os.environ.get("GEMINI_MODEL", "gemini-1.5-flash")

    gemini_key = os.environ.get("GEMINI_API_KEY", "")

    try:
        if not gemini_key and "GEMINI_API_KEY" in st.secrets:
            gemini_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass

    return {
        "ollama_model": ollama_model,
        "gemini_model": gemini_model,
        "gemini_key": gemini_key,
    }


def construir_prompt_redaccion(texto, contexto="general", tarea="acuerdos", sensible=True):
    """
    Corrector conservador: NO genera información nueva.
    """
    contexto_txt = {
        "hp": "entrevista de hogar/apoderado respecto de un estudiante",
        "e": "entrevista directa a estudiante",
        "compromiso_hp": "compromisos asumidos por hogar/apoderado respecto del estudiante",
        "compromiso_e": "compromisos asumidos por estudiante",
        "antecedentes_hp": "antecedentes entregados por hogar/apoderado respecto de un estudiante",
        "antecedentes_e": "antecedentes declarados por estudiante",
        "general": "redacción institucional escolar general"
    }.get(contexto, "redacción institucional escolar general")

    return f"""
Eres un corrector de redacción institucional escolar chilena.

NO eres generador de contenido.
NO debes crear acuerdos, compromisos ni antecedentes nuevos.

Tu tarea es SOLO corregir ortografía, puntuación, claridad y tono institucional, manteniendo exactamente la información escrita.

CONTEXTO:
{contexto_txt}

REGLAS OBLIGATORIAS:
1. Conserva exactamente la intención del texto original.
2. No agregues información nueva.
3. No elimines información escrita por el usuario.
4. No cambies una idea por otra.
5. No resumas.
6. No fusiones líneas.
7. Mantén cada línea o viñeta como un punto independiente.
8. Si el usuario entrega 4 líneas, devuelve 4 viñetas.
9. Si aparece "psicólogo", debe aparecer "psicólogo".
10. Si aparece "PIE", debe aparecer "PIE".
11. Si aparece "apoderado", debe aparecer "apoderado".
12. Si aparece "seguimiento", debe aparecer "seguimiento".
13. Si aparece "derivación", debe aparecer "derivación".
14. Si aparece "protocolo", debe aparecer "protocolo".
15. No uses la palabra "usuario".
16. Devuelve SOLO viñetas.

TEXTO ORIGINAL:
{texto}
""".strip()

def llamar_ollama(prompt, timeout=90):
    """
    Llama a Ollama local. No envía datos a internet.
    """
    cfg = obtener_config_ia()
    payload = {
        "model": cfg["ollama_model"],
        "messages": [
            {"role": "system", "content": "Eres un asistente institucional escolar. Redactas con precisión, sin inventar hechos."},
            {"role": "user", "content": prompt}
        ],
        "stream": False
    }

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        "http://localhost:11434/api/chat",
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST"
    )

    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read().decode("utf-8")
        result = json.loads(raw)
        return result.get("message", {}).get("content", "").strip()


def llamar_gemini(prompt, timeout=60):
    """
    Llama a Gemini API para tareas generales o anonimizadas.
    No debe usarse con datos sensibles reales.
    """
    cfg = obtener_config_ia()
    api_key = cfg["gemini_key"]

    if not api_key:
        raise RuntimeError("No hay GEMINI_API_KEY configurada.")

    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        + cfg["gemini_model"]
        + ":generateContent?key="
        + api_key
    )

    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt}
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.35,
            "topP": 0.85,
            "maxOutputTokens": 900
        }
    }

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST"
    )

    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read().decode("utf-8")
        result = json.loads(raw)
        candidates = result.get("candidates", [])
        if not candidates:
            return ""
        parts = candidates[0].get("content", {}).get("parts", [])
        return "\n".join([p.get("text", "") for p in parts]).strip()


def contiene_datos_sensibles(texto):
    """
    Detector preventivo simple.
    Para entrevistas reales, se recomienda considerar sensible por defecto.
    """
    texto = str(texto).lower()
    claves = [
        "apoderado", "estudiante", "curso", "rut", "run", "autolesión",
        "autoagresión", "vulneración", "familia", "madre", "padre",
        "psicólogo", "asistente social", "tribunal", "fiscalía",
        "cesfam", "opd", "nombre", "7a", "7b", "6a", "6b"
    ]
    return any(c in texto for c in claves)



def palabras_clave_minimas(texto):
    base = corregir_ortografia_basica(str(texto).lower())
    claves = [
        "psicólogo", "psicologa", "seguimiento", "apoderado", "estudiante",
        "familia", "protocolo", "derivación", "derivacion", "pie",
        "profesor", "profesora", "dirección", "convivencia", "orientación",
        "utp", "inspectoría", "mediación", "conciliación", "autolesión",
        "riesgo", "evidencia", "citación", "monitoreo", "académico",
        "emocional", "conductual"
    ]
    salida = []
    for c in claves:
        if c in base and c not in salida:
            salida.append(c)
    return salida


def respuesta_conserva_claves(original, respuesta):
    claves = palabras_clave_minimas(original)
    resp = corregir_ortografia_basica(str(respuesta).lower())
    equivalencias = {"psicologa": "psicólogo", "derivacion": "derivación"}
    for c in claves:
        cn = equivalencias.get(c, c)
        if cn not in resp:
            return False
    return True


def fallback_conservador(linea, contexto="general"):
    linea = corregir_ortografia_basica(str(linea).strip(" •-*"))
    if not linea:
        return ""
    contenido = linea[0].lower() + linea[1:] if len(linea) > 1 else linea
    if contexto == "compromiso_hp":
        return f"• La familia se compromete a {contenido}."
    if contexto == "compromiso_e":
        return f"• El estudiante se compromete a {contenido}."
    if contexto in ["antecedentes_hp", "hp"]:
        return f"• Se consigna que {contenido}."
    if contexto in ["antecedentes_e", "e"]:
        return f"• El estudiante señala que {contenido}."
    return f"• {linea}."


def mejorar_linea_conservadora(linea, contexto="general", tarea="acuerdos", sensible=True):
    linea = corregir_ortografia_basica(str(linea).strip(" •-*"))
    if not linea:
        return ""
    prompt = construir_prompt_redaccion(f"• {linea}", contexto=contexto, tarea=tarea, sensible=sensible)
    try:
        if sensible or contiene_datos_sensibles(linea):
            respuesta = llamar_ollama(prompt)
        else:
            cfg = obtener_config_ia()
            respuesta = llamar_gemini(prompt) if cfg.get("gemini_key") else llamar_ollama(prompt)
        respuesta = normalizar_salida_ia(respuesta)
        primera = next((x.strip() for x in respuesta.splitlines() if x.strip()), "")
        if primera and respuesta_conserva_claves(linea, primera):
            return primera
    except Exception:
        pass
    return fallback_conservador(linea, contexto=contexto)


def mejorar_texto_conservador_por_lineas(texto, contexto="general", tarea="acuerdos", sensible=True):
    lineas = separar_acuerdos(texto)
    salida = []
    for linea in lineas:
        m = mejorar_linea_conservadora(linea, contexto=contexto, tarea=tarea, sensible=sensible)
        if m:
            salida.append(m)
    return "\\n".join(salida)



def dividir_lineas_conservadoras(texto):
    texto = corregir_ortografia_basica(str(texto).strip())
    if not texto:
        return []
    lineas = []
    for raw in texto.splitlines():
        item = raw.strip().strip("•-–—* ").strip()
        if not item:
            continue
        if len(item) > 2 and item[0].isdigit() and item[1] in [".", ")"]:
            item = item[2:].strip()
        if item:
            lineas.append(item)
    return lineas if lineas else [texto]


def palabras_clave_minimas(texto):
    base = corregir_ortografia_basica(str(texto).lower())
    claves = ["psicólogo","psicologa","seguimiento","apoderado","estudiante","familia","protocolo","derivación","derivacion","pie","profesor","profesora","dirección","convivencia","orientación","utp","inspectoría","mediación","conciliación","autolesión","riesgo","evidencia","citación","monitoreo","académico","emocional","conductual"]
    salida = []
    for c in claves:
        if c in base and c not in salida:
            salida.append(c)
    return salida


def respuesta_conserva_claves(original, respuesta):
    resp = corregir_ortografia_basica(str(respuesta).lower())
    equivalencias = {"psicologa": "psicólogo", "derivacion": "derivación"}
    for c in palabras_clave_minimas(original):
        if equivalencias.get(c, c) not in resp:
            return False
    return True


def fallback_redaccion_conservadora(linea, contexto="general"):
    linea = corregir_ortografia_basica(str(linea).strip(" •-*"))
    if not linea:
        return ""
    contenido = linea[0].lower() + linea[1:] if len(linea) > 1 else linea
    if contexto == "compromiso_hp":
        return f"• La familia se compromete a {contenido}."
    if contexto == "compromiso_e":
        return f"• El estudiante se compromete a {contenido}."
    if contexto in ["antecedentes_hp", "hp"]:
        return f"• Se consigna que {contenido}."
    if contexto in ["antecedentes_e", "e"]:
        return f"• El estudiante señala que {contenido}."
    return f"• {linea}."


def mejorar_linea_redaccion_local(linea, contexto="general", tarea="acuerdos", sensible=True):
    linea = corregir_ortografia_basica(str(linea).strip(" •-*"))
    if not linea:
        return ""
    prompt = construir_prompt_redaccion(f"• {linea}", contexto=contexto, tarea=tarea, sensible=sensible)
    try:
        respuesta = normalizar_salida_ia(llamar_ollama(prompt))
        primera = next((x.strip() for x in respuesta.splitlines() if x.strip()), "")
        if primera and respuesta_conserva_claves(linea, primera):
            return primera if primera.startswith("•") else "• " + primera
    except Exception:
        pass
    return fallback_redaccion_conservadora(linea, contexto=contexto)


def mejorar_redaccion_conservadora(texto, contexto="general", tarea="acuerdos", sensible=True):
    salida = []
    for linea in dividir_lineas_conservadoras(texto):
        mejorada = mejorar_linea_redaccion_local(linea, contexto=contexto, tarea=tarea, sensible=True)
        if mejorada:
            salida.append(mejorada)
    return "\n".join(salida)


def mejorar_texto_mixto(texto, contexto="general", tarea="acuerdos", sensible=True, permitir_externa=False):
    """
    Redacción local sensible.
    No genera acuerdos, antecedentes ni compromisos.
    Solo mejora lo ingresado por el usuario, línea por línea.
    """
    texto = str(texto).strip()
    if not texto:
        return ""
    return mejorar_redaccion_conservadora(
        texto,
        contexto=contexto,
        tarea=tarea,
        sensible=True
    )


def normalizar_salida_ia(texto):
    """
    Asegura salida con viñetas limpias.
    """
    lineas = []
    for raw in str(texto).split("\n"):
        item = raw.strip()
        if not item:
            continue
        item = item.strip("-•* ").strip()
        if len(item) > 2 and item[0].isdigit() and item[1] in [".", ")"]:
            item = item[2:].strip()
        if item:
            lineas.append(f"• {item}")

    return "\n".join(lineas)


def panel_estado_ia():
    """
    Panel informativo breve para Administración.
    """
    cfg = obtener_config_ia()
    st.markdown("### Configuración IA mixta: ial / iae")
    st.write(f"ial local Ollama: `{cfg['ollama_model']}`")
    st.write(f"iae externa Gemini: `{cfg['gemini_model']}`")
    if cfg["gemini_key"]:
        st.success("GEMINI_API_KEY detectada.")
    else:
        st.info("GEMINI_API_KEY no configurada. La app usará Ollama local y reglas internas.")




# ==================================================
# EDITOR DE TEXTO TIPO WORD / QUILL
# ==================================================

def normalizar_html_a_texto(valor):
    import re as _re
    texto = str(valor or "")
    texto = texto.replace("</p><p>", "\n")
    texto = texto.replace("<br>", "\n").replace("<br/>", "\n").replace("<br />", "\n")
    texto = texto.replace("</li><li>", "\n• ")
    texto = texto.replace("<li>", "• ").replace("</li>", "")
    texto = texto.replace("<ul>", "").replace("</ul>", "")
    texto = texto.replace("<ol>", "").replace("</ol>", "")
    texto = texto.replace("</p>", "\n").replace("<p>", "")
    texto = _re.sub(r"<[^>]+>", "", texto)
    texto = texto.replace("&nbsp;", " ").replace("&amp;", "&")
    texto = texto.replace("&lt;", "<").replace("&gt;", ">")
    lineas = [" ".join(x.strip().split()) for x in texto.split("\n")]
    return "\n".join([x for x in lineas if x])


def editor_texto_word(label, key, value="", height=260, help_text=""):
    """
    Editor visual tipo Word usando streamlit-quill.
    Diseñado para quedar visible en Nuevo e Histórico.
    """
    st.markdown(f"**{label}**")

    if help_text:
        st.caption(help_text)

    # Espaciador mínimo para evitar que Streamlit colapse el iframe al cargar.
    st.markdown(
        """
        <style>
        iframe {
            min-height: 285px !important;
            height: 285px !important;
            display: block !important;
            visibility: visible !important;
            opacity: 1 !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    if QUILL_DISPONIBLE:
        toolbar_config = [
            ["bold", "italic", "underline"],
            [{"list": "ordered"}, {"list": "bullet"}],
            ["clean"]
        ]

        contenido = st_quill(
            value=value or "",
            html=True,
            key=key,
            toolbar=toolbar_config,
            placeholder="Escriba aquí..."
        )

        return normalizar_html_a_texto(contenido)

    st.error("No está instalado streamlit-quill en el PC servidor. Ejecuta: pip install streamlit-quill")
    return st.text_area(
        label,
        value=value,
        height=height,
        key=key + "_fallback",
        help=help_text
    )


# ==================================================
# HERRAMIENTAS SIMPLES DE TEXTO
# ==================================================

def limpiar_formato_texto(texto):
    """
    Limpia espacios, líneas vacías repetidas y caracteres básicos.
    """
    lineas = []
    for linea in str(texto).replace("\r", "\n").split("\n"):
        limpia = " ".join(linea.strip().split())
        if limpia:
            lineas.append(limpia)
    return "\n".join(lineas)


def convertir_a_vinetas(texto):
    """
    Convierte cada línea o frase en viñeta.
    """
    partes = separar_acuerdos(texto) if "separar_acuerdos" in globals() else str(texto).split("\n")
    salida = []
    for item in partes:
        item = str(item).strip(" •-*0123456789.").strip()
        if item:
            salida.append(f"• {item}")
    return "\n".join(salida)


def ordenar_texto_institucional(texto, contexto="general"):
    """
    Ordena texto sin inventar antecedentes.
    """
    texto = corregir_ortografia_basica(str(texto).strip()) if "corregir_ortografia_basica" in globals() else str(texto).strip()

    if not texto:
        return ""

    lineas = [x.strip() for x in texto.split("\n") if x.strip()]

    if len(lineas) > 1:
        cuerpo = "\n".join([f"• {x.strip('•- ')}" for x in lineas])
    else:
        cuerpo = texto

    if contexto == "hp":
        return "Antecedentes informados por el/la apoderado/a respecto de su estudiante:\n" + cuerpo
    if contexto == "e":
        return "Antecedentes declarados por el/la estudiante:\n" + cuerpo

    return cuerpo


def herramientas_texto(texto_base, prefix, contexto="general", sensible=True):
    """
    Panel de edición rápida.
    No modifica directamente el widget original; muestra resultado en un cuadro separado.
    """
    st.markdown("##### Herramientas de texto")

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        if st.button("Ordenar texto", key=f"{prefix}_ordenar_texto"):
            st.session_state[f"{prefix}_texto_editado"] = ordenar_texto_institucional(texto_base, contexto=contexto)
            st.rerun()

    with c2:
        if st.button("Corregir ortografía", key=f"{prefix}_corregir_texto"):
            st.session_state[f"{prefix}_texto_editado"] = corregir_ortografia_basica(texto_base)
            st.rerun()

    with c3:
        if st.button("Convertir a viñetas", key=f"{prefix}_vinetas_texto"):
            st.session_state[f"{prefix}_texto_editado"] = convertir_a_vinetas(texto_base)
            st.rerun()

    with c4:
        if st.button("Limpiar texto", key=f"{prefix}_limpiar_texto"):
            st.session_state[f"{prefix}_texto_editado"] = limpiar_formato_texto(texto_base)
            st.rerun()

    if st.button("Mejorar redacción", key=f"{prefix}_mejorar_redaccion"):
        st.session_state[f"{prefix}_texto_editado"] = mejorar_texto_mixto(
            texto_base,
            contexto=contexto,
            tarea="antecedentes",
            sensible=sensible
        )
        st.rerun()

    if st.session_state.get(f"{prefix}_texto_editado"):
        return st.text_area(
            "Texto editado sugerido",
            value=st.session_state[f"{prefix}_texto_editado"],
            height=150,
            key=f"{prefix}_texto_editado_visible"
        )

    return texto_base


# ==================================================
# CARGA HISTÓRICA DE CASOS
# ==================================================

def selector_tipo_registro_historico(prefix="general"):
    tipo_ingreso = st.radio(
        "Modo de registro",
        ["Nuevo", "Histórico"],
        horizontal=True,
        key=f"{prefix}_modo_registro"
    )
    return tipo_ingreso


def datos_historicos(prefix="general"):
    st.info("Modo histórico activado: ingresa los datos reales ya existentes del caso.")

    col1, col2, col3 = st.columns(3)

    with col1:
        cc_manual = st.text_input(
            "CC existente",
            key=f"{prefix}_cc_historico",
            help="Ejemplo: CC-2026-0012"
        )

    with col2:
        folio_manual = st.text_input(
            "Folio histórico",
            key=f"{prefix}_folio_historico"
        )

    with col3:
        entrevista_manual = st.text_input(
            "N° entrevista histórica",
            key=f"{prefix}_entrevista_historica",
            help="Ejemplo: NE-CC-2026-0001-001 o número usado previamente."
        )

    fecha_manual = st.date_input(
        "Fecha histórica",
        key=f"{prefix}_fecha_historica"
    )

    return cc_manual, folio_manual, entrevista_manual, fecha_manual



# ==================================================
# LOGIN FLOW
# ==================================================

if "logueado" not in st.session_state:
    st.session_state.logueado = False
    st.session_state.usuario = None
    st.session_state.nombre = None
    st.session_state.rol = None

if not st.session_state.logueado:
    encabezado("App Convivencia Pukaray", "Acceso institucional")

    with st.form("form_login_institucional"):
        usuario = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        entrar = st.form_submit_button("Ingresar")

    if entrar:
        data = autenticar(usuario, password)
        if data:
            st.session_state.logueado = True
            st.session_state.usuario = usuario
            st.session_state.nombre = data[0]
            st.session_state.rol = data[1]
            st.rerun()
        else:
            st.error("Credenciales incorrectas.")

    st.caption("Acceso restringido a funcionarios autorizados del Colegio Pukaray.")
    st.stop()


# ==================================================
# MENÚ
# ==================================================

st.sidebar.markdown("### 🌳 Convivencia Pukaray")
st.sidebar.caption(f"Usuario: {st.session_state.nombre}")
st.sidebar.caption(f"Rol: {st.session_state.rol}")
st.sidebar.caption("Acceso normativo")
boton_notebooklm_normativo()
st.sidebar.caption("IA mixta: ial sensible / iae general")
st.sidebar.divider()
boton_notebooklm_normativo()

if st.session_state.get("rol") == "Administrador":
    opciones_menu = [
        "Inicio",
        "Hogar / Apoderado",
        "Entrevista Estudiante",
        "Gestión de Casos",
        "Registros",
        "Administración",
        "Salir"
    ]
else:
    opciones_menu = [
        "Inicio",
        "Hogar / Apoderado",
        "Entrevista Estudiante",
        "Gestión de Casos",
        "Salir"
    ]

menu = st.sidebar.radio(
    "Menú",
    opciones_menu
)

if menu == "Salir":
    st.session_state.logueado = False
    st.rerun()


# ==================================================
# PÁGINAS
# ==================================================

if menu == "Inicio":
    encabezado("App Convivencia Pukaray", "Fase 1: Hogar / Apoderado estable")

    st.markdown("""
    <div class="pukaray-card">
    <span class="pukaray-badge">Sistema institucional</span><br>
    <b>Plataforma de gestión de convivencia escolar y seguimiento de casos.</b><br><br>
    ✓ Entrevistas hp y e con ficha imprimible institucional.<br>
    ✓ Código de caso <b>cc</b>, folio y número de entrevista.<br>
    ✓ Protocolos, checklist, condición del caso y cierre institucional.<br>
    ✓ Gestión de casos con avance, línea de tiempo y expediente descargable.<br>
    ✓ Redacción técnica con enfoque RICE y revisión básica de ortografía.<br>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Estudiantes", len(df_estudiantes))
    c2.metric("Entrevistadores", len(obtener_entrevistadores()))
    c3.metric("Departamentos", len(obtener_departamentos()))
    c4.metric("Tipos de falta", max(len(obtener_tipo_falta_categoria_entorno_tipo()) - 1, 0))


elif menu == "Hogar / Apoderado":
    encabezado("Hogar / Apoderado", "Ficha de entrevista de apoderado")


    tipo_registro_hist = selector_tipo_registro_historico("hp")


    st.subheader("Datos básicos")

    cc = selector_cc("hogar")

    if tipo_registro_hist == "Histórico":
        cc_manual, folio_manual, entrevista_manual, fecha_manual = datos_historicos("hp")

        if cc_manual.strip():
            cc = cc_manual.strip()


    col_folio, col_ne = st.columns(2)
    with col_folio:
        folio = selector_correlativo("hogar", "Folio", "F", "Folio")
    with col_ne:
        numero_entrevista = selector_correlativo("hogar", "N° entrevista", "NE", "Numero_entrevista", cc=cc)

    col1, col2 = st.columns(2)

    with col1:
        fecha = st.date_input("Fecha", value=date.today(), key="fecha_hogar")
        hora = st.time_input("Hora", value=datetime.now().time(), key="hora_hogar")
        estudiante, curso, apoderado_base, run = seleccionar_estudiante("hogar")

    with col2:
        departamento_sel = st.multiselect(
            "Departamento que cita",
            obtener_departamentos(),
            key="departamento_hogar",
            placeholder="Seleccione..."
        )
        departamento = texto_lista(departamento_sel)

        participantes, cargo_auto = seleccionar_participantes("hogar")

        cargo_entrevistador = cargo_auto

    st.subheader("Persona entrevistada")

    col3, col4, col5 = st.columns(3)

    with col3:
        apoderado = st.text_area(
            "Nombre(s) apoderado(s) / persona(s) entrevistada(s)",
            value=apoderado_base,
            height=90,
            key="apoderado_hogar",
            help="Puedes ingresar uno o más apoderados, uno por línea si corresponde."
        )

    with col4:
        vinculo = st.selectbox(
            "Vínculo",
            opciones_con_seleccione(VINCULOS),
            key="vinculo_hogar"
        )
        vinculo = limpiar_seleccione(vinculo)

    with col5:
        asiste_apoderado = st.selectbox(
            "Asiste apoderado",
            ["Sí", "No"],
            key="asiste_apoderado"
        )
        asiste_estudiante = st.selectbox(
            "Asiste estudiante",
            ["No", "Sí"],
            key="asiste_estudiante"
        )

    st.subheader("Clasificación del registro")

    motivo_sel = st.multiselect(
        "Motivo de la entrevista",
        obtener_motivos(),
        key="motivo_hogar",
        placeholder="Seleccione un motivo de entrevista..."
    )
    motivo = texto_lista(motivo_sel)
    tipo_registro_rice = ""
    st.caption("Seleccione el motivo principal que origina esta entrevista.")

    tipo_falta_sel = st.multiselect(
        "Tipo Falta RICE",
        obtener_tipo_falta_categoria_entorno_tipo(),
        key="tipo_falta_hogar",
        placeholder="Seleccione un tipo de falta RICE..."
    )
    tipo_falta = texto_lista(tipo_falta_sel)
    st.caption("Orden visible: Tipo | Entorno | Categoría.")

    protocolos_sel = st.multiselect(
        "Protocolos aplicados",
        obtener_protocolos_formateados(),
        key="protocolos_hogar",
        placeholder="Seleccione los protocolos aplicados..."
    )
    protocolos_aplicados = texto_lista(protocolos_sel)
    st.caption("Orden visible: Detalle | Protocolo | Categoría.")

    estado_caso = st.selectbox(
        "Estado del caso",
        opciones_con_seleccione(obtener_estados_caso()),
        key="estado_caso_hogar"
    )
    estado_caso = limpiar_seleccione(estado_caso)

    condicion_caso = st.selectbox(
        "Condición del caso",
        opciones_con_seleccione(["General", "Reservado", "Altamente sensible"]),
        key="condicion_caso_hogar",
        help="Personalizado queda reservado para administración."
    )
    condicion_caso = limpiar_seleccione(condicion_caso)

    checklist_cierre, estado_sugerido_checklist, pendientes_checklist = checklist_cierre_interactivo("hogar")
    estado_institucional, detalle_estado, conclusion_estado = selector_estado_detalle_caso("hogar")

    st.subheader("Desarrollo de la entrevista")

    bloque_levantamiento_ocr_formulario("hp_form_ocr", destino="hp")


    st.markdown("#### Apoyo normativo")
    st.caption("Consulta RICE, protocolos y normativa. No pegues datos sensibles de estudiantes en NotebookLM.")
    boton_notebooklm_normativo()



    antecedentes = editor_texto_word(
        "Antecedentes / relato de hechos",
        key="hp_antecedentes_editor",
        value=st.session_state.get("antecedentes_ordenados_hogar", st.session_state.get("hp_antecedentes_valor", "")),
        height=220,
        help_text="Use la barra para negrita, cursiva, subrayado, viñetas y numeración."
    )

    datos_previos = {
        "cc": cc,
        "folio": folio,
        "numero_entrevista": numero_entrevista,
        "resumen_caso": f"{estudiante} | {curso} | {motivo} | {estado_caso}",
        "fecha": fecha.strftime("%d-%m-%Y"),
        "hora": str(hora)[:5],
        "departamento": departamento,
        "curso": curso,
        "estudiante": estudiante,
        "run": run,
        "apoderado": apoderado,
        "vinculo": vinculo,
        "entrevistador": participantes,
        "participantes": participantes,
        "cargo_entrevistador": cargo_entrevistador or cargo_por_entrevistadores(participantes),
        "motivo": motivo,
        "tipo_registro_rice": tipo_registro_rice,
        "tipo_falta": tipo_falta,
        "protocolos_aplicados": protocolos_aplicados,
        "estado_caso": estado_caso,
        "condicion_caso": condicion_caso,
        "checklist_cierre": checklist_cierre,
        "pendientes_checklist": texto_lista(pendientes_checklist),
        "estado_sugerido_checklist": estado_sugerido_checklist,
        "estado_institucional": estado_institucional,
        "detalle_estado": detalle_estado,
        "conclusion_estado": conclusion_estado,
        "antecedentes": antecedentes,
        "asiste_apoderado": asiste_apoderado,
        "asiste_estudiante": asiste_estudiante,
    }

    if st.button("Generar resumen libro", key="btn_resumen_base"):
        st.session_state["resumen_hogar"] = generar_resumen_libro(datos_previos)
        st.rerun()

    antecedentes_para_word = antecedentes

    if st.session_state.get("antecedentes_ordenados_hogar"):
        antecedentes_para_word = editor_texto_word(
            "Antecedentes ordenados sugeridos",
            key="hp_antecedentes_ordenados_editor",
            value=st.session_state["antecedentes_ordenados_hogar"],
            height=220,
            help_text="Texto sugerido. Puede editarlo antes de guardar o generar Word."
        )

    acuerdos_raw = editor_texto_word(
        "Acuerdos o conclusiones",
        key="hp_acuerdos_editor",
        value=st.session_state.get("acuerdos_hogar", ""),
        height=220,
        help_text="Ingrese un acuerdo por viñeta o línea."
    )
    acuerdos_lista = []
    for linea in acuerdos_raw.split("\n"):
        linea = linea.strip()
        if linea:
            if linea.startswith("•") or (len(linea) > 2 and linea[0].isdigit() and linea[1] in [".", ")"]):
                acuerdos_lista.append(linea)
            else:
                acuerdos_lista.append(f"• {linea}")

    acuerdos = "\n".join(acuerdos_lista)

    compromisos_raw = editor_texto_word(
        "Compromisos hogar/apoderado",
        key="hp_compromisos_editor",
        value=st.session_state.get("compromisos_hogar_mejorado", ""),
        height=200,
        help_text="Ingrese un compromiso por viñeta o línea."
    )

    compromisos = compromisos_raw

    observaciones = st.text_area(
        "Observaciones internas",
        height=90,
        key="observaciones_hogar"
    )

    resumen_libro = st.text_area(
        "Resumen para libro de clases",
        value=st.session_state.get("resumen_hogar", ""),
        height=120,
        key="resumen_hogar",
        help="Este resumen NO se incorpora al Word. Es solo para copiar y pegar manualmente."
    )

    motivos_formateados = []
    if motivo:
        for item in motivo.split(","):
            item = item.strip()
            if item:
                motivos_formateados.append(f"• {item}")

    faltas_formateadas = []
    if tipo_falta:
        for item in tipo_falta.split(","):
            item = item.strip()
            if item:
                faltas_formateadas.append(f"• {item}")

    motivo_word = (
        f"CÓDIGO DE CASO (cc): {cc}\n"
        + f"Folio: {folio}\n"
        + f"N° entrevista: {numero_entrevista}\n\n"
        + "MOTIVOS DE LA ENTREVISTA:\n"
        + ("\n".join(motivos_formateados) if motivos_formateados else "• No informado")
        + "\n\n"
        + "TIPO FALTA RICE:\n"
        + ("\n".join(faltas_formateadas) if faltas_formateadas else "• No aplica")
        + "\n\n"
        + "PROTOCOLOS APLICADOS:\n"
        + (protocolos_aplicados if protocolos_aplicados else "No aplica")
        + "\n\n"
        + "CONDICIÓN DEL CASO:\n"
        + condicion_caso
        + "\n\n"
        + "ANTECEDENTES:\n"
        + (antecedentes_para_word or antecedentes or "No informado")
    )

    datos_previos["antecedentes"] = antecedentes_para_word or antecedentes

    datos = {
        **datos_previos,
        "acuerdos": acuerdos,
        "compromisos": compromisos,
        "observaciones": observaciones,
        "resumen_libro": resumen_libro,
        "motivo_word": motivo_word,
    }

    st.divider()

    col_guardar, col_word = st.columns(2)

    with col_guardar:
        if st.button("Guardar registro mínimo", key="guardar_hogar"):
            if not estudiante:
                st.error("Debes seleccionar estudiante.")
            elif not apoderado:
                st.error("Debes ingresar apoderado/persona entrevistada.")
            elif datos.get("estado_caso") == "Cerrado" and datos.get("pendientes_checklist"):
                st.error("No se puede registrar como Cerrado porque existen pasos pendientes en el checklist.")
            else:
                guardar_registro(datos)

    with col_word:
        archivo_word = rellenar_docx_por_tablas(datos)
        nombre_archivo = limpiar_nombre_archivo(
            f"Ficha_Apoderado_{estudiante or 'estudiante'}_{curso}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        )

        st.download_button(
            "Descargar Word ficha apoderado",
            data=archivo_word,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="descargar_word_hogar"
        )

    panel_registros_en_formulario("hogar")


    st.markdown("### Resumen para copiar")
    st.caption("No se incluye en el Word.")
    st.code(resumen_libro or "Sin resumen generado.")


elif menu == "Entrevista Estudiante":
    encabezado("Entrevista Estudiante", "Registro institucional de entrevista estudiantil")


    tipo_registro_hist = selector_tipo_registro_historico("e")


    st.subheader("Datos básicos")

    cc = selector_cc("estudiante")

    if tipo_registro_hist == "Histórico":
        cc_manual, folio_manual, entrevista_manual, fecha_manual = datos_historicos("e")

        if cc_manual.strip():
            cc = cc_manual.strip()


    col_folio, col_ne = st.columns(2)
    with col_folio:
        folio = selector_correlativo("estudiante", "Folio", "F", "Folio")
    with col_ne:
        numero_entrevista = selector_correlativo("estudiante", "N° entrevista", "NE", "Numero_entrevista", cc=cc)

    col1, col2 = st.columns(2)

    with col1:
        fecha = st.date_input("Fecha", value=date.today(), key="fecha_estudiante")
        hora = st.time_input("Hora", value=datetime.now().time(), key="hora_estudiante")
        estudiante, curso, apoderado_base, run = seleccionar_estudiante("estudiante")

    with col2:
        departamento_sel = st.multiselect(
            "Departamento que cita",
            obtener_departamentos(),
            key="departamento_estudiante",
            placeholder="Seleccione..."
        )
        departamento = texto_lista(departamento_sel)

        participantes, cargo_auto = seleccionar_participantes("estudiante")

        cargo_entrevistador = cargo_auto

    st.subheader("Clasificación del registro")

    tipo_relato = st.selectbox(
        "Tipo de entrevista",
        [
            "Relato espontáneo",
            "Orientación",
            "Intervención",
            "Entrevista Estudiante",
            "Intervención académica / seguimiento pedagógico",
            "Declaración estudiante",
            "Seguimiento convivencia",
            "Derivación docente",
            "Entrevista preventiva",
            "Otro"
        ],
        key="tipo_relato_estudiante"
    )

    motivo_sel = st.multiselect(
        "Motivo de la entrevista",
        obtener_motivos(),
        key="motivo_estudiante",
        placeholder="Seleccione un motivo de entrevista..."
    )
    motivo = texto_lista(motivo_sel)
    st.caption("Seleccione el motivo principal que origina esta entrevista.")

    tipo_falta_sel = st.multiselect(
        "Tipo Falta RICE",
        obtener_tipo_falta_categoria_entorno_tipo(),
        key="tipo_falta_estudiante",
        placeholder="Seleccione un tipo de falta RICE..."
    )
    tipo_falta = texto_lista(tipo_falta_sel)
    st.caption("Orden visible: Tipo | Entorno | Categoría.")

    protocolos_sel = st.multiselect(
        "Protocolos aplicados",
        obtener_protocolos_formateados(),
        key="protocolos_estudiante",
        placeholder="Seleccione los protocolos aplicados..."
    )
    protocolos_aplicados = texto_lista(protocolos_sel)
    st.caption("Orden visible: Detalle | Protocolo | Categoría.")

    estado_caso = st.selectbox(
        "Estado del caso",
        opciones_con_seleccione(obtener_estados_caso()),
        key="estado_caso_estudiante"
    )
    estado_caso = limpiar_seleccione(estado_caso)

    condicion_caso = st.selectbox(
        "Condición del caso",
        opciones_con_seleccione(["General", "Reservado", "Altamente sensible"]),
        key="condicion_caso_estudiante",
        help="Personalizado queda reservado para administración."
    )
    condicion_caso = limpiar_seleccione(condicion_caso)

    checklist_cierre, estado_sugerido_checklist, pendientes_checklist = checklist_cierre_interactivo("estudiante")

    estado_institucional, detalle_estado, conclusion_estado = selector_estado_detalle_caso("estudiante")

    st.subheader("Desarrollo de la entrevista")

    bloque_levantamiento_ocr_formulario("e_form_ocr", destino="e")


    st.markdown("#### Apoyo normativo")
    st.caption("Consulta RICE, protocolos y normativa. No pegues datos sensibles de estudiantes en NotebookLM.")
    boton_notebooklm_normativo()



    antecedentes_para_word = ""

    antecedentes = editor_texto_word(
        "Antecedentes / relato de hechos",
        key="e_antecedentes_editor",
        value=st.session_state.get("antecedentes_ordenados_estudiante", st.session_state.get("e_antecedentes_valor", "")),
        height=220,
        help_text="Use la barra para negrita, cursiva, subrayado, viñetas y numeración."
    )

    datos_previos = {
        "tipo_registro": "Entrevista Estudiante",
        "tipo_relato": tipo_relato,
        "cc": cc,
        "folio": folio,
        "numero_entrevista": numero_entrevista,
        "resumen_caso": f"{estudiante} | {curso} | {motivo if 'motivo' in locals() else ''} | {estado_caso if 'estado_caso' in locals() else ''}",
        "fecha": fecha.strftime("%d-%m-%Y"),
        "hora": str(hora)[:5],
        "departamento": departamento,
        "curso": curso,
        "estudiante": estudiante,
        "run": run,
        "apoderado": "",
        "vinculo": "Estudiante",
        "entrevistador": participantes,
        "participantes": participantes,
        "cargo_entrevistador": cargo_entrevistador or cargo_por_entrevistadores(participantes),
        "motivo": motivo,
        "tipo_registro_rice": "",
        "tipo_falta": tipo_falta,
        "protocolos_aplicados": protocolos_aplicados,
        "estado_caso": estado_caso,
        "protocolos_aplicados": protocolos_aplicados,
        "checklist_cierre": checklist_cierre,
        "pendientes_checklist": texto_lista(pendientes_checklist),
        "estado_sugerido_checklist": estado_sugerido_checklist,
        "estado_institucional": estado_institucional,
        "detalle_estado": detalle_estado,
        "conclusion_estado": conclusion_estado,
        "antecedentes": antecedentes,
    }

    if st.button("Generar resumen libro", key="btn_resumen_base_estudiante"):
        st.session_state["resumen_estudiante"] = generar_resumen_libro_estudiante(datos_previos)
        st.rerun()

    antecedentes_para_word = antecedentes

    if st.session_state.get("antecedentes_ordenados_estudiante"):
        antecedentes_para_word = editor_texto_word(
            "Antecedentes ordenados sugeridos",
            key="e_antecedentes_ordenados_editor",
            value=st.session_state["antecedentes_ordenados_estudiante"],
            height=220,
            help_text="Texto sugerido. Puede editarlo antes de guardar o generar Word."
        )

    acuerdos = editor_texto_word(
        "Acuerdos o conclusiones",
        key="e_acuerdos_editor",
        value=st.session_state.get("acuerdos_estudiante", ""),
        height=220,
        help_text="Ingrese un acuerdo por viñeta o línea."
    )
    compromisos_raw = editor_texto_word(
        "Compromisos estudiante",
        key="e_compromisos_editor",
        value=st.session_state.get("compromisos_estudiante_mejorado", ""),
        height=200,
        help_text="Ingrese un compromiso por viñeta o línea."
    )

    compromisos = compromisos_raw

    observaciones = st.text_area(
        "Observaciones internas",
        height=90,
        key="observaciones_estudiante"
    )

    resumen_libro = st.text_area(
        "Resumen para libro de clases",
        height=120,
        key="resumen_estudiante",
        help="Este resumen NO se incorpora al Word. Es solo para copiar y pegar manualmente."
    )

    motivos_formateados = []
    if motivo:
        for item in motivo.split(","):
            item = item.strip()
            if item:
                motivos_formateados.append(f"• {item}")

    faltas_formateadas = []
    if tipo_falta:
        for item in tipo_falta.split(","):
            item = item.strip()
            if item:
                faltas_formateadas.append(f"• {item}")

    acuerdos_lista = []
    for linea in acuerdos.split("\n"):
        linea = linea.strip()
        if linea:
            if not linea.startswith("•"):
                acuerdos_lista.append(f"• {linea}")
            else:
                acuerdos_lista.append(linea)

    acuerdos_word = "\n".join(acuerdos_lista)

    motivo_word = (
        f"CÓDIGO DE CASO (cc): {cc}\n"
        + f"Folio: {folio}\n"
        + f"N° entrevista: {numero_entrevista}\n\n"
        + f"TIPO DE ENTREVISTA:\n• {tipo_relato}\n\n"
        + "MOTIVOS DE LA ENTREVISTA:\n"
        + ("\n".join(motivos_formateados) if motivos_formateados else "• No informado")
        + "\n\n"
        + "TIPO FALTA RICE:\n"
        + ("\n".join(faltas_formateadas) if faltas_formateadas else "• No aplica")
        + "\n\n"
        + "PROTOCOLOS APLICADOS:\n"
        + (protocolos_aplicados if protocolos_aplicados else "No aplica")
        + "\n\n"
        + "CONDICIÓN DEL CASO:\n"
        + condicion_caso
        + "\n\n"
        + "ANTECEDENTES:\n"
        + (locals().get("antecedentes_para_word", antecedentes) or "No informado")
    )

    datos_previos["antecedentes"] = antecedentes_para_word

    acuerdos_word_final = acuerdos if acuerdos.strip() else acuerdos_word

    datos = {
        **datos_previos,
        "acuerdos": acuerdos_word_final,
        "compromisos": compromisos,
        "observaciones": observaciones,
        "resumen_libro": resumen_libro,
        "motivo_word": motivo_word,
    }

    st.divider()

    col_guardar, col_word = st.columns(2)

    with col_guardar:
        if st.button("Guardar registro mínimo", key="guardar_estudiante"):
            if not estudiante:
                st.error("Debes seleccionar estudiante.")
            elif datos.get("estado_caso") == "Cerrado" and datos.get("pendientes_checklist"):
                st.error("No se puede registrar como Cerrado porque existen pasos pendientes en el checklist.")
            else:
                guardar_registro(datos)

    with col_word:
        archivo_word = rellenar_docx_estudiante(datos)
        nombre_archivo = limpiar_nombre_archivo(
            f"Ficha_Estudiante_{estudiante or 'estudiante'}_{curso}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        )

        st.download_button(
            "Descargar Word ficha estudiante",
            data=archivo_word,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="descargar_word_estudiante"
        )

    panel_registros_en_formulario("estudiante")


    st.markdown("### Resumen para copiar")
    st.caption("No se incluye en el Word.")
    st.code(resumen_libro or "Sin resumen generado.")



elif menu == "Gestión de Casos":
    encabezado("Gestión de Casos", "Seguimiento institucional por código de caso")

    df_reg = cargar_registros()

    if df_reg.empty:
        st.info("Aún no existen registros guardados. La gestión de casos se activará al guardar entrevistas hp/e.")
    elif "CC" not in df_reg.columns:
        st.warning("El registro existe, pero aún no contiene columna CC. Guarda nuevos registros con la versión actualizada.")
    else:
        df_casos = construir_resumen_gestion_casos(df_reg)

        if df_casos.empty:
            st.info("No existen casos con CC registrado.")
        else:
            st.subheader("Panel central de casos")

            col_a, col_b, col_c, col_d = st.columns(4)
            col_a.metric("Casos", len(df_casos))
            col_b.metric("Cerrados", len(df_casos[df_casos["Estado"] == "Cerrado"]))
            col_c.metric("En proceso/seguimiento", len(df_casos[df_casos["Estado"].isin(["En proceso", "En seguimiento"])]))
            col_d.metric("Con pendientes", len(df_casos[df_casos["Pendientes"] > 0]))

            f_estado = st.selectbox("Filtrar por estado", ["Todos"] + sorted(set(df_casos["Estado"].astype(str).tolist())))
            df_vista = df_casos.copy()
            if f_estado != "Todos":
                df_vista = df_vista[df_vista["Estado"].astype(str) == f_estado]

            st.dataframe(df_vista, use_container_width=True)

            opciones = [
                f"{row['CC']} | {row['Estudiante']} | {row['Curso']} | {row['Estado']} | {row['Avance %']}%"
                for _, row in df_casos.iterrows()
            ]

            st.subheader("Detalle del caso")
            seleccion = st.selectbox("Seleccionar caso", opciones, key="gestion_cc_selector")
            cc_sel = seleccion.split(" | ")[0].strip()
            df_cc = df_reg[df_reg["CC"].astype(str).str.strip() == cc_sel].copy()

            if not df_cc.empty:
                estados = consolidar_checklist_por_cc(df_cc)
                avance, cumplidos, no_aplica, pendientes = calcular_avance_checklist(estados)
                ultima = df_cc.iloc[-1]
                primera = df_cc.iloc[0]

                st.markdown(f"### {cc_sel}")
                st.write(f"**Estudiante:** {primera.get('Estudiante', '')}")
                st.write(f"**Curso:** {primera.get('Curso', '')}")
                st.write(f"**Estado institucional actual:** {ultima.get('Estado_institucional', ultima.get('Estado_caso', ''))}")
                st.write(f"**Última acción:** {ultima.get('Tipo_registro', '')}")
                st.write(f"**Fecha última:** {ultima.get('Fecha_entrevista', '')}")

                st.progress(int(avance) / 100)
                st.caption(f"Avance del caso: {avance}%")

                c1, c2, c3 = st.columns(3)
                c1.metric("Cumplidos", cumplidos)
                c2.metric("No aplica", no_aplica)
                c3.metric("Pendientes", pendientes)

                st.markdown("### Checklist consolidado del caso")
                render_checklist_consolidado(estados)

                st.markdown("### Línea de tiempo")
                columnas = [
                    c for c in [
                        "Fecha_entrevista", "Hora", "Tipo_registro", "Departamento_que_cita",
                        "Motivo", "Estado_institucional", "Protocolos_aplicados", "Usuario_registro"
                    ] if c in df_cc.columns
                ]
                st.dataframe(df_cc[columnas], use_container_width=True)

                st.markdown("### Resumen del estado actual")
                resumen_actual = generar_resumen_estado_actual(df_cc, estados)
                st.text_area(
                    "Resumen generado",
                    value=resumen_actual,
                    height=140,
                    key=f"resumen_estado_caso_{cc_sel}"
                )

                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_cc.to_excel(writer, sheet_name="Linea_tiempo", index=False)
                    pd.DataFrame([{"Paso": k, "Estado": v} for k, v in estados.items()]).to_excel(
                        writer, sheet_name="Checklist_consolidado", index=False
                    )
                buffer.seek(0)

                st.download_button(
                    "Descargar expediente del caso",
                    data=buffer,
                    file_name=f"expediente_{cc_sel}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )



elif menu == "Registros":
    if st.session_state.get("rol") != "Administrador":
        encabezado("Acceso restringido", "Solo rol Administrador")
        st.error("Acceso restringido a Registros. Solo rol Administrador.")
        st.stop()

    encabezado("Registros", "Seguimiento mínimo de entrevistas")

    if REGISTRO_EXCEL.exists():
        try:
            df_reg = pd.read_excel(REGISTRO_EXCEL, sheet_name="Registro_Entrevistas").fillna("")
            st.success(f"Registros encontrados: {len(df_reg)}")

            if not df_reg.empty:
                col1, col2, col3, col4 = st.columns(4)

                cc_reg = ["Todos"] + sorted(set(df_reg["CC"].astype(str).tolist())) if "CC" in df_reg.columns else ["Todos"]
                departamentos_reg = ["Todos"] + sorted(set(df_reg["Departamento_que_cita"].astype(str).tolist()))
                cursos_reg = ["Todos"] + sorted(set(df_reg["Curso"].astype(str).tolist()))
                motivos_reg = ["Todos"] + sorted(set(df_reg["Motivo"].astype(str).tolist()))

                f_cc = col1.selectbox("Filtrar cc", cc_reg)
                f_depto = col2.selectbox("Filtrar departamento", departamentos_reg)
                f_curso = col3.selectbox("Filtrar curso", cursos_reg)
                f_motivo = col4.selectbox("Filtrar motivo", motivos_reg)

                df_filtrado = df_reg.copy()

                if "CC" in df_filtrado.columns and f_cc != "Todos":
                    df_filtrado = df_filtrado[df_filtrado["CC"].astype(str) == f_cc]

                if f_depto != "Todos":
                    df_filtrado = df_filtrado[df_filtrado["Departamento_que_cita"].astype(str) == f_depto]

                if f_curso != "Todos":
                    df_filtrado = df_filtrado[df_filtrado["Curso"].astype(str) == f_curso]

                if f_motivo != "Todos":
                    df_filtrado = df_filtrado[df_filtrado["Motivo"].astype(str) == f_motivo]

                st.dataframe(df_filtrado, use_container_width=True)

                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_filtrado.to_excel(writer, sheet_name="Registros_filtrados", index=False)
                buffer.seek(0)

                st.download_button(
                    "Descargar registros filtrados",
                    data=buffer,
                    file_name=f"registro_entrevistas_filtrado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"No fue posible leer registros: {e}")

    else:
        st.info("Aún no existe datos/registro_entrevistas.xlsx. Se creará al guardar el primer registro.")



elif menu == "Levantamiento de información":
    modulo_levantamiento_informacion()

elif menu == "Administración":

    st.markdown("### NotebookLM normativo")
    st.caption("Uso recomendado: consulta documental de RICE, protocolos y normativa. No pegar datos sensibles de estudiantes.")
    if NOTEBOOKLM_URL:
        st.success("NotebookLM configurado.")
        st.link_button("📚 Abrir cuaderno normativo", NOTEBOOKLM_URL)
    else:
        st.info("Para activar el botón, configurar en PowerShell: setx NOTEBOOKLM_URL \"URL_DEL_CUADERNO_NOTEBOOKLM\"")

    if st.session_state.get("rol") != "Administrador":
        encabezado("Acceso restringido", "Solo rol Administrador")
        st.error("Acceso restringido a rol Administrador.")
        st.stop()

    encabezado("Administración", "Verificación de archivos de Fase 1")

    st.subheader("Archivos requeridos")

    st.write(f"Base principal: `{BASE_EXCEL}`")
    st.write(f"Base definitiva de listas y datos: `{BASE_EXCEL}`")
    st.write(f"Plantilla apoderado: `{PLANTILLA_APODERADO}`")
    st.write(f"Plantilla estudiante: `{PLANTILLA_ESTUDIANTE}`")
    st.write(f"Registro separado: `{REGISTRO_EXCEL}`")

    col1, col2, col3, col4 = st.columns(4)

    col1.success(
        "Base encontrada"
        if BASE_EXCEL.exists()
        else "Falta base_datos_pukaray.xlsx"
    )

    col2.success(
        "Motivos en base definitiva"
        if not df_motivos.empty
        else "Falta hoja lista_motivo_entrevista"
    )

    col3.success(
        "Plantilla encontrada"
        if PLANTILLA_APODERADO.exists()
        else "Falta Ficha_Entrevista_APODERADO.docx"
    )

    st.markdown("### Hojas base_datos_pukaray.xlsx")
    st.write(hojas_excel(str(BASE_EXCEL)))

    st.markdown("### Vista previa estudiantes")
    st.dataframe(df_estudiantes.head(40), use_container_width=True)

    st.markdown("### Vista previa departamentos")
    st.dataframe(df_departamentos.head(40), use_container_width=True)

    st.markdown("### Vista previa lista_faltas")
    st.caption("Selector usa orden: categoría + entorno + tipo")
    st.dataframe(df_lista_faltas.head(40), use_container_width=True)

    st.markdown("### Vista previa tipo registro RICE / lista_faltas")
    st.caption("Orden visible: categoría + entorno + tipo")
    st.dataframe(df_tipo_registro.head(40), use_container_width=True)

    st.markdown("### Vista previa motivos")
    st.dataframe(df_motivos.head(40), use_container_width=True)

    st.markdown("### Vista previa estados de caso")
    st.caption("Origen: base_datos_pukaray.xlsx / hoja lista_estado_casos")
    st.dataframe(leer_hoja(str(BASE_EXCEL), "lista_estado_casos").head(40), use_container_width=True)

    st.markdown("### Vista previa checklist cierre de caso")
    st.caption("Origen: base_datos_pukaray.xlsx / hoja lista_checklist_cierre_caso")
    st.dataframe(leer_hoja(str(BASE_EXCEL), "lista_checklist_cierre_caso").head(40), use_container_width=True)

    st.markdown("### Vista previa estado/detalle de caso")
    st.caption("Origen: base_datos_pukaray.xlsx / hoja lista_estado_detalle_caso")
    st.dataframe(leer_hoja(str(BASE_EXCEL), "lista_estado_detalle_caso").head(40), use_container_width=True)

