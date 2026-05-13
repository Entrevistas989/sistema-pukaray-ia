import re
import json
import unicodedata
import shutil
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from openpyxl import load_workbook

from motor_rice import analizar_rice
from redactor_institucional import mejorar_antecedentes


TEMPLATE_PARTICIPANTE = "plantilla_ficha_entrevista_apoderado.docx"
TEMPLATE_ESTUDIANTE = "Ficha_Entrevista_ESTUDIANTE.docx"
TEMPLATE_FUNCIONARIO = "Ficha_Entrevista_FUNCIONARIO.docx"

DB_PATH = "base_datos_pukaray.xlsx"
USERS_PATH = "usuarios.json"
LOGO_PATH = "logo_pukaray.png"


st.set_page_config(
    page_title="Sistema Pukaray IA",
    page_icon="📄",
    layout="centered"
)


def normalizar(texto):
    texto = str(texto or "").strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    return texto.upper().replace(" ", "")


def normalizar_clave(texto):
    texto = str(texto or "").strip().lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    texto = re.sub(r"[^a-z0-9]+", "_", texto).strip("_")
    return texto


def limpiar_nombre_archivo(texto):
    texto = unicodedata.normalize("NFD", texto or "")
    texto = texto.encode("ascii", "ignore").decode("utf-8")
    texto = re.sub(r"[^a-zA-Z0-9]+", "_", texto).strip("_")
    return texto or "Documento"


def limpiar_para_word(texto):
    return str(texto or "").replace("\\n", "\n").replace("\\t", "\t")


def obtener_iniciales_usuario():
    usuario_actual = st.session_state.get("usuario_nombre", "")
    iniciales = "".join([palabra[0].upper() for palabra in usuario_actual.split() if palabra])
    return iniciales or "US"


def cargar_usuarios():
    with open(USERS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def pantalla_login():
    if Path(LOGO_PATH).exists():
        st.image(LOGO_PATH, width=160)

    st.markdown(
        """
        <div style="background-color:#1a542a;padding:22px;border-radius:14px;margin-bottom:20px;">
            <h1 style="color:white;text-align:center;margin:0;">Sistema Pukaray IA</h1>
            <p style="color:#f5f3eb;text-align:center;margin:6px 0 0 0;">Ingreso funcionarios autorizados</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    usuarios = cargar_usuarios()
    usuario = st.text_input("Usuario")
    clave = st.text_input("Contraseña", type="password")

    if st.button("Ingresar", type="primary"):
        if usuario in usuarios and usuarios[usuario]["password"] == clave:
            st.session_state["autenticado"] = True
            st.session_state["usuario_id"] = usuario
            st.session_state["usuario_nombre"] = usuarios[usuario]["nombre"]
            st.session_state["usuario_cargo"] = usuarios[usuario]["cargo"]
            st.session_state["usuario_permisos"] = usuarios[usuario].get("permisos", [])
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrecta.")


if not st.session_state.get("autenticado"):
    pantalla_login()
    st.stop()


def limpiar_datos():
    sesion = {
        "autenticado": st.session_state.get("autenticado"),
        "usuario_id": st.session_state.get("usuario_id"),
        "usuario_nombre": st.session_state.get("usuario_nombre"),
        "usuario_cargo": st.session_state.get("usuario_cargo"),
        "usuario_permisos": st.session_state.get("usuario_permisos", []),
        "reset_form": st.session_state.get("reset_form", 0) + 1,
    }
    st.session_state.clear()
    st.session_state.update(sesion)
    st.rerun()


def cerrar_sesion():
    st.session_state.clear()
    st.rerun()


def leer_hoja(nombre_hoja):
    wb = load_workbook(DB_PATH, data_only=True)
    ws = wb[nombre_hoja]
    headers = [str(c.value or "").strip() for c in ws[1]]
    registros = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        item = {headers[i]: row[i] if i < len(row) else "" for i in range(len(headers))}
        estado = str(item.get("Estado", "Activo") or "Activo").strip().lower()
        if estado == "activo":
            registros.append(item)
    return registros


def cargar_historial_dataframe():
    wb = load_workbook(DB_PATH, data_only=True)
    ws = wb["Seguimiento_Intervenciones"]
    filas = list(ws.iter_rows(values_only=True))

    if len(filas) <= 1:
        return pd.DataFrame()

    encabezados_originales = [str(x or "") for x in filas[0]]
    encabezados_tecnicos = [normalizar_clave(x) for x in encabezados_originales]
    datos = filas[1:]

    return pd.DataFrame(datos, columns=encabezados_tecnicos)


def contar_intervenciones_previas(nombre_estudiante):
    df = cargar_historial_dataframe()
    if df.empty or "nombre_estudiante" not in df.columns:
        return 0
    return int((df["nombre_estudiante"].fillna("").astype(str).str.strip().str.lower() == str(nombre_estudiante).strip().lower()).sum())


def resumen_personas(registros, nombre_key, cargo_key, depto_key=None, apoyo_key=None):
    nombres, cargos, deptos, apoyos = [], [], [], []
    for r in registros:
        if r.get(nombre_key):
            nombres.append(str(r.get(nombre_key, "") or ""))
        if r.get(cargo_key):
            cargos.append(str(r.get(cargo_key, "") or ""))
        if depto_key and r.get(depto_key):
            deptos.append(str(r.get(depto_key, "") or ""))
        if apoyo_key and r.get(apoyo_key):
            apoyos.append(str(r.get(apoyo_key, "") or ""))

    return {
        "nombres": ", ".join(nombres),
        "cargos": ", ".join(cargos),
        "deptos": ", ".join(deptos),
        "apoyos": "\n".join(apoyos),
    }


def seleccionar_plantilla(tipo_registro):
    if tipo_registro == "Entrevista participante":
        return TEMPLATE_PARTICIPANTE
    if tipo_registro == "Atención estudiante":
        return TEMPLATE_ESTUDIANTE
    return TEMPLATE_FUNCIONARIO


def detectar_reincidencia_institucional(cantidad_intervenciones_previas):
    if cantidad_intervenciones_previas >= 5:
        return {
            "nivel": "ALTA",
            "texto": (
                "Se advierte reiteración significativa de intervenciones previas asociadas al estudiante. "
                "Se recomienda seguimiento prioritario, revisión del caso por el equipo de convivencia escolar "
                "y evaluación de progresión de medidas formativas, disciplinarias o de apoyo según corresponda."
            )
        }

    if cantidad_intervenciones_previas >= 2:
        return {
            "nivel": "MODERADA",
            "texto": (
                "Se observa registro previo de intervenciones asociadas al estudiante. "
                "Se sugiere mantener seguimiento institucional, reforzar acuerdos y monitorear la evolución del caso."
            )
        }

    if cantidad_intervenciones_previas == 1:
        return {
            "nivel": "BAJA",
            "texto": (
                "Existe un antecedente previo registrado, por lo que se recomienda seguimiento preventivo "
                "y observación de la evolución conductual o formativa."
            )
        }

    return {
        "nivel": "SIN REGISTROS PREVIOS",
        "texto": ""
    }


def detectar_acuerdos_inteligentes(texto, tipo_registro, gravedad):
    texto_base = normalizar_clave(texto).replace("_", " ")
    acuerdos_extra = []

    def contiene(*palabras):
        return any(p in texto_base for p in palabras)

    if contiene("golpe", "pega", "pego", "agresion", "agrede", "empuja", "zancadilla", "patada"):
        acuerdos_extra += [
            "Se acuerda promover una acción reparatoria proporcional al daño causado, resguardando la dignidad de las personas involucradas.",
            "Se realizará seguimiento conductual para verificar que no exista reiteración de acciones físicas o de riesgo.",
        ]

    if contiene("insulto", "garabato", "ofensa", "amenaza", "amenazo", "humilla", "hiriente"):
        acuerdos_extra += [
            "Se acuerda reforzar el uso de lenguaje respetuoso y estrategias de comunicación adecuadas.",
            "Se orienta a evitar respuestas impulsivas, amenazas o verbalizaciones que puedan afectar la convivencia escolar.",
        ]

    if contiene("burla", "burlas", "apodo", "sobrenombre", "ridiculiza", "molesta"):
        acuerdos_extra += [
            "Se acuerda trabajar el buen trato, la empatía y el reconocimiento del impacto que generan las burlas en otros integrantes de la comunidad.",
            "Se realizará seguimiento preventivo para evitar nuevas situaciones de menoscabo o trato despectivo.",
        ]

    if contiene("desregulacion", "crisis", "ansiedad", "llanto", "emocional", "impulsivo", "impulsividad"):
        acuerdos_extra += [
            "Se acuerda fortalecer estrategias de autorregulación emocional y solicitud oportuna de apoyo adulto.",
            "Se evaluará la pertinencia de acompañamiento socioemocional o coordinación con profesionales de apoyo.",
        ]

    if contiene("interrumpe", "interrupciones", "disruptivo", "disruptiva", "clase", "ruidos", "molesta"):
        acuerdos_extra += [
            "Se acuerda reforzar normas de participación en clases y respeto por el proceso de aprendizaje del grupo curso.",
            "Se realizará monitoreo de la conducta en aula y retroalimentación formativa según evolución.",
        ]

    if contiene("redes", "whatsapp", "instagram", "mensaje", "foto", "video", "grupo"):
        acuerdos_extra += [
            "Se acuerda orientar sobre el uso responsable de redes sociales y medios digitales, resguardando la privacidad y dignidad de terceros.",
            "Se reforzará que cualquier conflicto digital sea informado oportunamente a adultos responsables del establecimiento.",
        ]

    if contiene("inasistencia", "atraso", "atrasos", "licencia", "reintegro", "ausencia"):
        acuerdos_extra += [
            "Se acuerda mantener seguimiento de asistencia, puntualidad y proceso de reintegro escolar.",
            "Se coordinarán apoyos pedagógicos o formativos si la situación afecta la continuidad del proceso educativo.",
        ]

    if contiene("funcionario", "docente", "profesor", "asistente", "laboral", "equipo") or tipo_registro == "Atención funcionario":
        acuerdos_extra += [
            "Se acuerda resguardar canales formales de comunicación institucional y mantener registro de los acuerdos adoptados.",
            "Se reforzará la coordinación interna, el buen trato y el cumplimiento de procedimientos establecidos por el colegio.",
        ]

    if gravedad in ["ALTA", "GRAVE", "CRÍTICA", "CRITICA"]:
        acuerdos_extra += [
            "Debido a la gravedad referencial, se sugiere seguimiento prioritario y revisión del caso por el equipo correspondiente.",
            "Se deberá dejar registro de nuevas acciones, avances o incumplimientos asociados al caso.",
        ]

    salida = []
    for acuerdo in acuerdos_extra:
        if acuerdo not in salida:
            salida.append(acuerdo)

    return salida


def redactar_textos(tipo_registro, antecedentes_mejorados, responsables_apoyo, tipo_apoyo, rice, intervenciones_previas=0):

    gravedad = str(rice.get("gravedad", "BAJA")).upper()
    categorias = rice.get("categoria", []) or []
    normas = rice.get("normas", []) or []
    medidas = rice.get("medidas", []) or []
    alertas = rice.get("alertas", []) or []

    categorias_txt = "; ".join([str(x) for x in categorias]) if categorias else "Sin clasificación automática específica."
    normas_txt = "; ".join([str(x) for x in normas]) if normas else "Sin norma específica asociada automáticamente."
    requiere_seguimiento_prioritario = gravedad in ["ALTA", "GRAVE", "CRÍTICA", "CRITICA"]

    reincidencia = detectar_reincidencia_institucional(intervenciones_previas)
    texto_reincidencia = reincidencia.get("texto", "")
    nivel_reincidencia = reincidencia.get("nivel", "SIN REGISTROS PREVIOS")

    acuerdos_inteligentes = detectar_acuerdos_inteligentes(
        antecedentes_mejorados,
        tipo_registro,
        gravedad
    )

    if tipo_registro == "Entrevista participante":

        motivo = (
            "Se realiza entrevista con participante convocado para informar, contextualizar y abordar "
            "antecedentes asociados al proceso formativo y de convivencia escolar del estudiante.\n\n"
            "ANTECEDENTES EXPUESTOS:\n"
            f"{antecedentes_mejorados}"
        )

        analisis = [
            "ANÁLISIS INSTITUCIONAL",
            "1. La entrevista se orienta a entregar información clara al participante, recoger antecedentes complementarios y favorecer la coordinación entre familia y establecimiento.",
            "2. El análisis considera el impacto de los hechos en la convivencia escolar, el proceso formativo del estudiante y la necesidad de acompañamiento institucional.",
            f"3. Según revisión referencial del RICE, la situación se vincula con: {categorias_txt}",
            f"4. Normativa o criterios institucionales asociados: {normas_txt}",
        ]

        if requiere_seguimiento_prioritario:
            analisis.append(
                "5. Dada la gravedad referencial del caso, se requiere seguimiento prioritario, registro sistemático y coordinación oportuna con los responsables de apoyo."
            )
        else:
            analisis.append(
                "5. Se sugiere seguimiento formativo y comunicación permanente para prevenir reiteración de la conducta o agravamiento de la situación."
            )

        acuerdos = [
            "ACUERDOS Y COMPROMISOS",
            "1. El participante toma conocimiento formal de los antecedentes expuestos y de la orientación institucional entregada.",
            "2. Se acuerda reforzar desde el hogar normas de respeto, buen trato, responsabilidad y resolución pacífica de conflictos.",
            "3. El establecimiento mantendrá seguimiento del caso y comunicará avances o nuevas situaciones relevantes.",
        ]

    elif tipo_registro == "Atención estudiante":

        motivo = (
            "Se realiza atención individual de estudiante para promover reflexión formativa, identificar factores asociados "
            "a la situación y orientar conductas de reparación, autocontrol y buen trato.\n\n"
            "RELATO Y ANTECEDENTES ABORDADOS:\n"
            f"{antecedentes_mejorados}"
        )

        analisis = [
            "ANÁLISIS FORMATIVO",
            "1. La intervención se centra en que el estudiante reconozca los hechos abordados, sus consecuencias y el impacto que pueden generar en otros integrantes de la comunidad escolar.",
            "2. Se refuerza la importancia del autocontrol, la responsabilidad personal, la empatía y el cumplimiento de las normas de convivencia.",
            f"3. Según revisión referencial del RICE, la situación se vincula con: {categorias_txt}",
            f"4. Normativa o criterios institucionales asociados: {normas_txt}",
        ]

        if requiere_seguimiento_prioritario:
            analisis.append(
                "5. Por la gravedad referencial detectada, se requiere seguimiento cercano, eventual coordinación con apoderado y evaluación de medidas formativas o disciplinarias según corresponda."
            )
        else:
            analisis.append(
                "5. Se recomienda acompañamiento formativo y monitoreo preventivo para fortalecer cambios conductuales sostenidos."
            )

        acuerdos = [
            "ACUERDOS Y REFLEXIONES",
            "1. El estudiante toma conocimiento de los antecedentes abordados y participa en una reflexión guiada sobre su conducta.",
            "2. Se compromete a evitar la reiteración de la situación y a utilizar canales adecuados para resolver conflictos o solicitar ayuda.",
            "3. Se acuerda seguimiento formativo por parte del establecimiento, considerando la evolución de la conducta y el cumplimiento de los compromisos.",
        ]

    else:

        motivo = (
            "Se realiza atención individual de funcionario para registrar antecedentes, entregar orientación institucional "
            "y coordinar acciones que favorezcan el resguardo, la comunicación adecuada y el buen funcionamiento interno.\n\n"
            "ANTECEDENTES REGISTRADOS:\n"
            f"{antecedentes_mejorados}"
        )

        analisis = [
            "ANÁLISIS INSTITUCIONAL",
            "1. La atención se aborda desde criterios de resguardo institucional, convivencia interna, responsabilidad profesional y coordinación entre funcionarios.",
            "2. Se releva la necesidad de mantener comunicación formal, trato respetuoso y cumplimiento de protocolos o acuerdos internos.",
            f"3. Referencia institucional asociada al análisis: {categorias_txt}",
            f"4. Criterios normativos o institucionales relacionados: {normas_txt}",
        ]

        if requiere_seguimiento_prioritario:
            analisis.append(
                "5. Por la gravedad referencial del antecedente, se sugiere seguimiento directivo o de convivencia, resguardando la confidencialidad y la trazabilidad del caso."
            )
        else:
            analisis.append(
                "5. Se recomienda mantener seguimiento institucional preventivo y registro de los acuerdos adoptados."
            )

        acuerdos = [
            "ACUERDOS Y ORIENTACIONES",
            "1. El funcionario toma conocimiento formal de los antecedentes tratados y de las orientaciones institucionales entregadas.",
            "2. Se acuerda mantener canales formales de comunicación y coordinación según corresponda.",
            "3. Se reforzarán criterios de buen trato, resguardo institucional y cumplimiento de funciones o protocolos relacionados.",
        ]

    if texto_reincidencia:
        acuerdos.append("ANTECEDENTES DE REINCIDENCIA O SEGUIMIENTO:")
        acuerdos.append(f"   • {texto_reincidencia}")
        acuerdos.append(f"   • Nivel referencial de reincidencia institucional: {nivel_reincidencia}.")

    if acuerdos_inteligentes:
        acuerdos.append("ACUERDOS ESPECÍFICOS SEGÚN ANTECEDENTES DEL CASO:")
        acuerdos += [f"   • {x}" for x in acuerdos_inteligentes]

    if medidas:
        acuerdos.append("MEDIDAS Y ACCIONES SUGERIDAS SEGÚN RICE:")
        acuerdos += [f"   • {x}" for x in medidas]

    if responsables_apoyo:
        acuerdos.append(f"RESPONSABLES DE APOYO Y SEGUIMIENTO: {responsables_apoyo}.")

    if tipo_apoyo:
        acuerdos.append("APOYOS COMPROMETIDOS:")
        acuerdos += [
            f"   • {linea.strip()}"
            for linea in str(tipo_apoyo).splitlines()
            if linea.strip()
        ]

    acuerdos.append(f"CLASIFICACIÓN REFERENCIAL DE GRAVEDAD INSTITUCIONAL: {gravedad}.")

    if alertas:
        acuerdos.append("ALERTAS INSTITUCIONALES PARA SEGUIMIENTO:")
        acuerdos += [f"   • {x}" for x in alertas]

    firma = (
        f"\n\nDocumento generado por:\n"
        f"{st.session_state.get('usuario_nombre', '')}\n"
        f"{st.session_state.get('usuario_cargo', '')}"
    )

    return motivo, "\n".join(analisis), "\n".join(acuerdos) + firma


def reemplazar_texto_en_doc(doc):
    reemplazos = {
        "APODERADO": "PARTICIPANTE",
        "Apoderado": "Participante",
        "apoderado": "participante",
    }
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    for run in parrafo.runs:
                        for viejo, nuevo in reemplazos.items():
                            if viejo in run.text:
                                run.text = run.text.replace(viejo, nuevo)


def completar_plantilla(datos, motivo, analisis, acuerdos, tipo_registro):
    plantilla = seleccionar_plantilla(tipo_registro)
    doc = Document(plantilla)

    if tipo_registro == "Entrevista participante":
        reemplazar_texto_en_doc(doc)

    def put(cell, text):
        cell.text = limpiar_para_word(text)

    # Folio institucional: parte superior derecha, antes del título y fuera del motivo.
    parrafo_folio = doc.paragraphs[0].insert_paragraph_before()
    parrafo_folio.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run_folio = parrafo_folio.add_run(f"Folio N° {datos['numero_entrevista']}")
    run_folio.bold = True

    if tipo_registro == "Entrevista participante":
        put(doc.tables[0].cell(0, 1), datos["nombre_estudiante"])
        put(doc.tables[1].cell(0, 1), datos["curso"])
        put(doc.tables[1].cell(0, 3), datos["fecha"])
        put(doc.tables[1].cell(0, 5), datos["hora"])
        put(doc.tables[2].cell(0, 1), datos["entrevistadores"])
        put(doc.tables[2].cell(0, 3), datos["cargos_entrevistadores"])
        put(doc.tables[3].cell(0, 1), datos["departamentos"])
        put(doc.tables[3].cell(0, 5), datos["numero_entrevista"])
        put(doc.tables[3].cell(1, 1 if datos["asiste_participante"] == "Sí" else 2), " X")
        put(doc.tables[3].cell(1, 5 if datos["asiste_estudiante"] == "Sí" else 6), " X")
        put(doc.tables[4].cell(0, 1), motivo)
        put(doc.tables[5].cell(0, 1), f"{analisis}\n\n{acuerdos}")

    elif tipo_registro == "Atención estudiante":
        put(doc.tables[0].cell(0, 1), datos["nombre_estudiante"])
        put(doc.tables[1].cell(0, 1), datos["curso"])
        put(doc.tables[1].cell(0, 3), datos["fecha"])
        put(doc.tables[1].cell(0, 5), datos["hora"])
        put(doc.tables[2].cell(0, 1), datos["entrevistadores"])
        put(doc.tables[2].cell(0, 3), datos["cargos_entrevistadores"])
        put(doc.tables[3].cell(0, 1), motivo)
        put(doc.tables[4].cell(0, 1), f"{analisis}\n\n{acuerdos}")

    else:
        put(doc.tables[0].cell(0, 1), datos["participante_entrevista"])
        put(doc.tables[1].cell(0, 1), datos["vinculo_persona"])
        put(doc.tables[1].cell(0, 3), datos["fecha"])
        put(doc.tables[1].cell(0, 5), datos["hora"])
        put(doc.tables[2].cell(0, 1), datos["entrevistadores"])
        put(doc.tables[2].cell(0, 3), datos["cargos_entrevistadores"])
        put(doc.tables[3].cell(0, 1), motivo)
        put(doc.tables[4].cell(0, 1), f"{analisis}\n\n{acuerdos}")

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


def generar_folio():
    df = cargar_historial_dataframe()

    if df.empty or "numero_entrevista" not in df.columns:
        return "000001"

    numeros = []

    for valor in df["numero_entrevista"].dropna():
        valor = str(valor).strip()
        solo_num = "".join(filter(str.isdigit, valor))

        if solo_num:
            numeros.append(int(solo_num))

    if not numeros:
        return "000001"

    siguiente = max(numeros) + 1
    return str(siguiente).zfill(6)


def registrar(registro):
    wb = load_workbook(DB_PATH)
    ws = wb["Seguimiento_Intervenciones"]
    encabezados = [normalizar_clave(celda.value) for celda in ws[1]]
    nueva_fila = [registro.get(encabezado, "") for encabezado in encabezados]
    ws.append(nueva_fila)
    wb.save(DB_PATH)


def crear_respaldo():
    carpeta = Path("RESPALDOS")
    carpeta.mkdir(exist_ok=True)
    fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    destino = carpeta / f"respaldo_base_datos_pukaray_{fecha}.xlsx"
    shutil.copy(DB_PATH, destino)
    return destino


def generar_informe_word(df_filtrado, titulo):
    doc = Document()
    doc.add_heading("Informe Institucional de Intervenciones", level=1)
    doc.add_paragraph("Colegio Pukaray")
    doc.add_paragraph(f"Fecha de generación: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    doc.add_paragraph(f"Generado por: {st.session_state.get('usuario_nombre', '')} - {st.session_state.get('usuario_cargo', '')}")
    doc.add_heading(titulo, level=2)
    doc.add_paragraph(f"Total de registros: {len(df_filtrado)}")

    for _, fila in df_filtrado.iterrows():
        doc.add_paragraph(
            (
                f"Fecha: {fila.get('fecha_registro', '')}\n"
                f"Tipo registro: {fila.get('tipo_registro', '')}\n"
                f"Estudiante: {fila.get('nombre_estudiante', '')}\n"
                f"Curso: {fila.get('curso', '')}\n"
                f"Participante: {fila.get('participante_entrevista', '')}\n"
                f"Vínculo: {fila.get('vinculo_persona', '')}\n"
                f"Gravedad: {fila.get('gravedad', '')}\n"
                f"Categoría RICE: {fila.get('categoria_rice', '')}\n"
                f"Medidas: {fila.get('medidas_rice', '')}\n"
            )
        )

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# =========================================================
# INTERFAZ PRINCIPAL
# =========================================================

col1, col2 = st.columns([2, 1])

with col1:
    if Path(LOGO_PATH).exists():
        st.image(LOGO_PATH, width=120)
    st.markdown(
        """
        <div style="background-color:#f5f3eb;border-left:8px solid #1a542a;padding:16px;border-radius:10px;">
            <h2 style="margin:0;color:#1a542a;">Sistema Pukaray IA</h2>
            <p style="margin:4px 0 0 0;color:#6b1e11;">Entrevistas · RICE · Seguimiento institucional</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(f"Usuario conectado: {st.session_state.get('usuario_nombre')} · {st.session_state.get('usuario_cargo')}")

with col2:
    if st.button("Limpiar datos"):
        limpiar_datos()
    if st.button("Salir"):
        cerrar_sesion()


if "crear" not in st.session_state.get("usuario_permisos", []):
    st.error("Su usuario no tiene permiso para crear entrevistas.")
    st.stop()


reset_form = st.session_state.get("reset_form", 0)
estudiantes = leer_hoja("Estudiantes")
entrevistadores = leer_hoja("Entrevistadores")
responsables = leer_hoja("Responsables_Apoyo")
df_historial = cargar_historial_dataframe()


# =========================================================
# FORMULARIO
# =========================================================

st.divider()
st.subheader("Datos de la entrevista")

tipo_registro = st.selectbox(
    "Tipo de registro",
    [
        "Seleccione tipo entrevista",
        "Entrevista participante",
        "Atención estudiante",
        "Atención funcionario"
    ],
    index=0,
    key=f"tipo_registro_{reset_form}",
)

st.subheader("Selección según tipo de registro")

curso_sel = "No aplica"
estudiante_sel = "No aplica"
estudiante = {}
funcionario_sel = "Seleccione funcionario"
funcionario_data = {}

if tipo_registro == "Seleccione tipo entrevista":
    st.warning("Debe seleccionar un tipo de entrevista para continuar.")
    st.stop()

if tipo_registro in ["Entrevista participante", "Atención estudiante"]:
    cursos = sorted({str(e.get("Curso", "")).strip() for e in estudiantes if str(e.get("Curso", "")).strip()})
    curso_sel = st.selectbox("Curso", ["Seleccione curso"] + cursos, index=0, key=f"curso_sel_{reset_form}")

    estudiantes_filtrados = [
        e for e in estudiantes
        if curso_sel != "Seleccione curso" and normalizar(e.get("Curso", "")) == normalizar(curso_sel)
    ]

    nombres_estudiantes = [str(e.get("Nombre Estudiante", "")).strip() for e in estudiantes_filtrados]

    estudiante_sel = st.selectbox(
        "Estudiante",
        ["Seleccione estudiante"] + nombres_estudiantes,
        index=0,
        key=f"estudiante_sel_{reset_form}_{normalizar(curso_sel)}",
    )

    estudiante = next(
        (e for e in estudiantes_filtrados if str(e.get("Nombre Estudiante", "")).strip() == estudiante_sel),
        {},
    )

else:
    nombres_funcionarios = [e.get("Nombre Entrevistador", "") for e in entrevistadores if e.get("Nombre Entrevistador")]
    funcionario_sel = st.selectbox(
        "Funcionario a entrevistar",
        ["Seleccione funcionario"] + nombres_funcionarios,
        index=0,
        key=f"funcionario_sel_{reset_form}",
    )
    funcionario_data = next(
        (e for e in entrevistadores if str(e.get("Nombre Entrevistador", "")).strip() == funcionario_sel),
        {},
    )


# =========================================================
# HISTORIAL DEL ESTUDIANTE
# =========================================================

if tipo_registro in ["Entrevista participante", "Atención estudiante"]:
    st.subheader("Historial del estudiante")

    if estudiante_sel == "Seleccione estudiante":
        st.info("Seleccione un estudiante para ver su historial.")
    elif df_historial.empty:
        st.info("No existen intervenciones registradas.")
    elif "nombre_estudiante" not in df_historial.columns:
        st.warning("La hoja Seguimiento_Intervenciones no contiene la columna nombre_estudiante.")
    else:
        df_estudiante = df_historial[
            df_historial["nombre_estudiante"].fillna("").astype(str).str.strip().str.lower()
            == str(estudiante_sel).strip().lower()
        ]

        df_estudiante = df_estudiante.drop_duplicates()
        if df_estudiante.empty:
            st.info("No existen intervenciones previas registradas para este estudiante.")
        else:
            columnas_mostrar = [
                "fecha_registro",
                "hora_registro",
                "tipo_registro",
                "curso",
                "nombre_estudiante",
                "gravedad",
                "categoria_rice",
                "medidas_rice",
                "archivo_generado",
            ]
            columnas_mostrar = [c for c in columnas_mostrar if c in df_estudiante.columns]
            st.dataframe(df_estudiante[columnas_mostrar], use_container_width=True)


# =========================================================
# DATOS DE REGISTRO
# =========================================================

numero_entrevista = generar_folio()

st.info(f"Folio automático asignado: {numero_entrevista}")

fecha = st.date_input(
    "Fecha entrevista",
    value=date.today(),
    format="DD/MM/YYYY",
    key=f"fecha_{reset_form}",
)

hora = st.text_input(
    "Hora entrevista",
    placeholder="Ej: 17:00 hrs",
    key=f"hora_{reset_form}",
)

# =========================================================
# DATOS SEGÚN TIPO DE REGISTRO
# =========================================================

if tipo_registro == "Entrevista participante":
    participante_entrevista = st.text_input(
        "Participante de la entrevista",
        key=f"participante_entrevista_{reset_form}",
    )

    vinculo_persona = st.text_input(
        "Vínculo con el estudiante",
        key=f"vinculo_persona_{reset_form}",
    )

    asiste_participante = st.selectbox(
        "Asiste participante de la entrevista",
        ["Sí", "No"],
        key=f"asiste_participante_{reset_form}",
    )

    asiste_estudiante = st.selectbox(
        "Asiste estudiante",
        ["No", "Sí"],
        key=f"asiste_estudiante_{reset_form}",
    )

elif tipo_registro == "Atención estudiante":
    participante_entrevista = estudiante_sel
    vinculo_persona = "Estudiante"
    asiste_participante = "Sí"
    asiste_estudiante = "Sí"
    st.info("Registro de atención individual de estudiante.")

else:
    participante_entrevista = funcionario_sel
    vinculo_persona = funcionario_data.get("Cargo", "")
    asiste_participante = "Sí"
    asiste_estudiante = "No"

    st.text_input(
        "Funcionario seleccionado",
        value=participante_entrevista,
        disabled=True,
    )

    st.text_input(
        "Cargo o función",
        value=vinculo_persona,
        disabled=True,
    )

    st.info("Registro de atención individual de funcionario.")

# =========================================================
# PARTICIPANTES INSTITUCIONALES
# =========================================================

st.subheader("Participantes institucionales")

nombres_entrevistadores = [e.get("Nombre Entrevistador", "") for e in entrevistadores if e.get("Nombre Entrevistador")]
entrevistadores_sel = st.multiselect(
    "Entrevistadores participantes",
    nombres_entrevistadores,
    default=[],
    key=f"entrevistadores_{reset_form}",
)

resumen_ent = resumen_personas(
    [e for e in entrevistadores if e.get("Nombre Entrevistador") in entrevistadores_sel],
    "Nombre Entrevistador",
    "Cargo",
    "Departamento",
)

nombres_responsables = [r.get("Nombre Responsable", "") for r in responsables if r.get("Nombre Responsable")]
responsables_sel = st.multiselect(
    "Responsables a cargo de ejecutar apoyos",
    nombres_responsables,
    default=[],
    key=f"responsables_{reset_form}",
)

resumen_resp = resumen_personas(
    [r for r in responsables if r.get("Nombre Responsable") in responsables_sel],
    "Nombre Responsable",
    "Cargo/Rol",
    "Área",
    "Tipo de Apoyo",
)

tipo_apoyo_extra = st.text_area(
    "Ajuste o detalle del apoyo a ejecutar",
    value=resumen_resp["apoyos"],
    height=110,
    key=f"tipo_apoyo_{reset_form}",
)


# =========================================================
# ANTECEDENTES
# =========================================================

st.subheader("Antecedentes")

antecedentes = st.text_area(
    "Antecedentes breves del caso",
    height=160,
    placeholder="Ej: le pegó a otro compañero y lo insultó",
    key=f"antecedentes_{reset_form}",
)

mejorar_texto = st.checkbox("Mejorar automáticamente la redacción institucional", value=True, key=f"mejorar_texto_{reset_form}")
incluir_rice = st.checkbox("Analizar antecedentes según RICE", value=True, key=f"incluir_rice_{reset_form}")


# =========================================================
# GENERAR DOCUMENTO
# =========================================================

if st.button("Generar documento y registrar seguimiento", type="primary"):
    if tipo_registro in ["Entrevista participante", "Atención estudiante"]:
        if curso_sel == "Seleccione curso" or estudiante_sel == "Seleccione estudiante":
            st.error("Debe seleccionar curso y estudiante.")
            st.stop()

    if tipo_registro == "Atención funcionario" and funcionario_sel == "Seleccione funcionario":
        st.error("Debe seleccionar funcionario.")
        st.stop()

    nombre_estudiante = estudiante.get("Nombre Estudiante", "") if estudiante else ""
    run = estudiante.get("RUN", "") if estudiante else ""
    curso = curso_sel

    antecedentes_mejorados = mejorar_antecedentes(antecedentes) if mejorar_texto else antecedentes

    rice = (
        analizar_rice(f"{antecedentes} {antecedentes_mejorados}", contar_intervenciones_previas(nombre_estudiante))
        if incluir_rice
        else {"categoria": ["No solicitado"], "normas": ["No solicitado"], "medidas": [], "alertas": [], "gravedad": "BAJA"}
    )

    intervenciones_previas = contar_intervenciones_previas(nombre_estudiante)

    motivo, analisis, acuerdos = redactar_textos(
        tipo_registro,
        antecedentes_mejorados,
        resumen_resp["nombres"],
        tipo_apoyo_extra,
        rice,
        intervenciones_previas,
    )

    iniciales = obtener_iniciales_usuario()
    nombre_base = participante_entrevista if tipo_registro == "Atención funcionario" else nombre_estudiante

    nombre_archivo = (
        f"{limpiar_nombre_archivo(nombre_base)}_"
        f"{limpiar_nombre_archivo(tipo_registro)}_"
        f"{fecha.strftime('%d-%m-%Y')}_"
        f"{iniciales}.docx"
    )

    archivo = completar_plantilla(
        {
            "nombre_estudiante": nombre_estudiante,
            "curso": curso,
            "fecha": fecha.strftime("%d.%m.%Y"),
            "hora": hora,
            "participante_entrevista": participante_entrevista,
            "vinculo_persona": vinculo_persona,
            "entrevistadores": resumen_ent["nombres"],
            "cargos_entrevistadores": resumen_ent["cargos"],
            "departamentos": resumen_ent["deptos"],
            "numero_entrevista": numero_entrevista,
            "asiste_participante": asiste_participante,
            "asiste_estudiante": asiste_estudiante,
        },
        motivo,
        analisis,
        acuerdos,
        tipo_registro,
    )

    ahora = datetime.now()

    registrar(
        {
            "fecha_registro": ahora.strftime("%d.%m.%Y"),
            "hora_registro": ahora.strftime("%H:%M:%S"),
            "usuario_sistema": st.session_state.get("usuario_id"),
            "nombre_funcionario": st.session_state.get("usuario_nombre"),
            "cargo_funcionario": st.session_state.get("usuario_cargo"),
            "tipo_registro": tipo_registro,
            "curso": curso,
            "nombre_estudiante": nombre_estudiante,
            "run": run,
            "participante_entrevista": participante_entrevista,
            "vinculo_persona": vinculo_persona,
            "entrevistadores": resumen_ent["nombres"],
            "cargos_entrevistadores": resumen_ent["cargos"],
            "departamentos": resumen_ent["deptos"],
            "responsables_apoyo": resumen_resp["nombres"],
            "roles_responsables": resumen_resp["cargos"],
            "tipos_apoyo": tipo_apoyo_extra,
            "asiste_participante": asiste_participante,
            "asiste_estudiante": asiste_estudiante,
            "antecedentes_originales": antecedentes,
            "antecedentes_mejorados": antecedentes_mejorados,
            "motivo": motivo,
            "analisis": analisis,
            "acuerdos": acuerdos,
            "categoria_rice": "\n".join(rice.get("categoria", [])),
            "normas_rice": "\n".join(rice.get("normas", [])),
            "medidas_rice": "\n".join(rice.get("medidas", [])),
            "alertas_rice": "\n".join(rice.get("alertas", [])),
            "gravedad": rice.get("gravedad", "BAJA"),
            "archivo_generado": nombre_archivo,
            "numero_entrevista": numero_entrevista,
            "intervenciones_previas": intervenciones_previas,
            "nivel_reincidencia": detectar_reincidencia_institucional(intervenciones_previas).get("nivel", ""),
        }
    )

    st.success("Documento generado y seguimiento registrado correctamente.")
    st.download_button(
        "Descargar Word listo para imprimir",
        archivo,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# =========================================================
# INFORMES INSTITUCIONALES
# =========================================================

st.divider()
st.header("Informes institucionales")

df_historial = cargar_historial_dataframe()

tipo_informe = st.selectbox(
    "Tipo de informe",
    ["Por estudiante", "Por curso", "General colegio"],
    key="tipo_informe",
)

df_filtrado = pd.DataFrame()
titulo_informe = ""

if df_historial.empty:
    st.info("Aún no existen registros para generar informes.")
else:
    if tipo_informe == "Por estudiante":
        opciones = sorted(df_historial["nombre_estudiante"].dropna().astype(str).unique().tolist()) if "nombre_estudiante" in df_historial.columns else []
        informe_estudiante = st.selectbox("Seleccione estudiante", ["Seleccione estudiante"] + opciones, key="informe_estudiante")
        if informe_estudiante != "Seleccione estudiante":
            df_filtrado = df_historial[df_historial["nombre_estudiante"].astype(str) == informe_estudiante]
            titulo_informe = f"Informe por estudiante: {informe_estudiante}"

    elif tipo_informe == "Por curso":
        opciones = sorted(df_historial["curso"].dropna().astype(str).unique().tolist()) if "curso" in df_historial.columns else []
        informe_curso = st.selectbox("Seleccione curso", ["Seleccione curso"] + opciones, key="informe_curso")
        if informe_curso != "Seleccione curso":
            df_filtrado = df_historial[df_historial["curso"].astype(str) == informe_curso]
            titulo_informe = f"Informe por curso: {informe_curso}"

    else:
        df_filtrado = df_historial
        titulo_informe = "Informe general del colegio"

    if st.button("Generar informe Word"):
        if df_filtrado.empty:
            st.warning("No hay registros seleccionados para generar informe.")
        else:
            informe = generar_informe_word(df_filtrado, titulo_informe)
            nombre_informe = f"{limpiar_nombre_archivo(titulo_informe)}_{datetime.now().strftime('%d-%m-%Y')}.docx"
            st.download_button(
                "Descargar informe Word",
                informe,
                file_name=nombre_informe,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )


# =========================================================
# RESPALDOS
# =========================================================

st.divider()
st.header("Respaldos institucionales")

if st.button("Crear respaldo de base de datos"):
    respaldo = crear_respaldo()
    st.success(f"Respaldo creado correctamente: {respaldo.name}")

if Path(DB_PATH).exists():
    with open(DB_PATH, "rb") as archivo:
        st.download_button(
            "Descargar base de datos actual",
            archivo,
            file_name="base_datos_pukaray.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
