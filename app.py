
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
from openpyxl import load_workbook

from motor_rice import analizar_rice
from redactor_institucional import mejorar_antecedentes


TEMPLATE_PATH = "plantilla_ficha_entrevista_apoderado.docx"
DB_PATH = "base_datos_pukaray.xlsx"
USERS_PATH = "usuarios.json"
LOGO_PATH = "logo_pukaray.png"


st.set_page_config(
    page_title="Sistema Pukaray IA",
    page_icon="📄",
    layout="centered"
)


# =========================================================
# UTILIDADES GENERALES
# =========================================================

def normalizar(texto):
    texto = str(texto or "").strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    return texto.upper().replace(" ", "")


def normalizar_clave(texto):
    """
    Convierte encabezados de Excel a claves técnicas.
    Ejemplo:
    'Fecha Registro' -> 'fecha_registro'
    'Categoría RICE' -> 'categoria_rice'
    """
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
    iniciales = "".join([
        palabra[0].upper()
        for palabra in usuario_actual.split()
        if palabra
    ])
    return iniciales or "US"


# =========================================================
# LOGIN
# =========================================================

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
        unsafe_allow_html=True
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


def tiene_permiso(nombre_permiso):
    return nombre_permiso in st.session_state.get("usuario_permisos", [])


def limpiar_datos():
    sesion = {
        "autenticado": st.session_state.get("autenticado"),
        "usuario_id": st.session_state.get("usuario_id"),
        "usuario_nombre": st.session_state.get("usuario_nombre"),
        "usuario_cargo": st.session_state.get("usuario_cargo"),
        "usuario_permisos": st.session_state.get("usuario_permisos", []),
        "reset_form": st.session_state.get("reset_form", 0) + 1
    }
    st.session_state.clear()
    st.session_state.update(sesion)
    st.rerun()


def cerrar_sesion():
    st.session_state.clear()
    st.rerun()


# =========================================================
# LECTURA BASE DE DATOS
# =========================================================

def leer_hoja(nombre_hoja):
    wb = load_workbook(DB_PATH, data_only=True)
    ws = wb[nombre_hoja]

    headers = [str(c.value or "").strip() for c in ws[1]]
    registros = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue

        item = {
            headers[i]: row[i] if i < len(row) else ""
            for i in range(len(headers))
        }

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

    return int(
        (
            df["nombre_estudiante"].fillna("").astype(str).str.strip().str.lower()
            == str(nombre_estudiante).strip().lower()
        ).sum()
    )


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


# =========================================================
# REDACCIÓN Y WORD
# =========================================================

def redactar_textos(antecedentes_mejorados, responsables_apoyo, tipo_apoyo, rice):
    motivo = (
        "Se realiza entrevista con participante convocado, con el propósito de informar antecedentes "
        "asociados al proceso formativo y de convivencia escolar del estudiante.\n\n"
        "ANTECEDENTES INFORMADOS:\n"
        f"{antecedentes_mejorados}"
    )

    analisis = [
        "ANÁLISIS INSTITUCIONAL",
        "1. Los antecedentes descritos evidencian una situación que requiere abordaje formativo, resguardo de la convivencia escolar y coordinación con la familia.",
        "2. Se recomienda fortalecer la reflexión del estudiante respecto de sus acciones, promoviendo la reparación del daño y el cumplimiento de las normas institucionales.",
        "3. Clasificación referencial según RICE:"
    ]

    analisis += [f"   • {x}" for x in rice.get("categoria", [])]
    analisis.append("4. Normas posiblemente asociadas:")
    analisis += [f"   • {x}" for x in rice.get("normas", [])]

    acuerdos = [
        "ACUERDOS Y CONCLUSIONES",
        "1. El participante de la entrevista toma conocimiento formal de los antecedentes expuestos.",
        "2. Se acuerda reforzar desde el hogar y/o desde el rol correspondiente las normas de respeto, buen trato y resolución adecuada de conflictos.",
        "3. El establecimiento realizará seguimiento institucional del caso."
    ]

    if responsables_apoyo:
        acuerdos.append(f"4. Responsables de ejecución y seguimiento de apoyos: {responsables_apoyo}.")

    if tipo_apoyo:
        acuerdos.append("5. Apoyos comprometidos:")
        acuerdos += [f"   • {linea.strip()}" for linea in str(tipo_apoyo).splitlines() if linea.strip()]

    if rice.get("medidas"):
        acuerdos.append("6. Medidas formativas sugeridas según análisis RICE:")
        acuerdos += [f"   • {x}" for x in rice.get("medidas", [])]

    acuerdos.append(f"7. Nivel referencial de gravedad institucional: {rice.get('gravedad', 'BAJA')}.")

    if rice.get("alertas"):
        acuerdos.append("8. Alertas para revisión del equipo:")
        acuerdos += [f"   • {x}" for x in rice.get("alertas", [])]

    firma = (
        "\n\nDocumento generado por:\n"
        f"{st.session_state.get('usuario_nombre', '')}\n"
        f"{st.session_state.get('usuario_cargo', '')}"
    )

    return motivo, "\n".join(analisis), "\n".join(acuerdos) + firma


def reemplazar_texto_en_doc(doc):
    """
    Ajusta textos visibles heredados de la plantilla Word.
    Mantiene el formato general, pero reemplaza 'Apoderado' por 'Participante'.
    """
    reemplazos = {
        "APODERADO": "PARTICIPANTE",
        "Apoderado": "Participante",
        "apoderado": "participante"
    }

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    for run in parrafo.runs:
                        for viejo, nuevo in reemplazos.items():
                            if viejo in run.text:
                                run.text = run.text.replace(viejo, nuevo)


def completar_plantilla(datos, motivo, analisis, acuerdos):
    doc = Document(TEMPLATE_PATH)
    reemplazar_texto_en_doc(doc)

    def put(cell, text):
        cell.text = limpiar_para_word(text)

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

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# =========================================================
# REGISTRO Y RESPALDOS
# =========================================================

def registrar(registro):
    wb = load_workbook(DB_PATH)
    ws = wb["Seguimiento_Intervenciones"]

    encabezados = [
        normalizar_clave(celda.value)
        for celda in ws[1]
    ]

    nueva_fila = [
        registro.get(encabezado, "")
        for encabezado in encabezados
    ]

    ws.append(nueva_fila)
    wb.save(DB_PATH)


def crear_respaldo():
    carpeta = Path("RESPALDOS")
    carpeta.mkdir(exist_ok=True)

    fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    destino = carpeta / f"respaldo_base_datos_pukaray_{fecha}.xlsx"

    shutil.copy(DB_PATH, destino)

    return destino


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
        unsafe_allow_html=True
    )

    st.caption(
        f"Usuario conectado: {st.session_state.get('usuario_nombre')} · "
        f"{st.session_state.get('usuario_cargo')}"
    )

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


# =========================================================
# SELECCIÓN DE ESTUDIANTE
# =========================================================

st.subheader("Datos del estudiante")

cursos = sorted({
    str(e.get("Curso", "")).strip()
    for e in estudiantes
    if str(e.get("Curso", "")).strip()
})

curso_sel = st.selectbox(
    "Curso",
    ["Seleccione curso"] + cursos,
    index=0,
    key=f"curso_sel_{reset_form}"
)

estudiantes_filtrados = [
    e for e in estudiantes
    if curso_sel != "Seleccione curso"
    and normalizar(e.get("Curso", "")) == normalizar(curso_sel)
]

nombres_estudiantes = [
    str(e.get("Nombre Estudiante", "")).strip()
    for e in estudiantes_filtrados
]

estudiante_sel = st.selectbox(
    "Estudiante",
    ["Seleccione estudiante"] + nombres_estudiantes,
    index=0,
    key=f"estudiante_sel_{reset_form}_{normalizar(curso_sel)}"
)

estudiante = next(
    (
        e for e in estudiantes_filtrados
        if str(e.get("Nombre Estudiante", "")).strip() == estudiante_sel
    ),
    {}
)


# =========================================================
# HISTORIAL DEL ESTUDIANTE
# =========================================================

st.subheader("Historial del estudiante")

df_historial = cargar_historial_dataframe()

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

    if df_estudiante.empty:
        st.info("No existen intervenciones previas registradas para este estudiante.")
    else:
        columnas_mostrar = [
            "fecha_registro",
            "hora_registro",
            "curso",
            "nombre_estudiante",
            "gravedad",
            "categoria_rice",
            "medidas_rice",
            "archivo_generado"
        ]

        columnas_mostrar = [
            c for c in columnas_mostrar
            if c in df_estudiante.columns
        ]

        st.dataframe(
            df_estudiante[columnas_mostrar],
            use_container_width=True
        )


# =========================================================
# FORMULARIO ENTREVISTA
# =========================================================

st.divider()
st.subheader("Datos de la entrevista")
tipo_registro = st.selectbox(
    "Tipo de registro",
    [
        "Entrevista participante",
        "Atención estudiante",
        "Atención funcionario"
    ],
    key=f"tipo_registro_{reset_form}"
)
fecha = st.date_input(
    "Fecha entrevista",
    value=date.today(),
    format="DD/MM/YYYY",
    key=f"fecha_{reset_form}"
)

hora = st.text_input(
    "Hora entrevista",
    placeholder="Ej: 17:00 hrs",
    key=f"hora_{reset_form}"
)

numero_entrevista = st.text_input(
    "Número entrevista",
    placeholder="Ej: 001-2026",
    key=f"numero_{reset_form}"
)

participante_entrevista = st.text_input(
    "Participante de la entrevista",
    key=f"participante_entrevista_{reset_form}"
)

vinculo_persona = st.text_input(
    "Vínculo con el estudiante",
    key=f"vinculo_persona_{reset_form}"
)

asiste_participante = st.selectbox(
    "Asiste participante de la entrevista",
    ["Sí", "No"],
    key=f"asiste_participante_{reset_form}"
)

asiste_estudiante = st.selectbox(
    "Asiste estudiante",
    ["No", "Sí"],
    key=f"asiste_estudiante_{reset_form}"
)


st.subheader("Participantes institucionales")

nombres_entrevistadores = [
    e.get("Nombre Entrevistador", "")
    for e in entrevistadores
    if e.get("Nombre Entrevistador")
]

entrevistadores_sel = st.multiselect(
    "Entrevistadores participantes",
    nombres_entrevistadores,
    default=[],
    key=f"entrevistadores_{reset_form}"
)

resumen_ent = resumen_personas(
    [e for e in entrevistadores if e.get("Nombre Entrevistador") in entrevistadores_sel],
    "Nombre Entrevistador",
    "Cargo",
    "Departamento"
)


nombres_responsables = [
    r.get("Nombre Responsable", "")
    for r in responsables
    if r.get("Nombre Responsable")
]

responsables_sel = st.multiselect(
    "Responsables a cargo de ejecutar apoyos",
    nombres_responsables,
    default=[],
    key=f"responsables_{reset_form}"
)

resumen_resp = resumen_personas(
    [r for r in responsables if r.get("Nombre Responsable") in responsables_sel],
    "Nombre Responsable",
    "Cargo/Rol",
    "Área",
    "Tipo de Apoyo"
)

tipo_apoyo_extra = st.text_area(
    "Ajuste o detalle del apoyo a ejecutar",
    value=resumen_resp["apoyos"],
    height=110,
    key=f"tipo_apoyo_{reset_form}"
)


st.subheader("Antecedentes")

antecedentes = st.text_area(
    "Antecedentes breves del caso",
    height=160,
    placeholder="Ej: le pegó a otro compañero y lo insultó",
    key=f"antecedentes_{reset_form}"
)

mejorar_texto = st.checkbox(
    "Mejorar automáticamente la redacción institucional",
    value=True,
    key=f"mejorar_texto_{reset_form}"
)

incluir_rice = st.checkbox(
    "Analizar antecedentes según RICE",
    value=True,
    key=f"incluir_rice_{reset_form}"
)


if st.button("Generar documento y registrar seguimiento", type="primary"):

    if curso_sel == "Seleccione curso" or estudiante_sel == "Seleccione estudiante":
        st.error("Debe seleccionar curso y estudiante.")

    else:
        nombre_estudiante = estudiante.get("Nombre Estudiante", "")
        run = estudiante.get("RUN", "") or ""
        curso = curso_sel

        antecedentes_mejorados = (
            mejorar_antecedentes(antecedentes)
            if mejorar_texto
            else antecedentes
        )

        rice = (
            analizar_rice(
                f"{antecedentes} {antecedentes_mejorados}",
                contar_intervenciones_previas(nombre_estudiante)
            )
            if incluir_rice
            else {
                "categoria": ["No solicitado"],
                "normas": ["No solicitado"],
                "medidas": [],
                "alertas": [],
                "gravedad": "BAJA"
            }
        )

        motivo, analisis, acuerdos = redactar_textos(
            antecedentes_mejorados,
            resumen_resp["nombres"],
            tipo_apoyo_extra,
            rice
        )

        iniciales = obtener_iniciales_usuario()

        nombre_archivo = (
            f"{limpiar_nombre_archivo(nombre_estudiante)}_"
            f"{limpiar_nombre_archivo(curso)}_"
            f"{fecha.strftime('%d-%m-%Y')}_"
            f"{iniciales}.docx"
        )

        archivo = completar_plantilla(
            {
                "nombre_estudiante": nombre_estudiante,
                "curso": curso,
                "fecha": fecha.strftime("%d.%m.%Y"),
                "hora": hora,
                "entrevistadores": resumen_ent["nombres"],
                "cargos_entrevistadores": resumen_ent["cargos"],
                "departamentos": resumen_ent["deptos"],
                "numero_entrevista": numero_entrevista,
                "asiste_participante": asiste_participante,
                "asiste_estudiante": asiste_estudiante
            },
            motivo,
            analisis,
            acuerdos
        )

        ahora = datetime.now()

        registrar({
            "fecha_registro": ahora.strftime("%d.%m.%Y"),
            "hora_registro": ahora.strftime("%H:%M:%S"),
            "usuario_sistema": st.session_state.get("usuario_id"),
            "nombre_funcionario": st.session_state.get("usuario_nombre"),
            "cargo_funcionario": st.session_state.get("usuario_cargo"),
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
            "numero_entrevista": numero_entrevista
        })

        st.success("Documento generado y seguimiento registrado correctamente.")

        st.download_button(
            "Descargar Word listo para imprimir",
            archivo,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )


# =========================================================
# ESTADÍSTICAS
# =========================================================

st.divider()
st.header("Estadísticas institucionales")

df = cargar_historial_dataframe()

if df.empty:
    st.info("Aún no existen registros para estadísticas.")

else:
    col_a, col_b, col_c = st.columns(3)

    with col_a:
        st.metric("Total intervenciones", len(df))

    with col_b:
        if "nombre_estudiante" in df.columns:
            st.metric(
                "Estudiantes intervenidos",
                df["nombre_estudiante"].nunique()
            )

    with col_c:
        if "curso" in df.columns:
            st.metric(
                "Cursos con registros",
                df["curso"].nunique()
            )

    st.subheader("Intervenciones por curso")
    if "curso" in df.columns:
        st.bar_chart(df["curso"].value_counts())

    st.subheader("Intervenciones por gravedad")
    if "gravedad" in df.columns:
        st.bar_chart(df["gravedad"].value_counts())

    st.subheader("Estudiantes con más intervenciones")
    if "nombre_estudiante" in df.columns:
        st.bar_chart(df["nombre_estudiante"].value_counts().head(10))

    st.subheader("Funcionarios que registran entrevistas")
    if "nombre_funcionario" in df.columns:
        st.bar_chart(df["nombre_funcionario"].value_counts())


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
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
