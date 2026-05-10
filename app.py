import re, json, unicodedata
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

st.set_page_config(page_title="Sistema Pukaray IA", page_icon="📄", layout="centered")

def cargar_usuarios():
    with open(USERS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def pantalla_login():
    if Path("logo_pukaray.png").exists():
        st.image("logo_pukaray.png", width=160)  
    st.markdown("""
    <div style="background-color:#1a542a;padding:22px;border-radius:14px;margin-bottom:20px;">
        <h1 style="color:white;text-align:center;margin:0;">Sistema Pukaray IA</h1>
        <p style="color:#f5f3eb;text-align:center;margin:6px 0 0 0;">Ingreso funcionarios autorizados</p>
    </div>
    """, unsafe_allow_html=True)
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

def normalizar(texto):
    texto = str(texto or "").strip()
    texto = unicodedata.normalize("NFD", texto)
    return "".join(ch for ch in texto if unicodedata.category(ch) != "Mn").upper().replace(" ", "")

def limpiar_nombre_archivo(texto):
    texto = unicodedata.normalize("NFD", texto or "").encode("ascii", "ignore").decode("utf-8")
    return re.sub(r"[^a-zA-Z0-9]+", "_", texto).strip("_") or "Documento"

def limpiar_para_word(texto):
    return str(texto or "").replace("\\n", "\n").replace("\\t", "\t")

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

def leer_hoja(nombre):
    wb = load_workbook(DB_PATH, data_only=True)
    ws = wb[nombre]
    headers = [str(c.value or "").strip() for c in ws[1]]
    salida = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row): 
            continue
        item = {headers[i]: row[i] if i < len(row) else "" for i in range(len(headers))}
        if str(item.get("Estado", "Activo") or "Activo").strip().lower() == "activo":
            salida.append(item)
    return salida

def contar_intervenciones_previas(nombre_estudiante):
    wb = load_workbook(DB_PATH, data_only=True)
    ws = wb["Seguimiento_Intervenciones"]
    headers = [str(c.value or "").strip() for c in ws[1]]
    if "Nombre Estudiante" not in headers:
        return 0
    idx = headers.index("Nombre Estudiante")
    return sum(1 for row in ws.iter_rows(min_row=2, values_only=True) if row and len(row) > idx and str(row[idx]).strip().lower() == str(nombre_estudiante).strip().lower())

def resumen_personas(registros, nombre_key, cargo_key, depto_key=None, apoyo_key=None):
    return {
        "nombres": ", ".join([str(r.get(nombre_key, "") or "") for r in registros if r.get(nombre_key)]),
        "cargos": ", ".join([str(r.get(cargo_key, "") or "") for r in registros if r.get(cargo_key)]),
        "deptos": ", ".join([str(r.get(depto_key, "") or "") for r in registros if depto_key and r.get(depto_key)]),
        "apoyos": "\n".join([str(r.get(apoyo_key, "") or "") for r in registros if apoyo_key and r.get(apoyo_key)]),
    }

def redactar_textos(antecedentes_mejorados, responsables_apoyo, tipo_apoyo, rice):
    motivo = "Se realiza entrevista de apoderado con el propósito de informar antecedentes asociados al proceso formativo y de convivencia escolar del estudiante.\n\nANTECEDENTES INFORMADOS:\n" + antecedentes_mejorados
    analisis = ["ANÁLISIS INSTITUCIONAL",
        "1. Los antecedentes descritos evidencian una situación que requiere abordaje formativo, resguardo de la convivencia escolar y coordinación con la familia.",
        "2. Se recomienda fortalecer la reflexión del estudiante respecto de sus acciones, promoviendo la reparación del daño y el cumplimiento de las normas institucionales.",
        "3. Clasificación referencial según RICE:"]
    analisis += [f"   • {x}" for x in rice.get("categoria", [])]
    analisis.append("4. Normas posiblemente asociadas:")
    analisis += [f"   • {x}" for x in rice.get("normas", [])]
    acuerdos = ["ACUERDOS Y CONCLUSIONES",
        "1. El apoderado toma conocimiento formal de los antecedentes expuestos durante la entrevista.",
        "2. Se acuerda reforzar desde el hogar normas de respeto, buen trato y resolución adecuada de conflictos.",
        "3. El establecimiento realizará seguimiento institucional del caso."]
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
    return motivo, "\n".join(analisis), "\n".join(acuerdos)

def completar_plantilla(datos, motivo, analisis, acuerdos):
    doc = Document(TEMPLATE_PATH)
    def put(cell, text): cell.text = limpiar_para_word(text)
    put(doc.tables[0].cell(0, 1), datos["nombre_estudiante"])
    put(doc.tables[1].cell(0, 1), datos["curso"])
    put(doc.tables[1].cell(0, 3), datos["fecha"])
    put(doc.tables[1].cell(0, 5), datos["hora"])
    put(doc.tables[2].cell(0, 1), datos["entrevistadores"])
    put(doc.tables[2].cell(0, 3), datos["cargos_entrevistadores"])
    put(doc.tables[3].cell(0, 1), datos["departamentos"])
    put(doc.tables[3].cell(0, 5), datos["numero_entrevista"])
    put(doc.tables[3].cell(1, 1 if datos["asiste_apoderado"] == "Sí" else 2), " X")
    put(doc.tables[3].cell(1, 5 if datos["asiste_estudiante"] == "Sí" else 6), " X")
    put(doc.tables[4].cell(0, 1), motivo)
    put(doc.tables[5].cell(0, 1), f"{analisis}\n\n{acuerdos}")
    bio = BytesIO(); doc.save(bio); bio.seek(0); return bio

def registrar(registro):
    wb = load_workbook(DB_PATH)
    ws = wb["Seguimiento_Intervenciones"]
    ws.append([registro.get(k) for k in [
        "fecha_registro","hora_registro","usuario_sistema","nombre_funcionario","cargo_funcionario","curso","nombre_estudiante","run","nombre_apoderado","relacion_apoderado","telefono_apoderado","correo_apoderado","entrevistadores","cargos_entrevistadores","departamentos","responsables_apoyo","roles_responsables","tipos_apoyo","asiste_apoderado","asiste_estudiante","antecedentes_originales","antecedentes_mejorados","motivo","analisis","acuerdos","categoria_rice","normas_rice","medidas_rice","alertas_rice","gravedad","archivo_generado","numero_entrevista"]])
    wb.save(DB_PATH)

col1, col2 = st.columns([2, 1])
with col1:
    if Path("logo_pukaray.png").exists():
        st.image("logo_pukaray.png", width=120)
    st.markdown("""<div style="background-color:#f5f3eb;border-left:8px solid #1a542a;padding:16px;border-radius:10px;"><h2 style="margin:0;color:#1a542a;">Sistema Pukaray IA</h2><p style="margin:4px 0 0 0;color:#6b1e11;">Entrevistas · RICE · Seguimiento institucional</p></div>""", unsafe_allow_html=True)
    st.caption(f"Usuario conectado: {st.session_state.get('usuario_nombre')} · {st.session_state.get('usuario_cargo')}")
with col2:
    if st.button("Limpiar datos"): limpiar_datos()
    if st.button("Salir"): cerrar_sesion()

if "crear" not in st.session_state.get("usuario_permisos", []):
    st.error("Su usuario no tiene permiso para crear entrevistas.")
    st.stop()

reset_form = st.session_state.get("reset_form", 0)
estudiantes, entrevistadores, responsables = leer_hoja("Estudiantes"), leer_hoja("Entrevistadores"), leer_hoja("Responsables_Apoyo")


def cargar_historial_dataframe():
    wb = load_workbook(DB_PATH, data_only=True)
    ws = wb["Seguimiento_Intervenciones"]

    filas = list(ws.iter_rows(values_only=True))

    if len(filas) <= 1:
        return pd.DataFrame()

    encabezados = [str(x or "") for x in filas[0]]
    datos = filas[1:]

    return pd.DataFrame(datos, columns=encabezados)
    wb = load_workbook(DB_PATH, data_only=True)
    ws = wb["Seguimiento_Intervenciones"]

    filas = list(ws.iter_rows(values_only=True))
    if len(filas) <= 1:
        return []

    encabezados = [str(x or "") for x in filas[0]]
    registros = []

    for fila in filas[1:]:
        registro = {}
        for i, encabezado in enumerate(encabezados):
            registro[encabezado] = fila[i] if i < len(fila) else ""
        registros.append(registro)

    return registros
cursos = sorted({str(e.get("Curso", "")).strip() for e in estudiantes if str(e.get("Curso", "")).strip()})
curso_sel = st.selectbox("Curso", ["Seleccione curso"] + cursos, index=0, key=f"curso_sel_{reset_form}")
estudiantes_filtrados = [e for e in estudiantes if curso_sel != "Seleccione curso" and normalizar(e.get("Curso", "")) == normalizar(curso_sel)]
nombres_estudiantes = [str(e.get("Nombre Estudiante", "")).strip() for e in estudiantes_filtrados]
estudiante_sel = st.selectbox("Estudiante", ["Seleccione estudiante"] + nombres_estudiantes, index=0, key=f"estudiante_sel_{reset_form}_{normalizar(curso_sel)}")
st.subheader("Historial del estudiante")
st.divider()

st.header("Estadísticas institucionales")

df = cargar_historial_dataframe()

if df.empty:

    st.info("Aún no existen registros para estadísticas.")

else:

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Total intervenciones", len(df))

    with col2:
        if "Nombre Estudiante" in df.columns:
            st.metric(
                "Estudiantes intervenidos",
                df["Nombre Estudiante"].nunique()
            )

    with col3:
        if "Curso" in df.columns:
            st.metric(
                "Cursos con registros",
                df["Curso"].nunique()
            )

    st.subheader("Intervenciones por curso")

    if "Curso" in df.columns:
        st.bar_chart(df["Curso"].value_counts())

    st.subheader("Intervenciones por gravedad")

    if "Gravedad" in df.columns:
        st.bar_chart(df["Gravedad"].value_counts())

    st.subheader("Estudiantes con más intervenciones")

    if "Nombre Estudiante" in df.columns:
        top_estudiantes = (
            df["Nombre Estudiante"]
            .value_counts()
            .head(10)
        )

        st.bar_chart(top_estudiantes)

st.subheader("Funcionarios que registran entrevistas")

if "Nombre Funcionario" in df.columns:
    st.bar_chart(df["Nombre Funcionario"].value_counts())
else:
    st.info("No existe la columna Nombre Funcionario en la base de datos.")
historial = cargar_historial_dataframe()

st.subheader("Historial del estudiante")

df_historial = cargar_historial_dataframe()

if estudiante_sel == "Seleccione estudiante":
    st.info("Seleccione un estudiante para ver su historial.")

elif df_historial.empty:
    st.info("No existen intervenciones registradas.")

else:
    df_estudiante = df_historial[
        df_historial["Nombre Estudiante"].fillna("").str.strip().str.lower()
        == str(estudiante_sel).strip().lower()
    ]

    if df_estudiante.empty:
        st.info("No existen intervenciones previas registradas para este estudiante.")
    else:
        columnas_mostrar = [
            "Fecha Registro",
            "Hora Registro",
            "Curso",
            "Nombre Estudiante",
            "Gravedad",
            "Categoría RICE",
            "Medidas RICE",
            "Archivo Generado"
        ]

        columnas_mostrar = [
            c for c in columnas_mostrar
            if c in df_estudiante.columns
        ]

        st.dataframe(
            df_estudiante[columnas_mostrar],
            use_container_width=True
        )

if estudiante_sel == "Seleccione estudiante":
    st.info("Seleccione un estudiante para ver su historial.")
elif not historial_estudiante:
    st.info("No existen intervenciones previas registradas para este estudiante.")
else:
    for h in historial_estudiante:
        with st.expander(f"{h.get('Fecha Registro', '')} - {h.get('Gravedad', '')}"):
            st.write(f"**Curso:** {h.get('Curso', '')}")
            st.write(f"**Funcionario:** {h.get('Nombre Funcionario', '')}")
            st.write(f"**Antecedentes:** {h.get('Antecedentes Mejorados', '')}")
            st.write(f"**Categoría RICE:** {h.get('Categoría RICE', '')}")
            st.write(f"**Normas RICE:** {h.get('Normas RICE', '')}")
            st.write(f"**Medidas:** {h.get('Medidas RICE', '')}")
            st.write(f"**Archivo:** {h.get('Archivo Generado', '')}")
estudiante = next((e for e in estudiantes_filtrados if str(e.get("Nombre Estudiante", "")).strip() == estudiante_sel), {})

fecha = st.date_input("Fecha entrevista", value=date.today(), format="DD/MM/YYYY", key=f"fecha_{reset_form}")
hora = st.text_input("Hora entrevista", placeholder="Ej: 17:00 hrs", key=f"hora_{reset_form}")
numero_entrevista = st.text_input("Número entrevista", placeholder="Ej: 001-2026", key=f"numero_{reset_form}")
apoderado_nombre = st.text_input("Nombre apoderado entrevistado", key=f"apoderado_nombre_{reset_form}")
apoderado_relacion = st.text_input("Relación con estudiante", key=f"apoderado_relacion_{reset_form}")
apoderado_telefono = st.text_input("Teléfono apoderado", key=f"apoderado_telefono_{reset_form}")
apoderado_correo = st.text_input("Correo apoderado", key=f"apoderado_correo_{reset_form}")

nombres_entrevistadores = [e.get("Nombre Entrevistador", "") for e in entrevistadores if e.get("Nombre Entrevistador")]
entrevistadores_sel = st.multiselect("Entrevistadores participantes", nombres_entrevistadores, default=nombres_entrevistadores[:1], key=f"entrevistadores_{reset_form}")
resumen_ent = resumen_personas([e for e in entrevistadores if e.get("Nombre Entrevistador") in entrevistadores_sel], "Nombre Entrevistador", "Cargo", "Departamento")

nombres_responsables = [r.get("Nombre Responsable", "") for r in responsables if r.get("Nombre Responsable")]
responsables_sel = st.multiselect("Responsables a cargo de ejecutar apoyos", nombres_responsables, default=nombres_responsables[:1], key=f"responsables_{reset_form}")
resumen_resp = resumen_personas([r for r in responsables if r.get("Nombre Responsable") in responsables_sel], "Nombre Responsable", "Cargo/Rol", "Área", "Tipo de Apoyo")

tipo_apoyo_extra = st.text_area("Ajuste o detalle del apoyo a ejecutar", value=resumen_resp["apoyos"], height=110, key=f"tipo_apoyo_{reset_form}")
asiste_apoderado = st.selectbox("Asiste apoderado", ["Sí", "No"], key=f"asiste_apoderado_{reset_form}")
asiste_estudiante = st.selectbox("Asiste estudiante", ["No", "Sí"], key=f"asiste_estudiante_{reset_form}")
antecedentes = st.text_area("Antecedentes breves del caso", height=160, placeholder="Ej: le pegó a otro compañero y lo insultó", key=f"antecedentes_{reset_form}")
mejorar_texto = st.checkbox("Mejorar automáticamente la redacción institucional", value=True, key=f"mejorar_texto_{reset_form}")
incluir_rice = st.checkbox("Analizar antecedentes según RICE", value=True, key=f"incluir_rice_{reset_form}")

if st.button("Generar documento y registrar seguimiento", type="primary"):
    if curso_sel == "Seleccione curso" or estudiante_sel == "Seleccione estudiante":
        st.error("Debe seleccionar curso y estudiante.")
    else:
        nombre_estudiante, run, curso = estudiante.get("Nombre Estudiante", ""), estudiante.get("RUN", "") or "", curso_sel
        antecedentes_mejorados = mejorar_antecedentes(antecedentes) if mejorar_texto else antecedentes
        rice = analizar_rice(f"{antecedentes} {antecedentes_mejorados}", contar_intervenciones_previas(nombre_estudiante)) if incluir_rice else {"categoria":["No solicitado"],"normas":["No solicitado"],"medidas":[],"alertas":[],"gravedad":"BAJA"}
        motivo, analisis, acuerdos = redactar_textos(antecedentes_mejorados, resumen_resp["nombres"], tipo_apoyo_extra, rice)
        nombre_archivo = f"{limpiar_nombre_archivo(nombre_estudiante)}_{limpiar_nombre_archivo(curso)}_{fecha.strftime('%d-%m-%Y')}.docx"
        archivo = completar_plantilla({"nombre_estudiante":nombre_estudiante,"curso":curso,"fecha":fecha.strftime("%d.%m.%Y"),"hora":hora,"entrevistadores":resumen_ent["nombres"],"cargos_entrevistadores":resumen_ent["cargos"],"departamentos":resumen_ent["deptos"],"numero_entrevista":numero_entrevista,"asiste_apoderado":asiste_apoderado,"asiste_estudiante":asiste_estudiante}, motivo, analisis, acuerdos)
        ahora = datetime.now()
        registrar({"fecha_registro":ahora.strftime("%d.%m.%Y"),"hora_registro":ahora.strftime("%H:%M:%S"),"usuario_sistema":st.session_state.get("usuario_id"),"nombre_funcionario":st.session_state.get("usuario_nombre"),"cargo_funcionario":st.session_state.get("usuario_cargo"),"curso":curso,"nombre_estudiante":nombre_estudiante,"run":run,"nombre_apoderado":apoderado_nombre,"relacion_apoderado":apoderado_relacion,"telefono_apoderado":apoderado_telefono,"correo_apoderado":apoderado_correo,"entrevistadores":resumen_ent["nombres"],"cargos_entrevistadores":resumen_ent["cargos"],"departamentos":resumen_ent["deptos"],"responsables_apoyo":resumen_resp["nombres"],"roles_responsables":resumen_resp["cargos"],"tipos_apoyo":tipo_apoyo_extra,"asiste_apoderado":asiste_apoderado,"asiste_estudiante":asiste_estudiante,"antecedentes_originales":antecedentes,"antecedentes_mejorados":antecedentes_mejorados,"motivo":motivo,"analisis":analisis,"acuerdos":acuerdos,"categoria_rice":"\n".join(rice.get("categoria", [])),"normas_rice":"\n".join(rice.get("normas", [])),"medidas_rice":"\n".join(rice.get("medidas", [])),"alertas_rice":"\n".join(rice.get("alertas", [])),"gravedad":rice.get("gravedad","BAJA"),"archivo_generado":nombre_archivo,"numero_entrevista":numero_entrevista})
        st.success("Documento generado y seguimiento registrado correctamente.")
        st.download_button("Descargar Word listo para imprimir", archivo, file_name=nombre_archivo, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
