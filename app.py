
import re
import unicodedata
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import streamlit as st
from docx import Document
from openpyxl import load_workbook

from motor_rice import analizar_rice

TEMPLATE_PATH = "plantilla_ficha_entrevista_apoderado.docx"
DB_PATH = "base_datos_pukaray.xlsx"

st.set_page_config(page_title="Sistema Pukaray IA", page_icon="📄", layout="centered", initial_sidebar_state="collapsed")


def normalizar(texto):
    texto = str(texto or "").strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    return texto.upper().replace(" ", "")


def limpiar_nombre_archivo(texto):
    texto = unicodedata.normalize("NFD", texto or "")
    texto = texto.encode("ascii", "ignore").decode("utf-8")
    texto = re.sub(r"[^a-zA-Z0-9]+", "_", texto).strip("_")
    return texto or "Documento"


def limpiar_formulario():
    for clave in [
        "hora_entrevista", "numero_entrevista", "apoderado_nombre", "apoderado_relacion",
        "apoderado_telefono", "apoderado_correo", "tipo_apoyo_extra", "antecedentes",
        "asiste_apoderado", "asiste_estudiante", "incluir_rice", "estudiante_sel"
    ]:
        st.session_state.pop(clave, None)
    st.rerun()


def salir_programa():
    st.session_state["salir"] = True
    st.rerun()


def leer_hoja(nombre_hoja):
    if not Path(DB_PATH).exists():
        st.error(f"No se encontró {DB_PATH}.")
        return []
    wb = load_workbook(DB_PATH, data_only=True)
    if nombre_hoja not in wb.sheetnames:
        st.error(f"No existe la hoja {nombre_hoja}.")
        return []
    ws = wb[nombre_hoja]
    headers = [str(c.value or "").strip() for c in ws[1]]
    registros = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        item = {headers[i]: (row[i] if i < len(row) else "") for i in range(len(headers))}
        estado = str(item.get("Estado", "Activo") or "Activo").strip().lower()
        if estado == "activo":
            registros.append(item)
    return registros


def contar_intervenciones_previas(nombre_estudiante):
    if not Path(DB_PATH).exists():
        return 0
    wb = load_workbook(DB_PATH, data_only=True)
    if "Seguimiento_Intervenciones" not in wb.sheetnames:
        return 0
    ws = wb["Seguimiento_Intervenciones"]
    headers = [str(c.value or "").strip() for c in ws[1]]
    if "Nombre Estudiante" not in headers:
        return 0
    idx = headers.index("Nombre Estudiante")
    total = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and len(row) > idx and str(row[idx]).strip().lower() == str(nombre_estudiante).strip().lower():
            total += 1
    return total


def mostrar_dato(label, valor):
    st.markdown(f"**{label}:** {valor if valor else '—'}")


def resumen_personas(registros, nombre_key, cargo_key, depto_key=None, apoyo_key=None):
    nombres, cargos, deptos, apoyos = [], [], [], []
    for r in registros:
        nombres.append(str(r.get(nombre_key, "") or ""))
        cargos.append(str(r.get(cargo_key, "") or ""))
        if depto_key:
            deptos.append(str(r.get(depto_key, "") or ""))
        if apoyo_key:
            apoyos.append(str(r.get(apoyo_key, "") or ""))
    return {
        "nombres": ", ".join([x for x in nombres if x]),
        "cargos": ", ".join([x for x in cargos if x]),
        "deptos": ", ".join([x for x in deptos if x]),
        "apoyos": "\n".join([x for x in apoyos if x]),
    }


def redactar_textos(antecedentes, responsables_apoyo, tipo_apoyo, rice):
    motivo = "Se realiza entrevista de apoderado con el propósito de informar antecedentes asociados al proceso formativo y de convivencia escolar del estudiante."
    if antecedentes:
        motivo += f"\n\nAntecedentes informados: {antecedentes}"

    analisis = (
        "Desde el ámbito institucional, los antecedentes descritos sugieren la necesidad de fortalecer "
        "el acompañamiento formativo del estudiante, promoviendo la reflexión, la autorregulación y la "
        "coordinación permanente con la familia. "
        f"Según análisis referencial del RICE, la situación se vincula con: {'; '.join(rice.get('categoria', []))}. "
        f"Normas posiblemente asociadas: {'; '.join(rice.get('normas', []))}."
    )

    acuerdos = (
        "El apoderado toma conocimiento de los antecedentes expuestos. Se acuerda mantener comunicación "
        "permanente con el establecimiento, reforzar normas y compromisos desde el hogar y realizar seguimiento del caso."
    )
    if responsables_apoyo:
        acuerdos += f"\n\nLa ejecución y seguimiento de los apoyos quedará a cargo de: {responsables_apoyo}."
    if tipo_apoyo:
        acuerdos += f"\nTipo de apoyo comprometido:\n{tipo_apoyo}"
    if rice.get("medidas"):
        acuerdos += f"\n\nMedidas formativas sugeridas según análisis RICE: {'; '.join(rice.get('medidas', []))}."
    acuerdos += f"\n\nClasificación referencial de gravedad institucional: {rice.get('gravedad', 'BAJA')}."
    if rice.get("alertas"):
        acuerdos += f"\n\nAlertas para revisión del equipo: {'; '.join(rice.get('alertas', []))}"
    return motivo, analisis, acuerdos


def agregar_texto(celda, texto):
    if celda.paragraphs:
        celda.paragraphs[0].add_run(str(texto or ""))
    else:
        celda.text = str(texto or "")


def completar_plantilla(datos, motivo, acuerdos):
    doc = Document(TEMPLATE_PATH)
    agregar_texto(doc.tables[0].cell(0, 1), datos["nombre_estudiante"])
    agregar_texto(doc.tables[1].cell(0, 1), datos["curso"])
    agregar_texto(doc.tables[1].cell(0, 3), datos["fecha"])
    agregar_texto(doc.tables[1].cell(0, 5), datos["hora"])
    agregar_texto(doc.tables[2].cell(0, 1), datos["entrevistadores"])
    agregar_texto(doc.tables[2].cell(0, 3), datos["cargos_entrevistadores"])
    agregar_texto(doc.tables[3].cell(0, 1), datos["departamentos"])
    agregar_texto(doc.tables[3].cell(0, 5), datos["numero_entrevista"])
    agregar_texto(doc.tables[3].cell(1, 1 if datos["asiste_apoderado"] == "Sí" else 2), " X")
    agregar_texto(doc.tables[3].cell(1, 5 if datos["asiste_estudiante"] == "Sí" else 6), " X")
    doc.tables[4].cell(0, 1).text = motivo
    doc.tables[5].cell(0, 1).text = acuerdos
    salida = BytesIO()
    doc.save(salida)
    salida.seek(0)
    return salida


def registrar(registro):
    wb = load_workbook(DB_PATH)
    ws = wb["Seguimiento_Intervenciones"]
    ws.append([
        registro.get("fecha_registro"), registro.get("hora_registro"), registro.get("curso"),
        registro.get("nombre_estudiante"), registro.get("run"), registro.get("nombre_apoderado"),
        registro.get("relacion_apoderado"), registro.get("telefono_apoderado"), registro.get("correo_apoderado"),
        registro.get("entrevistadores"), registro.get("cargos_entrevistadores"), registro.get("departamentos"),
        registro.get("responsables_apoyo"), registro.get("roles_responsables"), registro.get("tipos_apoyo"),
        registro.get("asiste_apoderado"), registro.get("asiste_estudiante"), registro.get("antecedentes"),
        registro.get("motivo"), registro.get("analisis"), registro.get("acuerdos"),
        registro.get("categoria_rice"), registro.get("normas_rice"), registro.get("medidas_rice"), registro.get("alertas_rice"),
        registro.get("gravedad"), registro.get("archivo_generado"), registro.get("numero_entrevista"),
    ])
    wb.save(DB_PATH)


if st.session_state.get("salir"):
    st.title("Sistema Pukaray IA")
    st.success("Sesión cerrada en este equipo.")
    if st.button("Volver a ingresar"):
        st.session_state.clear()
        st.rerun()
    st.stop()

col1, col2 = st.columns([2, 1])
with col1:
    st.title("Sistema Pukaray IA")
    st.caption("Filtro inmediato: Curso → Estudiante")
with col2:
    if st.button("Limpiar datos"):
        limpiar_formulario()
    if st.button("Salir"):
        salir_programa()

estudiantes = leer_hoja("Estudiantes")
entrevistadores = leer_hoja("Entrevistadores")
responsables = leer_hoja("Responsables_Apoyo")

# SELECTORES FUERA DEL FORMULARIO: esto permite que el filtro se actualice inmediatamente.
st.subheader("1. Curso y estudiante")

cursos = sorted({str(e.get("Curso", "")).strip() for e in estudiantes if str(e.get("Curso", "")).strip()})
curso_sel = st.selectbox("Curso", ["Seleccione curso"] + cursos, key="curso_sel")

estudiantes_filtrados = []
if curso_sel != "Seleccione curso":
    curso_norm = normalizar(curso_sel)
    estudiantes_filtrados = [
        e for e in estudiantes
        if normalizar(e.get("Curso", "")) == curso_norm and str(e.get("Nombre Estudiante", "")).strip()
    ]

nombres_estudiantes = [str(e.get("Nombre Estudiante", "")).strip() for e in estudiantes_filtrados]

if curso_sel != "Seleccione curso" and not nombres_estudiantes:
    st.warning("No hay estudiantes activos registrados para este curso en la hoja Estudiantes.")

estudiante_sel = st.selectbox(
    "Estudiante",
    ["Seleccione estudiante"] + nombres_estudiantes,
    key=f"estudiante_sel_{normalizar(curso_sel)}"
)

estudiante = {}
if estudiante_sel != "Seleccione estudiante":
    estudiante = next((e for e in estudiantes_filtrados if str(e.get("Nombre Estudiante", "")).strip() == estudiante_sel), {})

with st.container(border=True):
    c1, c2, c3 = st.columns(3)
    with c1:
        mostrar_dato("Curso", curso_sel if curso_sel != "Seleccione curso" else "")
    with c2:
        mostrar_dato("RUN", estudiante.get("RUN", ""))
    with c3:
        mostrar_dato("Estudiante", estudiante.get("Nombre Estudiante", ""))

st.divider()

# El resto NO necesita actualizar estudiantes en vivo.
st.subheader("2. Datos de entrevista")

fecha = st.date_input("Fecha entrevista", value=date.today(), format="DD/MM/YYYY", key="fecha_entrevista")
col_a, col_b = st.columns(2)
with col_a:
    hora = st.text_input("Hora entrevista", placeholder="Ej: 17:00 hrs", key="hora_entrevista")
    numero_entrevista = st.text_input("Número entrevista", placeholder="Ej: 001-2026", key="numero_entrevista")
with col_b:
    apoderado_nombre = st.text_input("Nombre apoderado entrevistado", key="apoderado_nombre")
    apoderado_relacion = st.text_input("Relación con estudiante", key="apoderado_relacion")
    apoderado_telefono = st.text_input("Teléfono apoderado", key="apoderado_telefono")
    apoderado_correo = st.text_input("Correo apoderado", key="apoderado_correo")

st.subheader("3. Participantes y apoyos")
nombres_entrevistadores = [e.get("Nombre Entrevistador", "") for e in entrevistadores if e.get("Nombre Entrevistador")]
entrevistadores_sel = st.multiselect("Entrevistadores participantes", nombres_entrevistadores, default=nombres_entrevistadores[:1], key="entrevistadores_sel")
entrevistadores_data = [e for e in entrevistadores if e.get("Nombre Entrevistador") in entrevistadores_sel]
resumen_ent = resumen_personas(entrevistadores_data, "Nombre Entrevistador", "Cargo", "Departamento")

with st.container(border=True):
    mostrar_dato("Entrevistadores", resumen_ent["nombres"])
    mostrar_dato("Cargos", resumen_ent["cargos"])
    mostrar_dato("Departamentos", resumen_ent["deptos"])

nombres_responsables = [r.get("Nombre Responsable", "") for r in responsables if r.get("Nombre Responsable")]
responsables_sel = st.multiselect("Responsables a cargo de ejecutar apoyos", nombres_responsables, default=nombres_responsables[:1], key="responsables_sel")
responsables_data = [r for r in responsables if r.get("Nombre Responsable") in responsables_sel]
resumen_resp = resumen_personas(responsables_data, "Nombre Responsable", "Cargo/Rol", "Área", "Tipo de Apoyo")

with st.container(border=True):
    mostrar_dato("Responsables", resumen_resp["nombres"])
    mostrar_dato("Roles", resumen_resp["cargos"])
    mostrar_dato("Áreas", resumen_resp["deptos"])

tipo_apoyo_extra = st.text_area("Ajuste o detalle del apoyo a ejecutar", value=resumen_resp["apoyos"], height=110, key="tipo_apoyo_extra")

st.subheader("4. Asistencia y antecedentes")
col_x, col_y = st.columns(2)
with col_x:
    asiste_apoderado = st.selectbox("Asiste apoderado", ["Sí", "No"], key="asiste_apoderado")
with col_y:
    asiste_estudiante = st.selectbox("Asiste estudiante", ["No", "Sí"], key="asiste_estudiante")

antecedentes = st.text_area("Antecedentes breves del caso", height=180, key="antecedentes")
incluir_rice = st.checkbox("Analizar antecedentes según RICE", value=True, key="incluir_rice")

generar = st.button("Generar documento y registrar seguimiento", type="primary")

if generar:
    if curso_sel == "Seleccione curso":
        st.error("Debe seleccionar un curso.")
    elif estudiante_sel == "Seleccione estudiante" or not estudiante:
        st.error("Debe seleccionar un estudiante del curso.")
    elif not entrevistadores_sel:
        st.error("Debe seleccionar al menos un entrevistador.")
    elif not responsables_sel:
        st.error("Debe seleccionar al menos un responsable de apoyo.")
    else:
        nombre_estudiante = estudiante.get("Nombre Estudiante", "")
        run = estudiante.get("RUN", "") or ""
        curso = curso_sel

        previas = contar_intervenciones_previas(nombre_estudiante)
        rice = analizar_rice(antecedentes, previas) if incluir_rice else {"categoria": ["No solicitado"], "normas": ["No solicitado"], "medidas": [], "alertas": [], "gravedad": "BAJA"}

        motivo, analisis, acuerdos = redactar_textos(antecedentes, resumen_resp["nombres"], tipo_apoyo_extra, rice)
        acuerdos_final = f"Análisis institucional:\n{analisis}\n\nAcuerdos o conclusiones:\n{acuerdos}"

        archivo = completar_plantilla({
            "nombre_estudiante": nombre_estudiante,
            "curso": curso,
            "fecha": fecha.strftime("%d.%m.%Y"),
            "hora": hora,
            "entrevistadores": resumen_ent["nombres"],
            "cargos_entrevistadores": resumen_ent["cargos"],
            "departamentos": resumen_ent["deptos"],
            "numero_entrevista": numero_entrevista,
            "asiste_apoderado": asiste_apoderado,
            "asiste_estudiante": asiste_estudiante,
        }, motivo, acuerdos_final)

        nombre_archivo = f"{limpiar_nombre_archivo(nombre_estudiante)}_{limpiar_nombre_archivo(curso)}_{fecha.strftime('%d-%m-%Y')}.docx"

        ahora = datetime.now()
        registrar({
            "fecha_registro": ahora.strftime("%d.%m.%Y"),
            "hora_registro": ahora.strftime("%H:%M:%S"),
            "curso": curso,
            "nombre_estudiante": nombre_estudiante,
            "run": run,
            "nombre_apoderado": apoderado_nombre,
            "relacion_apoderado": apoderado_relacion,
            "telefono_apoderado": apoderado_telefono,
            "correo_apoderado": apoderado_correo,
            "entrevistadores": resumen_ent["nombres"],
            "cargos_entrevistadores": resumen_ent["cargos"],
            "departamentos": resumen_ent["deptos"],
            "responsables_apoyo": resumen_resp["nombres"],
            "roles_responsables": resumen_resp["cargos"],
            "tipos_apoyo": tipo_apoyo_extra,
            "asiste_apoderado": asiste_apoderado,
            "asiste_estudiante": asiste_estudiante,
            "antecedentes": antecedentes,
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
        })

        st.success("Documento generado y seguimiento registrado correctamente.")
        st.info(f"Nivel de gravedad referencial: {rice.get('gravedad')}")
        st.info(f"Intervenciones previas registradas: {previas}")

        with st.expander("Vista previa del análisis RICE y texto generado", expanded=True):
            st.markdown("### Normas posiblemente incumplidas")
            for item in rice.get("normas", []):
                st.write(f"- {item}")
            st.markdown("### Medidas sugeridas")
            for item in rice.get("medidas", []):
                st.write(f"- {item}")
            st.markdown("### Motivo")
            st.write(motivo)
            st.markdown("### Análisis")
            st.write(analisis)
            st.markdown("### Acuerdos")
            st.write(acuerdos)

        st.download_button("Descargar Word listo para imprimir", archivo, file_name=nombre_archivo, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

st.divider()
st.subheader("Base de datos")
st.write("Si un curso no muestra estudiantes, revise la hoja Estudiantes. Columnas obligatorias: Curso, Nombre Estudiante, RUN, Estado.")
st.code(DB_PATH)
