import re
import unicodedata
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import streamlit as st
from docx import Document
from openpyxl import load_workbook

from motor_rice import analizar_rice
from redactor_institucional import mejorar_antecedentes

TEMPLATE_PATH = "plantilla_ficha_entrevista_apoderado.docx"
DB_PATH = "base_datos_pukaray.xlsx"

st.set_page_config(page_title="Sistema Pukaray IA", page_icon="📄", layout="centered")
LOGO_PATH = "logo_pukaray.png"
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

def limpiar_para_word(texto):
    return str(texto or "").replace("\\n", "\n").replace("\\t", "\t")

def limpiar_datos():
    st.session_state["reset_form"] = st.session_state.get("reset_form", 0) + 1
    st.rerun()
    for clave in claves:
        if clave in st.session_state:
            del st.session_state[clave]
    st.rerun()

def salir():
    st.session_state["salir"] = True
    st.rerun()

def leer_hoja(nombre_hoja):
    wb = load_workbook(DB_PATH, data_only=True)
    ws = wb[nombre_hoja]
    headers = [str(c.value or "").strip() for c in ws[1]]
    registros = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        item = {headers[i]: (row[i] if i < len(row) else "") for i in range(len(headers))}
        if str(item.get("Estado", "Activo") or "Activo").strip().lower() == "activo":
            registros.append(item)
    return registros

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
    categorias = rice.get("categoria", [])
    normas = rice.get("normas", [])
    medidas = rice.get("medidas", [])
    alertas = rice.get("alertas", [])

    motivo = (
        "Se realiza entrevista de apoderado con el propósito de informar antecedentes asociados al proceso formativo y de convivencia escolar del estudiante.\n\n"
        "ANTECEDENTES INFORMADOS:\n"
        f"{antecedentes_mejorados}"
    )

    analisis_partes = [
        "ANÁLISIS INSTITUCIONAL",
        "1. Los antecedentes descritos evidencian una situación que requiere abordaje formativo, resguardo de la convivencia escolar y coordinación con la familia.",
        "2. Se recomienda fortalecer la reflexión del estudiante respecto de sus acciones, promoviendo la reparación del daño y el cumplimiento de las normas institucionales.",
        "3. Clasificación referencial según RICE:",
    ]
    for cat in categorias:
        analisis_partes.append(f"   • {cat}")

    analisis_partes.append("4. Normas posiblemente asociadas:")
    for norma in normas:
        analisis_partes.append(f"   • {norma}")

    acuerdos_partes = [
        "ACUERDOS Y CONCLUSIONES",
        "1. El apoderado toma conocimiento formal de los antecedentes expuestos durante la entrevista.",
        "2. Se acuerda reforzar desde el hogar normas de respeto, buen trato y resolución adecuada de conflictos.",
        "3. El establecimiento realizará seguimiento institucional del caso.",
    ]

    if responsables_apoyo:
        acuerdos_partes.append(f"4. Responsables de ejecución y seguimiento de apoyos: {responsables_apoyo}.")

    if tipo_apoyo:
        acuerdos_partes.append("5. Apoyos comprometidos:")
        for linea in str(tipo_apoyo).splitlines():
            if linea.strip():
                acuerdos_partes.append(f"   • {linea.strip()}")

    if medidas:
        acuerdos_partes.append("6. Medidas formativas sugeridas según análisis RICE:")
        for medida in medidas:
            acuerdos_partes.append(f"   • {medida}")

    acuerdos_partes.append(f"7. Nivel referencial de gravedad institucional: {rice.get('gravedad', 'BAJA')}.")

    if alertas:
        acuerdos_partes.append("8. Alertas para revisión del equipo:")
        for alerta in alertas:
            acuerdos_partes.append(f"   • {alerta}")

    return motivo, "\n".join(analisis_partes), "\n".join(acuerdos_partes)

def agregar_texto(celda, texto):
    celda.text = limpiar_para_word(texto)

def completar_plantilla(datos, motivo, analisis, acuerdos):
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
    agregar_texto(doc.tables[4].cell(0, 1), motivo)
    agregar_texto(doc.tables[5].cell(0, 1), f"{analisis}\n\n{acuerdos}")
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
        registro.get("asiste_apoderado"), registro.get("asiste_estudiante"),
        registro.get("antecedentes_originales"), registro.get("antecedentes_mejorados"),
        registro.get("motivo"), registro.get("analisis"), registro.get("acuerdos"),
        registro.get("categoria_rice"), registro.get("normas_rice"), registro.get("medidas_rice"), registro.get("alertas_rice"),
        registro.get("gravedad"), registro.get("archivo_generado"), registro.get("numero_entrevista"),
    ])
    wb.save(DB_PATH)

if st.session_state.get("salir"):
LOGO_PATH = "logo_pukaray.png"
st.image(LOGO_PATH, width=150)
    st.title("Sistema Pukaray IA")
    st.success("Sesión cerrada en este equipo.")
    st.write("Puede cerrar esta pestaña del navegador.")
    if st.button("Volver a ingresar"):
        st.session_state.clear()
        st.rerun()
    st.stop()

col1, col2 = st.columns([2, 1])
with col1:
    st.title("Sistema Pukaray IA")
    st.caption("Corrección RICE + texto ordenado + botones")
with col2:
    if st.button("Limpiar datos"):
        limpiar_datos()
    if st.button("Salir"):
        salir()

estudiantes = leer_hoja("Estudiantes")
entrevistadores = leer_hoja("Entrevistadores")
responsables = leer_hoja("Responsables_Apoyo")

cursos = sorted({str(e.get("Curso", "")).strip() for e in estudiantes if str(e.get("Curso", "")).strip()})
reset_form = st.session_state.get("reset_form", 0)
curso_sel = st.selectbox("Curso", ["Seleccione curso"] + cursos, index=0, key=f"curso_sel_{reset_form}")

estudiantes_filtrados = [e for e in estudiantes if curso_sel != "Seleccione curso" and normalizar(e.get("Curso", "")) == normalizar(curso_sel)]
nombres_estudiantes = [str(e.get("Nombre Estudiante", "")).strip() for e in estudiantes_filtrados]
estudiante_sel = st.selectbox("Estudiante",["Seleccione estudiante"] + nombres_estudiantes,index=0,key=f"estudiante_sel_{reset_form}_{normalizar(curso_sel)}")
estudiante = next((e for e in estudiantes_filtrados if str(e.get("Nombre Estudiante", "")).strip() == estudiante_sel), {})

fecha = st.date_input("Fecha entrevista", value=date.today(), format="DD/MM/YYYY")
hora = st.text_input("Hora entrevista", placeholder="Ej: 17:00 hrs", key="hora_entrevista")
numero_entrevista = st.text_input("Número entrevista", placeholder="Ej: 001-2026", key="numero_entrevista")

apoderado_nombre = st.text_input("Nombre apoderado entrevistado", key="apoderado_nombre")
apoderado_relacion = st.text_input("Relación con estudiante", key="apoderado_relacion")
apoderado_telefono = st.text_input("Teléfono apoderado", key="apoderado_telefono")
apoderado_correo = st.text_input("Correo apoderado", key="apoderado_correo")

nombres_entrevistadores = [e.get("Nombre Entrevistador", "") for e in entrevistadores if e.get("Nombre Entrevistador")]
entrevistadores_sel = st.multiselect("Entrevistadores participantes", nombres_entrevistadores, default=nombres_entrevistadores[:1])
entrevistadores_data = [e for e in entrevistadores if e.get("Nombre Entrevistador") in entrevistadores_sel]
resumen_ent = resumen_personas(entrevistadores_data, "Nombre Entrevistador", "Cargo", "Departamento")

nombres_responsables = [r.get("Nombre Responsable", "") for r in responsables if r.get("Nombre Responsable")]
responsables_sel = st.multiselect("Responsables a cargo de ejecutar apoyos", nombres_responsables, default=nombres_responsables[:1])
responsables_data = [r for r in responsables if r.get("Nombre Responsable") in responsables_sel]
resumen_resp = resumen_personas(responsables_data, "Nombre Responsable", "Cargo/Rol", "Área", "Tipo de Apoyo")

tipo_apoyo_extra = st.text_area("Ajuste o detalle del apoyo a ejecutar", value=resumen_resp["apoyos"], height=110, key="tipo_apoyo_extra")

asiste_apoderado = st.selectbox("Asiste apoderado", ["Sí", "No"], key="asiste_apoderado")
asiste_estudiante = st.selectbox("Asiste estudiante", ["No", "Sí"], key="asiste_estudiante")

antecedentes = st.text_area("Antecedentes breves del caso", height=160, placeholder="Ej: le pegó a otro compañero y lo insultó", key="antecedentes")

mejorar_texto = st.checkbox("Mejorar automáticamente la redacción institucional", value=True, key="mejorar_texto")
incluir_rice = st.checkbox("Analizar antecedentes según RICE", value=True, key="incluir_rice")

generar = st.button("Generar documento y registrar seguimiento", type="primary")

if generar:
    if curso_sel == "Seleccione curso" or estudiante_sel == "Seleccione estudiante":
        st.error("Debe seleccionar curso y estudiante.")
    else:
        nombre_estudiante = estudiante.get("Nombre Estudiante", "")
        run = estudiante.get("RUN", "") or ""
        curso = curso_sel

        antecedentes_mejorados = mejorar_antecedentes(antecedentes) if mejorar_texto else antecedentes
        previas = contar_intervenciones_previas(nombre_estudiante)
        rice = analizar_rice(f"{antecedentes} {antecedentes_mejorados}", previas) if incluir_rice else {"categoria": ["No solicitado"], "normas": ["No solicitado"], "medidas": [], "alertas": [], "gravedad": "BAJA"}

        motivo, analisis, acuerdos = redactar_textos(antecedentes_mejorados, resumen_resp["nombres"], tipo_apoyo_extra, rice)

        nombre_archivo = f"{limpiar_nombre_archivo(nombre_estudiante)}_{limpiar_nombre_archivo(curso)}_{fecha.strftime('%d-%m-%Y')}.docx"

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
        }, motivo, analisis, acuerdos)

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
        })

        st.success("Documento generado y seguimiento registrado correctamente.")
        st.markdown("### Antecedentes mejorados")
        st.write(antecedentes_mejorados)
        st.markdown("### RICE detectado")
        st.write(rice)
        st.download_button("Descargar Word listo para imprimir", archivo, file_name=nombre_archivo, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
