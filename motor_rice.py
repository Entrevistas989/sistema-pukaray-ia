import unicodedata

def normalizar_texto(texto):
    texto = str(texto or "").lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    return texto

def analizar_rice(texto, cantidad_intervenciones_previas=0):
    t = normalizar_texto(texto)
    categorias, normas, medidas, alertas = [], [], [], []
    puntaje = 0

    def agregar(categoria, norma, medida_lista, puntos):
        if categoria not in categorias:
            categorias.append(categoria)
        if norma not in normas:
            normas.append(norma)
        for medida in medida_lista:
            if medida not in medidas:
                medidas.append(medida)
        return puntos

    if any(p in t for p in ["pego", "pega", "golpea", "golpe", "agrede", "agresion", "patada", "empuja", "zancadilla", "lesion"]):
        puntaje += agregar(
            "Falta muy grave en torno al respeto",
            "Agrede física, verbal, gestualmente y/o a través de medios tecnológicos a personas de la comunidad escolar.",
            [
                "Citación inmediata de apoderado.",
                "Registro formal de los antecedentes.",
                "Reflexión formativa con el estudiante.",
                "Acción reparatoria hacia el estudiante afectado.",
                "Seguimiento desde Convivencia Escolar.",
                "Evaluar medidas de resguardo según el contexto."
            ],
            4
        )

    if any(p in t for p in ["insulto", "insulta", "ofende", "garabato", "groseria", "hiriente", "amenaza"]):
        puntaje += agregar(
            "Falta muy grave en torno al respeto",
            "Insulta y/u ofende en conversación o discusión mediante garabatos, frases o palabras hirientes.",
            [
                "Entrevista con apoderado.",
                "Reflexión formativa sobre buen trato.",
                "Compromiso escrito de mejora conductual.",
                "Seguimiento de Convivencia Escolar."
            ],
            4
        )

    if any(p in t for p in ["interrumpe", "interrupciones", "disruptiva", "disruptivo", "grita", "ruidos", "impide clase", "molesta", "desafiante"]):
        puntaje += agregar(
            "Falta grave en torno al respeto",
            "Impide el desarrollo de la clase o perturba el ambiente adecuado de trabajo mediante ruidos, gestos o verbalizaciones.",
            [
                "Entrevista con apoderado.",
                "Reflexión formativa.",
                "Compromiso escrito.",
                "Seguimiento de Convivencia Escolar."
            ],
            2
        )

    if any(p in t for p in ["burla", "burlas", "burlesco", "apodo", "sobrenombre"]):
        puntaje += agregar(
            "Falta grave en torno al respeto",
            "Se burla de miembros de la comunidad educativa o insta a otros a burlarse mediante apodos, sobrenombres o burlas de distintos tipos.",
            [
                "Entrevista con apoderado.",
                "Acción reparatoria.",
                "Reflexión formativa sobre respeto y buen trato.",
                "Seguimiento de convivencia escolar."
            ],
            2
        )

    if cantidad_intervenciones_previas >= 2:
        alertas.append("Existe reiteración en el historial del estudiante. Evaluar agravante según RICE.")
        puntaje += 2
    elif cantidad_intervenciones_previas == 1:
        alertas.append("Existe un antecedente previo registrado. Considerar seguimiento preventivo.")
        puntaje += 1

    if not categorias:
        categorias.append("Revisión profesional requerida")
        normas.append("No se detectó coincidencia automática suficiente. El equipo debe revisar el caso con los antecedentes completos.")
        medidas.extend(["Entrevista formativa.", "Registro de antecedentes.", "Seguimiento según evolución del caso."])

    gravedad = "ALTA" if puntaje >= 6 else "MODERADA" if puntaje >= 2 else "BAJA"

    return {
        "categoria": categorias,
        "normas": normas,
        "medidas": medidas,
        "alertas": alertas,
        "gravedad": gravedad
    }
