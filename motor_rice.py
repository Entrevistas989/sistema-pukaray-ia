
def analizar_rice(texto, cantidad_intervenciones_previas=0):
    texto = (texto or "").lower()
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

    if any(p in texto for p in ["interrumpe", "interrupciones", "disruptiva", "disruptivo", "grita", "ruidos", "impide clase", "desafiante", "molesta", "burlesco", "se burla", "burlas", "apodo", "sobrenombre"]):
        puntaje += agregar(
            "Falta grave en torno al respeto",
            "Impide el desarrollo de la clase o perturba el ambiente adecuado de trabajo mediante ruidos, gestos o verbalizaciones.",
            ["Entrevista con apoderado", "Reflexión formativa", "Compromiso escrito", "Seguimiento de Convivencia Escolar"],
            2,
        )

    if any(p in texto for p in ["zancadilla", "empujón", "empujon", "patada", "golpe", "escupe", "escupitajo", "juego brusco", "caída", "caida"]):
        puntaje += agregar(
            "Falta leve o grave según daño y contexto",
            "Juego brusco o acción física que causa daño leve; evaluar agravantes, reiteración e intencionalidad.",
            ["Entrevista apoderado", "Acción reparatoria", "Seguimiento conductual", "Evaluar derivación a Convivencia Escolar"],
            2,
        )

    if any(p in texto for p in ["agresión", "agresion", "agrede", "amenaza", "lesión", "lesion", "daño físico", "daño psicológico", "pelea", "golpiza"]):
        puntaje += agregar(
            "Falta muy grave en torno al respeto",
            "Agresión física, verbal, gestual o psicológica hacia miembros de la comunidad escolar.",
            ["Citación inmediata de apoderado", "Derivación a Convivencia Escolar", "Medidas de protección y acompañamiento", "Evaluar procedimiento disciplinario"],
            4,
        )

    if cantidad_intervenciones_previas >= 2:
        alertas.append("Existe reiteración en el historial. Evaluar agravante según RICE.")
        puntaje += 2

    if not categorias:
        categorias.append("Sin coincidencia normativa automática clara")
        normas.append("Se requiere revisión profesional del caso según antecedentes completos.")
        medidas.extend(["Entrevista formativa", "Registro de antecedentes", "Seguimiento según evolución del caso"])

    gravedad = "ALTA" if puntaje >= 6 else "MODERADA" if puntaje >= 2 else "BAJA"
    return {"categoria": categorias, "normas": normas, "medidas": medidas, "alertas": alertas, "gravedad": gravedad}
