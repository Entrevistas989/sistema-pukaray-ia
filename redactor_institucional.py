
def mejorar_antecedentes(texto):
    texto_original = (texto or "").strip()
    t = texto_original.lower()

    if not texto_original:
        return "No se registran antecedentes específicos en el campo de relato breve."

    frases = []

    if any(p in t for p in ["interrumpe", "interrupciones", "molesta", "disruptiv", "grita", "ruidos", "no deja hacer clase"]):
        frases.append(
            "Se registran antecedentes asociados a conductas disruptivas durante el desarrollo de la clase, "
            "las que interfieren en el adecuado ambiente pedagógico y dificultan la continuidad de las actividades escolares."
        )

    if any(p in t for p in ["desafiante", "desafía", "desafia", "responde mal", "no obedece", "no sigue instrucciones", "se niega"]):
        frases.append(
            "Asimismo, se observan dificultades para acatar instrucciones y responder adecuadamente a las indicaciones entregadas por los adultos responsables."
        )

    if any(p in t for p in ["burla", "burlas", "burlesco", "apodo", "sobrenombre", "molesta compañeros", "se rie", "se ríe"]):
        frases.append(
            "Se consignan además actitudes burlescas o verbalizaciones inadecuadas hacia integrantes del grupo curso, "
            "situación que puede afectar la convivencia escolar y el bienestar socioemocional de los estudiantes involucrados."
        )

    if any(p in t for p in ["pega", "golpea", "golpe", "agrede", "agresión", "agresion", "patada", "empuja", "zancadilla"]):
        frases.append(
            "Se describen antecedentes vinculados a una interacción física inadecuada con otro estudiante, "
            "situación que requiere abordaje formativo, resguardo de los involucrados y seguimiento por parte del establecimiento."
        )

    if any(p in t for p in ["insulta", "garabato", "grosería", "groseria", "ofende", "amenaza"]):
        frases.append(
            "También se reporta uso de lenguaje verbal inadecuado u ofensivo, aspecto que resulta contrario a las normas de respeto promovidas por la comunidad educativa."
        )

    if any(p in t for p in ["reiterado", "reiterada", "varias veces", "constantemente", "frecuentemente", "siempre", "muchas anotaciones", "anotaciones"]):
        frases.append(
            "Los antecedentes señalan reiteración de la conducta, por lo que se estima necesario fortalecer las estrategias de seguimiento y acompañamiento institucional."
        )

    if not frases:
        return (
            "Se informa situación referida al proceso formativo y de convivencia escolar del estudiante, "
            "la cual requiere ser abordada mediante entrevista, análisis de antecedentes y seguimiento institucional. "
            f"Antecedente entregado por el funcionario: {texto_original}"
        )

    return " ".join(frases) + (
        " Estos antecedentes se presentan para conocimiento del apoderado, con el propósito de establecer acuerdos de apoyo, "
        "acompañamiento y mejora del proceso formativo del estudiante."
    )
