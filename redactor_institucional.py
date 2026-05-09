import unicodedata

def normalizar_texto(texto):
    texto = str(texto or "").lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    return texto

def mejorar_antecedentes(texto):
    texto_original = (texto or "").strip()
    t = normalizar_texto(texto_original)

    if not texto_original:
        return "No se registran antecedentes específicos en el campo de relato breve."

    frases = []

    if any(p in t for p in ["pego", "pega", "golpea", "golpe", "agrede", "agresion", "patada", "empuja", "zancadilla"]):
        frases.append(
            "Se registran antecedentes asociados a una interacción física inadecuada hacia otro estudiante, "
            "situación que requiere ser abordada desde el enfoque formativo, resguardando la integridad de los involucrados "
            "y promoviendo la reparación del daño causado."
        )

    if any(p in t for p in ["insulto", "insulta", "garabato", "groseria", "ofende", "amenaza", "hiriente"]):
        frases.append(
            "Asimismo, se informa el uso de lenguaje verbal inadecuado u ofensivo hacia un compañero, "
            "lo que resulta contrario a las normas de respeto y buen trato promovidas por la comunidad educativa."
        )

    if any(p in t for p in ["interrumpe", "interrupciones", "molesta", "disruptiv", "grita", "ruidos", "no deja hacer clase"]):
        frases.append(
            "Se observan conductas disruptivas durante el desarrollo de la clase, las que interfieren en el adecuado ambiente pedagógico "
            "y dificultan la continuidad de las actividades escolares."
        )

    if any(p in t for p in ["desafiante", "desafia", "responde mal", "no obedece", "no sigue instrucciones", "se niega"]):
        frases.append(
            "También se advierten dificultades para acatar instrucciones y responder adecuadamente a las indicaciones entregadas por los adultos responsables."
        )

    if any(p in t for p in ["burla", "burlas", "burlesco", "apodo", "sobrenombre", "se rie"]):
        frases.append(
            "Se consignan actitudes burlescas o verbalizaciones inadecuadas hacia integrantes del grupo curso, situación que puede afectar la convivencia escolar."
        )

    if any(p in t for p in ["reiterado", "reiterada", "varias veces", "constantemente", "frecuentemente", "siempre", "anotaciones"]):
        frases.append(
            "Los antecedentes señalan reiteración de la conducta, por lo que se estima necesario fortalecer las estrategias de seguimiento y acompañamiento institucional."
        )

    if not frases:
        frases.append(
            "Se informa una situación asociada al proceso formativo y de convivencia escolar del estudiante, la cual requiere entrevista, análisis de antecedentes y seguimiento institucional."
        )
        frases.append(f"Antecedente entregado por el funcionario: {texto_original}")

    return "\n\n".join(frases)
