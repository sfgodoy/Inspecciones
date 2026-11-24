"""
Lectura general de informes con IA (respaldo cuando ningún perfil reconoce el formato).

El modelo recibe el texto del PDF y la lista exacta de nombres de los catálogos de la planilla.
Devuelve una lista de observaciones con recinto, lugar, item y problema elegidos de esos catálogos.
Los valores que no existan en el catálogo se descartan (quedan vacíos y se marcan para revisión).
"""
import json
import os
import re

import pdfplumber

MODELO_POR_DEFECTO = "claude-sonnet-5-5"
MAX_CARACTERES_POR_LLAMADA = 120_000

INSTRUCCIONES = """Eres un asistente que extrae observaciones de informes de inspección de viviendas.

Tarea: devuelve TODAS las observaciones (defectos o hallazgos) del texto como una lista JSON.
Cada observación debe tener estas claves:
- "numero": número o identificador que el informe asigna a la observación (texto). Si no tiene, usa "".
- "recinto": nombre EXACTO de la lista RECINTO, o null si no corresponde o no está claro.
- "lugar": nombre EXACTO de la lista LUGAR, o null.
- "item": nombre EXACTO de la lista ITEM, o null.
- "problema": nombre EXACTO de la lista PROBLEMA, o null.
- "observacion": texto de la observación, en español, sin números de página ni firmas.
- "respuesta": respuesta o estado indicado por el constructor/revisor, o "" si no hay.

Reglas:
- Usa solo nombres que aparezcan exactamente en las listas. Si no hay coincidencia clara, usa null.
- No incluyas nombres de inspectores, supervisores, gerentes, ni datos de la carátula o del proyecto como observaciones.
- No inventes observaciones. Si un texto no describe un defecto o hallazgo, no lo incluyas.
- Responde SOLO con el JSON (una lista), sin texto antes ni después.

LISTA RECINTO:
{recinto}

LISTA LUGAR:
{lugar}

LISTA ITEM:
{item}

LISTA PROBLEMA:
{problema}

TEXTO DEL INFORME:
{texto}
"""


def texto_del_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join((p.extract_text() or "") for p in pdf.pages)


def trozos(texto, maximo=MAX_CARACTERES_POR_LLAMADA):
    return [texto[i:i + maximo] for i in range(0, len(texto), maximo)] or [""]


def construir_prompt(texto, catalogos):
    def lista(nombres):
        return "\n".join(f"- {n}" for n in nombres) or "(vacía)"
    return INSTRUCCIONES.format(
        recinto=lista(catalogos["recinto"]),
        lugar=lista(catalogos["lugar"]),
        item=lista(catalogos["item"]),
        problema=lista(catalogos["problema"]),
        texto=texto,
    )


def extraer_json(respuesta):
    """Saca la lista JSON aunque el modelo la envuelva en texto o en ```json."""
    m = re.search(r"\[.*\]", respuesta, re.S)
    if not m:
        return []
    try:
        datos = json.loads(m.group(0))
    except json.JSONDecodeError:
        return []
    return [d for d in datos if isinstance(d, dict)]


def validar(obs, catalogos):
    """Descarta valores que no estén en el catálogo exacto."""
    limpio = {
        "numero": str(obs.get("numero") or "").strip(),
        "observacion": str(obs.get("observacion") or "").strip(),
        "respuesta": str(obs.get("respuesta") or "").strip(),
    }
    for campo in ("recinto", "lugar", "item", "problema"):
        valor = obs.get(campo)
        limpio[campo] = valor if valor in catalogos[campo] else None
    return limpio


def extraer_con_ia(pdf_path, catalogos, api_key=None, modelo=MODELO_POR_DEFECTO, cliente=None):
    """Devuelve una lista de observaciones validadas. `cliente` permite inyectar un cliente (pruebas)."""
    if cliente is None:
        import anthropic
        api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            raise ValueError("Falta la clave ANTHROPIC_API_KEY para la lectura general.")
        cliente = anthropic.Anthropic(api_key=api_key)

    resultados = []
    for parte in trozos(texto_del_pdf(pdf_path)):
        respuesta = cliente.messages.create(
            model=modelo,
            max_tokens=16000,
            messages=[{"role": "user", "content": construir_prompt(parte, catalogos)}],
        )
        texto_respuesta = "".join(b.text for b in respuesta.content if getattr(b, "type", "") == "text")
        for obs in extraer_json(texto_respuesta):
            resultados.append(validar(obs, catalogos))
    return resultados
