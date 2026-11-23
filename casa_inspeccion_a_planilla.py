#!/usr/bin/env python3
"""
Carga de informes de inspección a la planilla Excel de requerimientos.

Uso:
  python casa_inspeccion_a_planilla.py informe.pdf planilla_base.xlsx salida.xlsx --perfil perfiles/X.json [--revisar revisar.csv]
  python casa_inspeccion_a_planilla.py informe.pdf --explorar --perfil perfiles/X.json

El perfil (JSON) describe el formato del informe. Hay dos tipos de lectura:
  "posicional": tabla con columnas en posiciones fijas (campos + columnas_x).
  "bloques":   cada observación es un bloque de texto que empieza con "N°  RECINTO" (inicio_bloque).

Requiere:  pip install pdfplumber openpyxl
"""
import argparse
import csv
import json
import re
import sys
import unicodedata
from collections import Counter

import openpyxl
import pdfplumber


def norm(s):
    s = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in s if not unicodedata.combining(c)).lower()


def clean(s):
    return re.sub(r"\s+", " ", (s or "")).strip()


class Perfil:
    def __init__(self, ruta):
        with open(ruta, encoding="utf-8") as f:
            d = json.load(f)
        self.nombre = d.get("nombre", ruta)
        self.proyecto = d.get("proyecto", "")
        self.tipo = d.get("tipo_lectura", "posicional")
        self.ignorar = [re.compile(p) for p in d.get("ignorar_lineas", [])]
        self.encabezados = [(re.compile(p), s) for p, s in d.get("encabezados_seccion", [])]
        self.recinto_encabezado = re.compile(d["recinto_encabezado"]) if d.get("recinto_encabezado") else None
        self.fin_bloques = re.compile(d["fin_bloques"]) if d.get("fin_bloques") else None
        self.seccion_recinto = d.get("seccion_recinto")
        self.secciones = d.get("secciones", {"rev0": 0})
        self.campos = d.get("campos", ["numero", "recinto", "observacion", "respuesta"])
        self.columnas_x = d.get("columnas_x", [])
        self.inicio_bloque = re.compile(d["inicio_bloque"]) if d.get("inicio_bloque") else None
        self.fin_bloques = re.compile(d["fin_bloques"]) if d.get("fin_bloques") else None
        self.marca_categoria = d.get("marca_categoria", "•")
        self.marca_defecto = d.get("marca_defecto", "−")
        self.recinto_reglas = d.get("recinto_reglas", [])
        self.problema_reglas = d.get("problema_reglas", [])
        self.categoria_reglas = d.get("categoria_reglas", [])


def recinto_de(perfil, texto):
    t = norm(texto)
    mejor = None
    for clave, nombre in perfil.recinto_reglas:
        i = t.find(norm(clave))
        if i != -1 and (mejor is None or i < mejor[0]):
            mejor = (i, nombre)
    return mejor[1] if mejor else None


def problema_de(perfil, texto):
    t = norm(texto)
    for clave, nombre in perfil.problema_reglas:
        if norm(clave) in t:
            return nombre
    return None


def categoria_de(perfil, texto):
    t = norm(texto)
    for r in perfil.categoria_reglas:
        if any(norm(c) in t for c in r["claves"]):
            return r["lugar"], r["item"]
    return None, None


# ---------------------------------------------------------------- lectura posicional
def leer_posicional(pdf_path, perfil):
    filas = []
    seccion = None
    recinto_heading = None
    cur = None
    cx = perfil.columnas_x
    campos = perfil.campos
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lineas = {}
            for w in page.extract_words():
                lineas.setdefault(round(w["top"] / 3), []).append(w)
            for _, ws in sorted(lineas.items()):
                ws = sorted(ws, key=lambda w: w["x0"])
                texto = " ".join(w["text"] for w in ws)
                if any(p.search(texto) for p in perfil.ignorar):
                    continue
                if perfil.fin_bloques and perfil.fin_bloques.search(texto):
                    seccion, cur, recinto_heading = None, None, None
                    continue
                if perfil.recinto_encabezado:
                    m = perfil.recinto_encabezado.search(texto)
                    if m:
                        recinto_heading = clean(m.group(1))
                        seccion, cur = perfil.seccion_recinto, None
                        continue
                cambio = False
                for patron, sec in perfil.encabezados:
                    if patron.search(texto):
                        seccion, cur, cambio = sec, None, True
                        break
                if cambio:
                    continue
                if seccion not in perfil.secciones:
                    cur = None
                    continue
                if "N°" in [w["text"] for w in ws] and "RECINTO" in texto:
                    cur = None
                    continue
                cols = [""] * len(campos)
                for w in ws:
                    x = w["x0"]
                    idx = sum(1 for limite in cx if x >= limite)
                    cols[idx] = (cols[idx] + " " + w["text"]).strip()
                num = cols[0].replace("ª", "a").replace(" ", "")
                if re.match(r"^\d", num) and not re.fullmatch(r"\d+(-\d+)?[a-z]?|\d+-\d*", num):
                    cur = None  # títulos internos del informe (ej. "4.RECINTO", "5.CONCLUSIONES")
                    continue
                if re.match(r"^\d", num):
                    if cur and cur["label"].endswith("-") and re.fullmatch(r"\d+[a-z]?", num):
                        cur["label"] += num
                    else:
                        cur = {"sec": perfil.secciones[seccion], "label": num, "recinto_heading": recinto_heading}
                        for c in campos[1:]:
                            cur[c] = ""
                        filas.append(cur)
                    cols[0] = ""
                elif num == "-" and cur and not cur["label"].endswith("-"):
                    cur["label"] += "-"
                    cols[0] = ""
                if cur is None:
                    continue
                for c, val in zip(campos[1:], cols[1:]):
                    cur[c] = (cur[c] + " " + val).strip()
    for f in filas:
        f["num"] = int(re.match(r"\d+", f["label"]).group(0))
    return filas


# ---------------------------------------------------------------- lectura por bloques de texto
def leer_bloques(pdf_path, perfil):
    lineas_pdf = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for linea in (page.extract_text() or "").splitlines():
                lineas_pdf.append(linea.rstrip())
    filas = []
    cur = None
    fin = False
    for i, linea in enumerate(lineas_pdf):
        if any(p.search(linea) for p in perfil.ignorar):
            continue
        if perfil.fin_bloques and perfil.fin_bloques.search(linea):
            fin = True
            cur = None
        if fin:
            continue
        m = perfil.inicio_bloque.match(linea) if perfil.inicio_bloque else None
        siguiente = next((l for l in lineas_pdf[i + 1:i + 3] if l.strip()), "")
        if m and re.match(r"^\s*Obs\.-", siguiente):
            cur = {"sec": 0, "label": m.group(1), "recinto": clean(m.group(2)),
                   "categoria": "", "defecto": "", "observacion": "", "respuesta": ""}
            filas.append(cur)
            continue
        if cur is None or re.match(r"^\s*Obs\.-", linea):
            continue
        t = linea.strip()
        if not t:
            continue
        cur["observacion"] = clean(cur["observacion"] + " " + t)
        if t.startswith(perfil.marca_categoria):
            cur["categoria"] = clean(cur["categoria"] + " " + t.lstrip(perfil.marca_categoria))
        elif t.startswith(perfil.marca_defecto):
            cur["defecto"] = clean(cur["defecto"] + " " + t.lstrip(perfil.marca_defecto))
        elif cur["defecto"]:
            cur["defecto"] = clean(cur["defecto"] + " " + t)
        else:
            cur["categoria"] = clean(cur["categoria"] + " " + t)
    for f in filas:
        f["num"] = int(f["label"])
        f["observacion"] = clean(re.sub(r"[•−]", " ", f["observacion"]))
    return filas


# ---------------------------------------------------------------- exploración
def explorar(pdf_path, perfil):
    print(f"Perfil: {perfil.nombre} (lectura: {perfil.tipo})")
    with pdfplumber.open(pdf_path) as pdf:
        print(f"Páginas: {len(pdf.pages)}")
        if perfil.tipo == "posicional":
            x_inicio = Counter()
            for page in pdf.pages:
                for w in page.extract_words():
                    if 90 < w["top"] < 740:
                        x_inicio[round(w["x0"] / 5) * 5] += 1
            print("Posiciones x más frecuentes (para columnas_x):")
            for x, n in sorted(x_inicio.items(), key=lambda t: -t[1])[:12]:
                print(f"   x≈{x}: {n} palabras")
        else:
            n = sum(1 for page in pdf.pages for l in (page.extract_text() or "").splitlines()
                    if perfil.inicio_bloque and perfil.inicio_bloque.match(l))
            print(f"Líneas que parecen inicio de bloque: {n}")
        print("Encabezados de sección encontrados:")
        for i, page in enumerate(pdf.pages, start=1):
            for linea in (page.extract_text() or "").splitlines():
                for patron, sec in perfil.encabezados:
                    if patron.search(linea):
                        print(f"   página {i}: '{linea.strip()}' -> {sec}")


def procesar(pdf_path, plantilla, salida, perfil_path):
    """Carga un informe en la planilla. Devuelve (total_filas, filas_para_revisar: lista de dicts)."""
    perfil = Perfil(perfil_path)
    filas = leer_bloques(pdf_path, perfil) if perfil.tipo == "bloques" else leer_posicional(pdf_path, perfil)
    if not filas:
        raise ValueError("No se encontraron observaciones. Revisa que el perfil corresponda al informe.")

    wb = openpyxl.load_workbook(plantilla)

    def catalogo(nombre):
        ws = wb[nombre]
        return {str(ws.cell(r, 2).value).strip(): ws.cell(r, 1).value
                for r in range(2, ws.max_row + 1) if ws.cell(r, 2).value}

    cat = {"rec": catalogo("Recinto"), "lug": catalogo("Lugar"),
           "item": catalogo("Item"), "prob": catalogo("Problema")}

    def codigo(tipo, nombre):
        if not nombre:
            return None
        if nombre not in cat[tipo]:
            raise ValueError(f"'{nombre}' no existe en el catálogo '{tipo}'. Corrige el perfil.")
        return cat[tipo][nombre]

    registros = []
    for f in filas:
        recinto_txt = f.get("recinto") or f.get("recinto_heading") or ""
        obs = f.get("observacion", "")
        if perfil.tipo == "bloques":
            cat_txt, prob_txt = f["categoria"], f["defecto"] or obs
        else:
            cat_txt = f.get("elemento") or f.get("categoria") or obs
            prob_txt = obs
        rec = recinto_de(perfil, recinto_txt)
        lugar, item = categoria_de(perfil, cat_txt)
        prob = problema_de(perfil, prob_txt)
        faltantes = [n for n, v in (("Recinto", rec), ("Lugar", lugar), ("Item", item), ("Problema", prob)) if not v]
        comentario = []
        for campo, etiqueta in (("estado", "Chequeo"), ("respuesta", "Respuesta")):
            if f.get(campo):
                comentario.append(f"{etiqueta}: {f[campo]}")
        if faltantes:
            comentario.append("[REVISAR: " + ", ".join(faltantes) + "]")
        registros.append({
            "orden": (rec or "ZZZ", f["sec"], f["num"], f["label"]),
            "rec": rec, "lugar": lugar, "item": item, "prob": prob,
            "desc": obs, "comentario": " ".join(comentario) or None,
            "etiqueta": f["label"], "faltantes": faltantes,
        })
    registros.sort(key=lambda r: r["orden"])

    ws = wb["Requerimientos"]
    if ws.max_row >= 3:
        ws.delete_rows(3, ws.max_row - 2)
    for i, r in enumerate(registros, start=3):
        ws.cell(i, 1, perfil.proyecto)
        ws.cell(i, 8, codigo("rec", r["rec"]))
        ws.cell(i, 9, codigo("lug", r["lugar"]))
        ws.cell(i, 10, codigo("item", r["item"]))
        ws.cell(i, 11, codigo("prob", r["prob"]))
        ws.cell(i, 12, r["desc"])
        ws.cell(i, 18, r["comentario"])
    wb.save(salida)
    pendientes = [r for r in registros if r["faltantes"]]
    return len(registros), pendientes


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf")
    ap.add_argument("plantilla", nargs="?")
    ap.add_argument("salida", nargs="?")
    ap.add_argument("--perfil", required=True)
    ap.add_argument("--explorar", action="store_true")
    ap.add_argument("--revisar", help="CSV con las filas que requieren revisión")
    args = ap.parse_args()

    if args.explorar:
        explorar(args.pdf, Perfil(args.perfil))
        return
    if not args.plantilla or not args.salida:
        ap.error("faltan la planilla base y el archivo de salida")

    try:
        total, pendientes = procesar(args.pdf, args.plantilla, args.salida, args.perfil)
    except ValueError as e:
        sys.exit(str(e))
    print(f"Filas escritas: {total}")
    print(f"Filas con [REVISAR]: {len(pendientes)}")
    print(f"Archivo: {args.salida}")
    if args.revisar:
        with open(args.revisar, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh, delimiter=";")
            w.writerow(["observacion", "faltan", "texto"])
            for r in pendientes:
                w.writerow([r["etiqueta"], ", ".join(r["faltantes"]), r["desc"]])
        print(f"Lista de revisión: {args.revisar}")


if __name__ == "__main__":
    main()
