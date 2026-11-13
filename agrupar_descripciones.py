"""
Agrupa descripciones que son el mismo producto escrito de distintas formas y propone una de ellas
(ya existente, sin reescribir nada) para usarla en todo el grupo.

- Se compara dentro del mismo tipo de existencia.
- Clave de comparación: sin mayúsculas, tildes, espacios, comillas, separador X, plural final
  ni palabras de relleno. Los números se mantienen exactos: 4" y 14" nunca se juntan.
- MISMA CLAVE: grupo seguro.  PARECIDA: mismos números y casi las mismas palabras; se revisa.
- Descripción propuesta: la de más compras; si empatan, la más corta. El usuario puede cambiarla.

Uso:  python agrupar_descripciones.py
"""
import re
import sys
import unicodedata
from collections import Counter
from difflib import SequenceMatcher
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ENTRADA = Path("Datos consolidados de compras_nivel1.xlsx")
SALIDA = Path("Grupos de descripciones.xlsx")

RELLENO = {"DE", "DEL", "PARA", "CON", "X", "EL", "LA", "LOS", "LAS", "Y", "EN", "P", "C", "-", "+"}
SIMILITUD_MINIMA = 0.8   # entre dos palabras que solo cambian en su escritura


def sin_tildes(texto):
    return "".join(c for c in unicodedata.normalize("NFD", texto) if not unicodedata.combining(c))


def palabras(descripcion):
    texto = sin_tildes(str(descripcion).upper()).replace('"', " ").replace("''", " ")
    texto = re.sub(r"(?<=\d)\s*X\s*(?=[\d.])", " ", texto)          # 2 X 4 -> 2 4
    texto = re.sub(r"(?<=\d)(?=[A-Z])|(?<=[A-Z])(?=\d)", " ", texto)  # 30KG -> 30 KG
    salida = []
    for p in re.findall(r"[A-Z0-9][A-Z0-9./]*", texto):
        p = p.strip(".")
        if p in RELLENO or not p:
            continue
        if len(p) > 3 and p.isalpha() and p.endswith("S"):          # plural simple
            p = p[:-2] if re.search(r"[RLNDZJ]ES$", p) else p[:-1]   # AISLADORES -> AISLADOR, GUANTES -> GUANTE
        salida.append(p)
    return salida


def clave(descripcion):
    return " ".join(palabras(descripcion))


def numeros(descripcion):
    return sorted(p for p in palabras(descripcion) if re.search(r"\d", p))


def parecidas(a, b):
    """Mismos números y las palabras que difieren son variantes de escritura una a una
    (IGNIFUGOS/IGNIFUGAS). Las cortas (tallas, LB/LR, colores cortos) deben ser iguales y no puede sobrar ninguna."""
    if numeros(a) != numeros(b):
        return False
    solo_a = list((Counter(palabras(a)) - Counter(palabras(b))).elements())
    solo_b = list((Counter(palabras(b)) - Counter(palabras(a))).elements())
    if not solo_a or len(solo_a) != len(solo_b):
        return False
    for p in solo_a:
        par = next((q for q in solo_b if min(len(p), len(q)) >= 4
                    and SequenceMatcher(None, p, q).ratio() >= SIMILITUD_MINIMA), None)
        if par is None:
            return False
        solo_b.remove(par)
    return True


def agrupar(tabla):
    """Une por clave idéntica y luego por similitud (union-find), dentro de cada tipo."""
    descripciones = (tabla.groupby(["tipo_existencia", "descripcion_n1"]).size()
                     .rename("filas").reset_index())
    padre = list(range(len(descripciones)))

    def raiz(i):
        while padre[i] != i:
            padre[i] = padre[padre[i]]
            i = padre[i]
        return i

    descripciones["clave"] = descripciones["descripcion_n1"].map(clave)
    motivo = {}
    for _, g in descripciones.groupby(["tipo_existencia", "clave"]):
        ids = list(g.index)
        for i in ids[1:]:
            padre[raiz(i)] = raiz(ids[0])
            motivo[i] = motivo[ids[0]] = "MISMA CLAVE"
    for _, g in descripciones.groupby("tipo_existencia"):
        ids = list(g.index)
        for x, i in enumerate(ids):
            for j in ids[x + 1:]:
                if raiz(i) != raiz(j) and parecidas(g.at[i, "descripcion_n1"], g.at[j, "descripcion_n1"]):
                    padre[raiz(j)] = raiz(i)
                    motivo.setdefault(i, "PARECIDA")
                    motivo.setdefault(j, "PARECIDA")
    descripciones["raiz"] = [raiz(i) for i in descripciones.index]
    tam = descripciones["raiz"].map(descripciones["raiz"].value_counts())
    grupos = descripciones[tam > 1].copy()
    grupos["motivo"] = [motivo.get(i, "PARECIDA") for i in grupos.index]
    filas = []
    for n, (_, g) in enumerate(grupos.groupby("raiz", sort=False), 1):
        elegida = g.sort_values(["filas", "descripcion_n1"], ascending=[False, True],
                                key=lambda s: s if s.name == "filas" else s.str.len()).iloc[0]["descripcion_n1"]
        confianza = "MISMA CLAVE" if (g["motivo"] == "MISMA CLAVE").all() and g["clave"].nunique() == 1 else "PARECIDA"
        for _, r in g.sort_values("descripcion_n1").iterrows():
            filas.append({"grupo": n, "confianza": confianza, "tipo_existencia": r["tipo_existencia"],
                          "descripcion_n1": r["descripcion_n1"], "filas": r["filas"],
                          "propuesta": "SI" if r["descripcion_n1"] == elegida else None,
                          "descripcion_elegida": elegida, "mismo_producto": None})
    resultado = pd.DataFrame(filas)
    orden = {"MISMA CLAVE": 0, "PARECIDA": 1}
    resultado["_o"] = resultado["confianza"].map(orden)
    grupos_orden = resultado.groupby("grupo")["_o"].first().sort_values(kind="stable").index
    renumerar = {g: i for i, g in enumerate(grupos_orden, 1)}
    resultado["grupo"] = resultado["grupo"].map(renumerar)
    return resultado.sort_values(["grupo", "descripcion_n1"]).drop(columns="_o")


def guardar(resultado):
    if SALIDA.exists():
        sys.exit(f"{SALIDA} ya existe y puede tener respuestas; bórralo a mano si quieres rehacerlo")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Grupos"
    columnas = list(resultado.columns)
    ws.append(columnas)
    for fila in resultado.itertuples(index=False):
        ws.append([None if pd.isna(v) else v for v in fila])
    anchos = {"descripcion_n1": 65, "descripcion_elegida": 65, "tipo_existencia": 14, "confianza": 14,
              "mismo_producto": 16, "propuesta": 10}
    colores = ["FFFFFF", "EAF1FB"]
    for fila in ws.iter_rows():
        for celda in fila:
            celda.font = Font(name="Arial", size=10)
        if fila[0].row > 1:
            color = colores[fila[0].value % 2]
            for celda in fila:
                celda.fill = PatternFill("solid", fgColor=color)
            if fila[columnas.index("propuesta")].value == "SI":
                fila[columnas.index("descripcion_n1")].font = Font(name="Arial", size=10, bold=True)
    for celda in ws[1]:
        celda.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")
        celda.alignment = Alignment(vertical="center")
    for i, campo in enumerate(columnas, 1):
        ws.column_dimensions[get_column_letter(i)].width = anchos.get(campo, 9)
    amarillo = PatternFill("solid", fgColor="FFF2CC")
    for celda in ws[get_column_letter(columnas.index("mismo_producto") + 1)][1:]:
        celda.fill = amarillo
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = ws.dimensions
    leeme = wb.create_sheet("Léeme", 0)
    for linea in [
        "Grupos de descripciones que parecen el mismo producto",
        "Cada grupo propone una descripción que ya existe (en negrita, propuesta = SI) para usarla en todo el grupo.",
        "MISMA CLAVE: solo cambian mayúsculas, tildes, espacios, comillas, la X de medidas, plurales o 'DE/PARA/CON'.",
        "PARECIDA: mismos números y casi las mismas palabras; revisar con cuidado.",
        "mismo_producto: escribe NO en la fila que no pertenece al grupo. Vacío = de acuerdo.",
        "Para elegir otra descripción del grupo, mueve el SI de la columna propuesta a esa fila.",
    ]:
        leeme.append([linea])
    leeme.column_dimensions["A"].width = 110
    wb.save(SALIDA)


if __name__ == "__main__":
    tabla = pd.read_excel(ENTRADA)
    resultado = agrupar(tabla)
    guardar(resultado)
    resumen = resultado.groupby("confianza").agg(grupos=("grupo", "nunique"), descripciones=("grupo", "size"))
    print(resumen.to_string())
    print(f"Guardado {SALIDA}")
