"""
Tabla de pares de descripciones parecidas, con sus precios, para decidir si son el mismo producto.

- Precio: unitario con IGV (total / cantidad), en su moneda y su unidad de compra.
  Se separan compras reales y cotizaciones; el precio de comparación es la mediana de las compras,
  o el de la cotización si la descripción no tiene compras.
- UNIR: solo cambia la escritura (mayúsculas, DE, plural, X, comillas) y el precio es parecido.
- REVISAR: misma escritura, pero el precio no se puede comparar (otra moneda u otra unidad),
  o es muy distinto y una de las dos es solo cotización.
- NO UNIR: el precio unitario de compras reales difiere más de PRECIO_DISTINTO veces.
- TALLA: solo cambia la talla (se resuelve al separarla). OTRO PRODUCTO: cambia una palabra que lo define.
El nombre final es una de las dos descripciones, la de más compras (si empatan, la más corta).

Uso:  python comparar_descripciones.py
"""
import re
from difflib import SequenceMatcher
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from agrupar_descripciones import clave, numeros, palabras, parecidas, sin_tildes

ENTRADA = Path("Datos consolidados de compras_nivel1.xlsx")
SALIDA = Path("Comparacion de descripciones.xlsx")
PRECIO_DISTINTO = 1.5      # veces
PARECIDO_CANDIDATO = 0.85  # similitud del texto para mostrar el par aunque cambie alguna palabra


def resumen_precios(tabla):
    """Por descripción: compras y cotizaciones por separado, unidad, moneda, proveedores y precios unitarios."""
    filas = {}
    for (tipo, d), g in tabla.groupby(["tipo_existencia", "descripcion_n1"]):
        unidades, monedas = sorted(set(g["unidad"].astype(str))), sorted(set(g["moneda"]))
        compras = g[g["tipo_registro"] == "COMPRA"]
        cotizaciones = g[g["tipo_registro"] == "COTIZACION"]
        una_moneda = len(monedas) == 1
        mediana = lambda s: round(s.median(), 2) if una_moneda and len(s) else None
        filas[(tipo, d)] = {
            "compras": len(compras), "cotizaciones": len(cotizaciones),
            "unidad": "/".join(unidades), "moneda": "/".join(monedas),
            "precio_compra": mediana(compras["precio_unitario"]),
            "minimo": round(compras["precio_unitario"].min(), 2) if una_moneda and len(compras) else None,
            "maximo": round(compras["precio_unitario"].max(), 2) if una_moneda and len(compras) else None,
            "precio_cotizacion": mediana(cotizaciones["precio_unitario"]),
            "proveedores": "; ".join(sorted(set(compras["proveedor"].dropna().astype(str).str.strip()))) or None,
        }
        filas[(tipo, d)]["precio"] = filas[(tipo, d)]["precio_compra"] or filas[(tipo, d)]["precio_cotizacion"]
    return filas


def que_cambia(a, b):
    """Explica en palabras qué diferencia a las dos descripciones."""
    pa = sin_tildes(a.upper()).replace('"', " ").split()
    pb = sin_tildes(b.upper()).replace('"', " ").split()
    solo_a = [p for p in pa if p not in pb]
    solo_b = [p for p in pb if p not in pa]
    if clave(a) == clave(b):
        cambios = []
        if any(p in ("DE", "PARA", "CON") for p in solo_a + solo_b):
            cambios.append("una tenía 'DE'")
        if any(p.rstrip("S") in [q.rstrip("S") for q in solo_b] for p in solo_a if p.endswith("S")) or \
           any(p.rstrip("S") in [q.rstrip("S") for q in solo_a] for p in solo_b if p.endswith("S")):
            cambios.append("singular / plural")
        if a.count('"') != b.count('"'):
            cambios.append('comillas de pulgada')
        if re.sub(r"\s", "", a) != re.sub(r"\s", "", b) and not cambios:
            cambios.append("la X o los espacios de la medida")
        return "Solo cambia la escritura: " + ", ".join(cambios or ["espacios / X"])
    return f"Cambian palabras: {' '.join(solo_a) or '-'}  vs  {' '.join(solo_b) or '-'}"


TALLAS = {"S", "M", "L", "XL", "XXL", "XXXL", "Y", "T", "TALLA", "T:"}


def solo_talla(a, b):
    """Las palabras que cambian son tallas (L, M, XL, T30, TALLA 30): se resuelve al separar la talla."""
    pa, pb = palabras(a), palabras(b)   # sin plural ni 'DE': GUANTE NITRILO TALLA M / GUANTES DE NITRILO TALLA L
    distintas = [p for p in pa if p not in pb] + [p for p in pb if p not in pa]
    es_talla = lambda p: p in TALLAS or re.fullmatch(r"T?\d{2}", p) is not None
    return bool(distintas) and all(es_talla(p) for p in distintas)


def comparar_precio(pa, pb):
    if pa["moneda"] != pb["moneda"] or "/" in pa["moneda"]:
        return None, "otra moneda"
    if pa["unidad"] != pb["unidad"]:
        return None, f"otra unidad ({pa['unidad']} vs {pb['unidad']})"
    if not pa["precio"] or not pb["precio"]:
        return None, "sin precio"
    veces = max(pa["precio"], pb["precio"]) / min(pa["precio"], pb["precio"])
    return round(veces, 2), None


def base_precio(pa, pb):
    if pa["precio_compra"] and pb["precio_compra"]:
        return "compra vs compra"
    return "incluye cotización"


def pares(tabla):
    precios = resumen_precios(tabla)
    filas = []
    for tipo, g in tabla.groupby("tipo_existencia"):
        descripciones = sorted(set(g["descripcion_n1"]))
        for i, a in enumerate(descripciones):
            for b in descripciones[i + 1:]:
                misma = clave(a) == clave(b)
                escritura = misma or parecidas(a, b)
                candidato = (numeros(a) == numeros(b) and palabras(a)[:1] == palabras(b)[:1]
                             and SequenceMatcher(None, clave(a), clave(b)).ratio() >= PARECIDO_CANDIDATO)
                if not (escritura or candidato):
                    continue
                pa, pb = precios[(tipo, a)], precios[(tipo, b)]
                veces, motivo_precio = comparar_precio(pa, pb)
                final = min([a, b], key=lambda d: (-precios[(tipo, d)]["compras"], len(d)))
                if solo_talla(a, b):
                    sugerencia = "TALLA"
                elif veces and veces > PRECIO_DISTINTO:
                    sugerencia = "NO UNIR" if base_precio(pa, pb) == "compra vs compra" else "REVISAR"
                elif not escritura:
                    sugerencia = "OTRO PRODUCTO"
                elif veces:
                    sugerencia = "UNIR"
                else:
                    sugerencia = "REVISAR"
                lado = lambda s, p: {
                    f"compras_{s}": p["compras"], f"cotizaciones_{s}": p["cotizaciones"],
                    f"unidad_{s}": p["unidad"], f"moneda_{s}": p["moneda"],
                    f"precio_compra_{s}": p["precio_compra"], f"minimo_{s}": p["minimo"], f"maximo_{s}": p["maximo"],
                    f"precio_cotizacion_{s}": p["precio_cotizacion"], f"proveedores_{s}": p["proveedores"]}
                filas.append({
                    "tipo": tipo,
                    "descripcion_A": a, **lado("A", pa),
                    "descripcion_B": b, **lado("B", pb),
                    "precio_veces": veces if veces else motivo_precio,
                    "base_precio": base_precio(pa, pb) if veces else None,
                    "que_cambia": que_cambia(a, b),
                    "nombre_final": final if sugerencia in ("UNIR", "REVISAR") else None,
                    "sugerencia": sugerencia, "decision": None,
                })
    orden = {"UNIR": 0, "REVISAR": 1, "TALLA": 2, "NO UNIR": 3, "OTRO PRODUCTO": 4}
    return pd.DataFrame(filas).sort_values(["sugerencia", "tipo", "descripcion_A"],
                                            key=lambda s: s.map(orden) if s.name == "sugerencia" else s)


def guardar(tabla):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Comparación"
    columnas = ["n"] + list(tabla.columns)
    ws.append(columnas)
    for n, fila in enumerate(tabla.itertuples(index=False), 1):
        ws.append([n] + [None if (not isinstance(v, str) and pd.isna(v)) else v for v in fila])
    colores = {"UNIR": "E2EFDA", "REVISAR": "FFF2CC", "TALLA": "DDEBF7", "NO UNIR": "F8CBAD", "OTRO PRODUCTO": "EDEDED"}
    col_sug = columnas.index("sugerencia")
    for fila in ws.iter_rows():
        for celda in fila:
            celda.font = Font(name="Arial", size=10)
            celda.alignment = Alignment(vertical="top", wrap_text=True)
        if fila[0].row > 1:
            fila[col_sug].fill = PatternFill("solid", fgColor=colores[fila[col_sug].value])
    for celda in ws[1]:
        celda.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")
    anchos = {"descripcion_A": 42, "descripcion_B": 42, "nombre_final": 42, "que_cambia": 34, "n": 5,
              "tipo": 11, "precio_veces": 12, "base_precio": 14, "sugerencia": 11, "decision": 14,
              "proveedores_A": 30, "proveedores_B": 30}
    for i, campo in enumerate(columnas, 1):
        letra = get_column_letter(i)
        ws.column_dimensions[letra].width = anchos.get(campo, 9)
        if campo.split("_")[0] in ("precio", "minimo", "maximo") and campo != "precio_veces":
            for celda in ws[letra][1:]:
                celda.number_format = "#,##0.00"
    # Lado A y lado B con fondo distinto en la cabecera para leerlos rápido
    for celda in ws[1]:
        if celda.value and celda.value.endswith("_A"):
            celda.fill = PatternFill("solid", fgColor="2E75B6")
        elif celda.value and celda.value.endswith("_B"):
            celda.fill = PatternFill("solid", fgColor="548235")
    for celda in ws[get_column_letter(columnas.index("decision") + 1)][1:]:
        celda.fill = PatternFill("solid", fgColor="DDEBF7")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    leeme = wb.create_sheet("Léeme", 0)
    for linea in [
        "Comparación de descripciones parecidas",
        "Cada fila compara dos descripciones del mismo tipo de existencia (lado A en azul, lado B en verde).",
        "Precios unitarios con IGV (total / cantidad), en su moneda y por la unidad con que se compró.",
        "precio_compra: mediana de las compras reales; minimo y maximo de esas compras; proveedores: quiénes las vendieron.",
        "precio_cotizacion: mediana de las cotizaciones (filas sin proveedor ni fecha en la hoja).",
        f"precio_veces: cuántas veces más caro es uno que el otro. base_precio dice si se compararon dos compras reales.",
        "Si la moneda o la unidad son distintas, el precio no se compara y se indica el motivo.",
        "UNIR: solo cambia la escritura y el precio es parecido.",
        "REVISAR: otra unidad o moneda, o precio muy distinto cuando una de las dos es solo cotización.",
        "TALLA: solo cambia la talla; se resuelve al separar la talla a su columna, no aquí.",
        f"NO UNIR: el precio de compras reales difiere más de {PRECIO_DISTINTO} veces.",
        "OTRO PRODUCTO: cambia una palabra que define el producto (tipo, color, forma, modelo).",
        "nombre_final: una de las dos descripciones, la de más compras.",
        "decision: escribe SI para unir, NO para dejarlas separadas, o el nombre que prefieras de entre las dos.",
    ]:
        leeme.append([linea])
    leeme.column_dimensions["A"].width = 115
    wb.save(SALIDA)


if __name__ == "__main__":
    tabla = pares(pd.read_excel(ENTRADA))
    guardar(tabla)
    print(tabla["sugerencia"].value_counts().to_string())
    print(f"Guardado {SALIDA} ({len(tabla)} pares)")
