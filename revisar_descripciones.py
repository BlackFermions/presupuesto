"""
Paso 3a: propone la descripción del item, la marca, el modelo y la talla de cada descripción
distinta del consolidado, para revisarla antes de crear items y productos.

- Las descripciones que ya se trabajaron en el maestro v3 (aprobado por el usuario) reutilizan esa
  separación: descripción del item, marca, modelo y talla.
- Las nuevas usan QUITAR: la descripción solo pierde texto; nunca se agrega ni se corrige nada.
- NO_ES_ITEM: filas que no describen un producto (cajas chicas, intereses, compras generales).
  Siguen en el consolidado; solo no generan item.

Uso:  python revisar_descripciones.py
"""
import re
import unicodedata
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

CONSOLIDADO = Path("Datos consolidados de compras.xlsx")
MAESTRO_V3 = Path("Maestro_items_v3.xlsx")
SALIDA = Path("Revision de descripciones.xlsx")

# Descripción original -> {"marca": texto, "modelo": texto, "talla": texto}
# Cada valor debe aparecer literal en la descripción; se quita de ella y pasa a su columna.
QUITAR = {
    # EPP
    "DOBLE LÍNEA DE VIDA HAUK": {"marca": "HAUK"},
    "GUANTES DE NITRILO TALLA L": {"talla": "TALLA L"},
    "GUANTES DE NITRILO TALLA M": {"talla": "TALLA M"},
    "TRAJE TYVEK TECSEG CON CINTA REFLECTIVA M Y L": {"marca": "TECSEG", "talla": "M Y L"},
    "TRAJE TYVEK TECSEG CON CINTA REFLECTIVA XL": {"marca": "TECSEG", "talla": "XL"},
    "TYVEK TALLA L": {"talla": "TALLA L"},
    "TYVEK TALLA M": {"talla": "TALLA M"},
    "ZAPATO FORTEMAX ( Mecánico y Dieléctrico)": {"marca": "FORTEMAX"},
    # EQUIPO
    "ANDAMIOS LAHYER CERTIFICADOS": {"marca": "LAHYER"},
    "Equipo Graco Neumático 70:1 Incluye manguera de 3/8”x 15 metros, chicote de 1/4” x 1.5 metros, "
    "pistola de pintar Graco XTR7 y con boquilla.": {"marca": "Graco"},
    # MATERIAL
    "BROCA  HSS 1/4\" Truper": {"marca": "Truper"},
    "BROCA HSS 1/8\" Truper": {"marca": "Truper"},
    "BROCA  HSS 3/8\" Truper": {"marca": "Truper"},
    "Brocha de 3\" TUMI": {"marca": "TUMI"},
    "BROCHA DE PINTOR TUMI DE 2\"": {"marca": "TUMI"},
    "BROCHA DE PINTOR TUMI DE 4\"": {"marca": "TUMI"},
    "CINTA AISLANTE ELECTRICA DE 19 MM X 18 M NEGRO TRUPER\"12500": {"marca": "TRUPER", "modelo": "\"12500"},
    "CIZALLA 24\" TRAMONTINA": {"marca": "TRAMONTINA"},
    "DISCOS PARA CORTE 4 1/2 X 1/8'' X 7/8'' DEWALT": {"marca": "DEWALT"},
    "DISCOS PARA CORTE 7 X 1/8'' X 7/8'' DEWALT": {"marca": "DEWALT"},
    "Epoxi novalaca resina fenolica con resina de curado amina cicloalinfática ergon armor": {"marca": "ergon armor"},
    "EXTENSION TELESCÓPICA 3.0 TRUPER (MODELO EXTM-30) CÓDIGO 10549": {"marca": "TRUPER",
                                                                      "modelo": "(MODELO EXTM-30) CÓDIGO 10549"},
    "GRAPAS CROSBY PARA CABLE DE 1/2\"": {"marca": "CROSBY"},
    "Grout epoxico polymerico (sikadur 42 CL juego A+B+C x 27kg)": {"modelo": "sikadur 42 CL"},
    "Pintura hardtop flexi RAL 1001": {"modelo": "hardtop flexi"},
    "Pintura SIGMADUR 550H RAL 1023 (AMARILLO)": {"modelo": "SIGMADUR 550H"},
    "Pintura SIGMADUR 550H RAL 6010 (VERDE)": {"modelo": "SIGMADUR 550H"},
    "PLANCHA TRIPLAY FENOLICO  B/Cp 18mm 1.22x2.44m Sumec": {"marca": "Sumec"},
    "Rodillos de 4\" TORO": {"marca": "TORO"},
    "THINNER SIGMADUR 60-06": {"modelo": "SIGMADUR"},
}

NO_ES_ITEM = {
    "COMPRA EPP", "COMPRA EPP PENDIENTE", "EPP POR CONFIRMAR", "COMPRA MATERIALES 24/08", "CONSUMIBLES",
    "MATERIALES", "INTERES", "INTERESES", "CAJA CHICA 05/08", "CAJA CHICA 08/06", "CAJA CHICA 11/08",
}

# Dudas que solo el usuario puede resolver
NOTAS = {
    "CURSO PBIP": "Es una capacitación; la hoja la marca MATERIAL. ¿Va como SERVICIO?",
    "TRASLADO MATERIALES FIERRO & ACERO CENTER": "Es un traslado; la hoja la marca MATERIAL. ¿Va como MOVILIDAD?",
    "TABLERO ELÉCTRICO": "Único con tipo FABRICACIÓN. ¿Es un tipo nuevo?",
    "SIRENA ELECTROMECANICA SUPERCELEREST 110VACDC 115DB": "¿SUPERCELEREST es marca o modelo?",
    "Pintura poliuretano color gris Ral 7040 (Sigmadur 550 o Hartop Flexi)": "Nombra dos productos posibles; se deja igual.",
    "SIGMADUR 550H(Z) RAL 7045": "Sin el nombre del producto quedaría solo 'RAL 7045'; se deja igual.",
    "SIGMADUR 550H(Z) RAL 9005": "Sin el nombre del producto quedaría solo 'RAL 9005'; se deja igual.",
    "SIKADUR 31 COMPONENTE A Y B X 5 KG": "Nombre comercial que da sentido a la descripción; se deja igual.",
}


def normalizar(texto):
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto).strip().upper()


def quitar(descripcion, fragmentos):
    """Quita cada fragmento literal; si alguno no está, se detiene para no adivinar."""
    texto = re.sub(r"\s+", " ", str(descripcion)).strip()
    fragmentos = list(fragmentos)
    if not fragmentos:
        return texto
    for fragmento in fragmentos:
        patron = rf"(?<![A-Za-z0-9]){re.escape(fragmento)}(?![A-Za-z0-9])"
        texto, cambios = re.subn(patron, " ", texto, count=1, flags=re.IGNORECASE)
        if not cambios:
            raise ValueError(f"'{fragmento}' no está en '{descripcion}'")
    texto = re.sub(r"\(\s+", "(", re.sub(r"\s+([,\.])", r"\1", re.sub(r"\s+", " ", texto)))
    return texto.strip(" -/:,.")


def comprobar_sin_inventar(original, generica):
    permitidas = set(normalizar(original).split())
    nuevas = {p.strip("-/:,.()\"") for p in normalizar(generica).split()} - {""}
    sobran = [p for p in nuevas if not any(p in q for q in permitidas)]
    if sobran:
        raise ValueError(f"'{generica}' tiene texto que no está en '{original}': {sobran}")


def propuesta_v3():
    """Descripción original normalizada -> separación aprobada en el maestro v3."""
    compras = pd.read_excel(MAESTRO_V3, sheet_name="Compras")
    items = pd.read_excel(MAESTRO_V3, sheet_name="Items").set_index("Código")
    compras["k"] = compras["Descripción original"].map(normalizar)
    v3 = {}
    for k, g in compras.groupby("k"):
        fila = g.iloc[0]
        v3[k] = {"descripcion_item": items.at[fila["Código"], "Descripción"],
                 "marca": fila["Marca"], "modelo": fila["Modelo"], "talla": fila["Talla"],
                 "codigo_v3": fila["Código"]}
    return v3


def construir():
    consolidado = pd.read_excel(CONSOLIDADO)
    consolidado["k"] = consolidado["descripcion"].map(normalizar)
    v3 = propuesta_v3()
    pendientes = set(QUITAR)
    filas = []
    for (tipo, k), g in consolidado.groupby(["tipo_existencia", "k"], sort=True):
        original = g["descripcion"].iloc[0]
        fila = {"tipo_existencia": tipo, "descripcion_original": original, "filas": len(g),
                "origenes": ", ".join(sorted({f"{h} {n}" for h, n in zip(g.origen_hoja, g.origen_fila)})[:4]),
                "descripcion_item": None, "marca": None, "modelo": None, "talla": None,
                "propuesta": None, "codigo_v3": None, "nota": NOTAS.get(original)}
        if original in NO_ES_ITEM:
            fila["propuesta"] = "NO ES ITEM"
        elif k in v3:
            fila.update({c: (None if pd.isna(v) else v) for c, v in v3[k].items()})
            fila["propuesta"] = "V3 APROBADO"
        else:
            separar = QUITAR.get(original, {})
            pendientes.discard(original)
            fila["descripcion_item"] = quitar(original, separar.values())
            comprobar_sin_inventar(original, fila["descripcion_item"])
            fila.update(separar)
            fila["propuesta"] = "NUEVA" if separar else "NUEVA SIN CAMBIOS"
        filas.append(fila)
    if pendientes:
        raise ValueError(f"QUITAR tiene descripciones que no están en el consolidado: {pendientes}")
    tabla = pd.DataFrame(filas)
    # Cuántas descripciones caen en el mismo item (mismo tipo y misma descripción genérica)
    clave = tabla["tipo_existencia"] + "|" + tabla["descripcion_item"].fillna("").map(normalizar)
    tabla["descripciones_en_el_item"] = clave.map(clave.value_counts()).where(tabla["descripcion_item"].notna())
    return tabla


def guardar(tabla):
    columnas = ["tipo_existencia", "descripcion_original", "filas", "origenes", "propuesta", "descripcion_item",
                "marca", "modelo", "talla", "descripciones_en_el_item", "codigo_v3", "nota", "decision"]
    tabla["decision"] = None
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Descripciones"
    ws.append(columnas)
    for _, f in tabla[columnas].iterrows():
        ws.append([None if pd.isna(v) else v for v in f])
    anchos = {"descripcion_original": 60, "descripcion_item": 50, "origenes": 26, "nota": 50, "marca": 16,
              "modelo": 18, "propuesta": 18, "tipo_existencia": 14, "decision": 30}
    colores = {"NUEVA": "FFF2CC", "NUEVA SIN CAMBIOS": "FFF9E6", "NO ES ITEM": "F8CBAD"}
    col_prop = columnas.index("propuesta") + 1
    for fila in ws.iter_rows():
        for celda in fila:
            celda.font = Font(name="Arial", size=10)
        color = colores.get(fila[col_prop - 1].value)
        if color and fila[0].row > 1:
            for celda in fila:
                celda.fill = PatternFill("solid", fgColor=color)
    for celda in ws[1]:
        celda.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")
        celda.alignment = Alignment(vertical="center")
    for i, campo in enumerate(columnas, 1):
        ws.column_dimensions[get_column_letter(i)].width = anchos.get(campo, 11)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(SALIDA)


if __name__ == "__main__":
    tabla = construir()
    guardar(tabla)
    print(tabla["propuesta"].value_counts().to_string())
    print(f"Guardado {SALIDA} ({len(tabla)} descripciones)")
