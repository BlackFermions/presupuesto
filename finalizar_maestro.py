import os
import re
import sys

import pandas as pd
from openpyxl import Workbook, load_workbook

import construir_maestro as cm

# Parte del borrador ya revisado por el usuario: sus respuestas están aquí y no se sobrescriben
ENTRADA = "Maestro_borrador.xlsx"
SALIDA = "Maestro_items_v3.xlsx"
PREGUNTAS = "Preguntas_logistica.xlsx"

# Descripciones generales que no identifican un producto (decisión del usuario)
ELIMINAR = ["MAT-0156", "EPP-0052", "MAT-0093", "MAT-0102", "EPP-0048", "EPP-0049"]

# Unidad elegida por el usuario. Solo cambia la etiqueta: el precio no se convierte,
# porque las filas UND cuestan lo mismo que el PAQ o el PAR (ver Preguntas_logistica)
UNIDAD_ELEGIDA = {"MAT-0012": "CAJA", "MAT-0039": "UND", "EPP-0029": "UND", "EPP-0058": "UND"}

# Código del borrador -> texto a quitar de la descripción original (marca, modelo o talla).
# La descripción solo puede perder texto: nunca se agrega ni se corrige nada.
# Lo que queda igual en dos códigos se une en un solo item
QUITAR = {
    "MAT-0008": ["MOBIL"],
    "MAT-0032": ["RUZC3M", "SHNEIDER"],
    "MAT-0037": ["SAI-C"],
    "MAT-0042": ["(TUMI)"], "MAT-0043": ["(TUMI)"], "MAT-0044": ["(TUMI)"],
    "MAT-0045": ["ABC"],
    "MAT-0047": ["AFIRENAS"],
    "MAT-0048": ["SIEMON", "9A6L4"], "MAT-0049": ["SIEMON", "9A6L4"],
    "MAT-0050": ["MIGUELEZ"],
    "MAT-0051": ["AFIRENAS-X RZ1-K 1KV MIGUELEZ"],
    "MAT-0065": ["WERKEN"],
    "MAT-0078": ["CADWELD"],
    "MAT-0079": ["Thor cem"],
    "MAT-0080": ["CADWELD"],
    "MAT-0082": ["3M - 1700 SCOTCH"],
    "MAT-0083": ["TRUPER 12503"],
    "MAT-0086": ["BROTHER PT-E300VP/PT-E550W", "- TZe251"],
    "MAT-0092": ["panduit t70ciw2"],
    "MAT-0094": ["Unicon Fácil"],
    "MAT-0107": ["CHEMA DESMOLD TB"],
    "MAT-0109": ["(DEWALT)"], "MAT-0111": ["DEWALT"], "MAT-0115": ["EN MARCA DeWALT"],
    "MAT-0133": ["MODELO SOLVEX 37-165", "TALLA M"],
    "MAT-0134": ["JOTUN"],
    "MAT-0139": ["BAHCO"],
    "MAT-0144": ["SCNEIDER ELECTRIC"],
    "MAT-0145": ["SIEMON", "Z.-MAX"],
    "MAT-0157": ["WELDOR"], "MAT-0158": ["WELDOR"],
    "MAT-0164": ["SIEMON"],
    "MAT-0166": ["SIEMON", "TERA Z-MAX"],
    "MAT-0168": ["OATEY"],
    "MAT-0183": ["Sigmashield 880 marca PPG"],
    "MAT-0184": ["Sigmadur 550"], "MAT-0185": ["Sigmadur 550"],
    "MAT-0186": ["ABRO"],
    "MAT-0194": ["/ MARCA SCHNEIDER"],
    "MAT-0195": ["CHEMA"],
    "MAT-0202": ["SIEMON"],
    "MAT-0204": ["Baseus"],
    "MAT-0207": ["FS-ONE MAX", "FO #2101535"], "MAT-0208": ["FS-ONE MAX", "FO HILTI"],
    "MAT-0209": ["SAM-LED MODELO: 221514010"],
    "MAT-0211": ["SIEMON 10GMX-HFP02-02B"],
    "MAT-0212": ["SIEMON CT4-BOX-02"],
    "MAT-0213": ["SIEMON HCM-A-2U"],
    "MAT-0214": ["SIEMON Z-TOOL"],
    "MAT-0215": ["SIEMON Z-TOOL-ZP"],
    "MAT-0216": ["SIEMON Z6A-S06B", "Z-MAX"],
    "MAT-0217": ["SIEMON Z6AS-PNL-24K Z-MAX"],
    "MAT-0218": ["SIEMON ZM6A-S03-06B", "Z-MAX"],
    "MAT-0239": ["JOTUN"],
    "MAT-0240": ["Epóxico"], "MAT-0242": ["poliuretano"],
    "MAT-0265": ["Werken"],
    "MAT-0272": ["MARTELL WORSAL"],

    "EPP-0001": ["2097", "3M"],
    "EPP-0002": ["MODELO: TURBINE MARCA: STEELPRO"],
    "EPP-0003": ["HAUK"],
    "EPP-0008": ["CLUTE"],
    "EPP-0009": ["// MARCA - WELLPRO"],
    "EPP-0013": ["TALLA: 42 Y 40"],
    "EPP-0014": ["HERMES 2 FLOTER", "T: 37 - SANDDER"],
    "EPP-0015": ["HERMES 2 FLOTER", "T: 41 - SANDDER"],
    "EPP-0016": ["HERMES 2 FLOTER", "T: 42 - SANDDER"],
    "EPP-0017": ["FORTEMAX NAZCA TALLA 35"], "EPP-0018": ["FORTEMAX NAZCA TALLA 36"], "EPP-0019": ["FORTEMAX NAZCA TALLA 37"],
    "EPP-0020": ["FORTEMAX NAZCA TALLA 38"], "EPP-0021": ["FORTEMAX NAZCA TALLA 39"], "EPP-0022": ["FORTEMAX NAZCA TALLA 44"],
    "EPP-0023": ["NORSEG TALLA 37"], "EPP-0024": ["NORSEG TALLA 41"], "EPP-0025": ["NORSEG TALLA 42"],
    "EPP-0026": ["T/39 (1 UND), T/41 (1 UND)"],
    "EPP-0029": ["MODELO: 60923"],
    "EPP-0030": ["60926", "COD - MARCA // 3M"],
    "EPP-0033": ["FORTE"],
    "EPP-0034": ["SECUREFIT", "- 3M"],
    "EPP-0036": ["- STEELPRO"],
    "EPP-0038": ["TALLA L"], "EPP-0039": ["TALLA M"], "EPP-0040": ["TALLA S"], "EPP-0041": ["TALLA XXL"],
    "EPP-0042": [": 1M, 1L"],
    "EPP-0043": [": 1M, 1L, 1XL"],
    "EPP-0045": ["3M TEMFLEX 165"],
    "EPP-0055": ["60923 3M"],
    "EPP-0056": ["60923"],
    "EPP-0057": ["60923 3M"],
    "EPP-0058": ["MODELO: 2097"],
    "EPP-0059": ["7093B /", "3M"],
    "EPP-0060": ["2097 3M"],
    "EPP-0061": ["MARCA: PORTWERST 3M / 3L"],
    "EPP-0062": ["PORTWEST 6/M 6/L"],
    "EPP-0063": ["TALLA M"],
    "EPP-0064": ["MODELO: A622 TALLA: M"],
    "EPP-0065": ["MARCA: CLUTE MODELO: CUT5 TALLA M"],
    "EPP-0066": ["36/M 12/L"],
    "EPP-0067": ["PORWEST TM /L"],
    "EPP-0069": ["MODELO: 92-600 TALLA: M"],
    "EPP-0070": ["MODELO: SOLVEX 37-165 MARCA: ANSELL TALLA: M"],
    "EPP-0071": ['TALLA "M"'],
    "EPP-0072": ["TOUCH N TUFF (92-600) T: L // MARCA - ANSELL"],
    "EPP-0073": ["HYFLEX 11-840 TALLA 8"],
    "EPP-0074": ["HYFLEX 11840 TALLA 8"],
    "EPP-0075": ["HYFLEX", "11-840 TALLA : 8 // MARCA - ANSELL"],
    "EPP-0076": ["MULTIFLEX", "T: L // MARCA - STEELPRO"],
    "EPP-0077": ["MULTIFLEX", "T: M // MARCA - STEELPRO"],
    "EPP-0078": ["MULTIFLEX", "TALLA M STEELPRO"],
    "EPP-0079": ["MULTIFLEX", "T: M - STEELPRO"],
    "EPP-0080": ["TALLA M"],
    "EPP-0081": ["N°9"],
    "EPP-0084": ["SEGPRO A622 M-L"],
    "EPP-0086": ["TOUCHNTUFF"],
    "EPP-0088": ["TALLA 32"],
    "EPP-0091": ["TURBINE STEEL PRO"],
    "EPP-0092": ["TURBINE", "// MARCA - STEELPRO"],
    "EPP-0093": ["TURBINE STEELPRO"],
    "EPP-0094": ["TURBINE", "- STEELPRO"],
    "EPP-0095": ["STEELPRO"],
    "EPP-0096": ["TURBINE"],
    "EPP-0097": ["TURBINE", "STEELPRO"],
    "EPP-0098": ["TURBINE"],
    "EPP-0100": ["MARCA K&N AMERICAN HOIST"],
    "EPP-0108": ["3M 98"],
    "EPP-0109": ["98"],
    "EPP-0111": ["(TALLA SEGÚN PERSONAL)"],
    "EPP-0112": ["T30"], "EPP-0113": ["TALLA 28"], "EPP-0114": ["TALLA 30"], "EPP-0115": ["TALLA 40"],
    "EPP-0117": ["(TALLAS COMPLETAS )"],
    "EPP-0118": ["XXL"],
    "EPP-0119": ["M Y L"],
    "EPP-0120": ["XL"],
    "EPP-0121": [": 2S, 6M , 4L"],
    "EPP-0122": ["TALLA L"], "EPP-0123": ["TALLA M"],
    "EPP-0124": ["3M"],
    "EPP-0125": ["6100", "SMALL // MARCA - 3M"],
    "EPP-0126": ["6200", "MEDIUM // MARCA - 3M"],
    "EPP-0127": ["7502", "MEDIUM - 3M"],
    "EPP-0128": ["MARCA: 3M MODELO: 6200"],
    "EPP-0129": ["MODELO: 6200"],
    "EPP-0130": ["3M"],
    "EPP-0131": ["7502", "60923 3M"],
    "EPP-0133": ["XXL"],
    "EPP-0134": ["SAFE 14308 TRUPER"],
    "EPP-0137": ["STEELPRO"],
    "EPP-0138": ["EP-T06C", "- STEELPRO"],
    "EPP-0139": ["1100 3M"],
    "EPP-0140": ["TSG15 BC. T: L - TECSEG"], "EPP-0141": ["TSG15 BC. T: M - TECSEG"], "EPP-0142": ["TSG15 BC. T: XL - TECSEG"],
    "EPP-0143": ["TALLA: L"], "EPP-0144": ["TALLA: M"],
    "EPP-0145": ["TYVEK MARCA: TECSEG MOD: TSG15BC TALLA: L (1) - M (2)"],
    "EPP-0146": ["TYVEK", "25M Y 25L"],
    "EPP-0147": ["TYVEX 50 (M) -50 (L)"],
    "EPP-0148": ["TYVECK", "TECSEG"],
    "EPP-0149": ["25M / 25L"],
    "EPP-0150": ["- STEELPRO"],
    "EPP-0153": ["FORTEMAX"],
    "EPP-0157": ["TALLA 35 Y 40"],

    "HER-0003": ["PLANO"], "HER-0004": ["PUNTA"],
    "HER-0012": ["Marco:BAHCO"],
    "HER-0013": ["BAHCO"],
    "HER-0014": ["bosch"],
    "EQU-0001": ["ESAB"],
    "EQU-0002": ["ESAB ROGUE ES 305i"],
    "EQU-0003": ["RANKOR RAB60001"],
    "EQU-0004": ["ABLE"],
    "SER-0001": ["GWS 18V-10 PC"],
    "SER-0002": ["9557HP"],
    "SER-0003": ["GWS 9-125 P"],
    "SER-0010": ["BOSCH"],
}

# Pares marcados con "si" en 'Posibles duplicados' que aún difieren en su texto: se usa el del segundo código
UNIONES = {
    "MAT-0253": "MAT-0249",
    "MAT-0069": "MAT-0068",
    "MAT-0053": "MAT-0057",
    "MAT-0054": "MAT-0059",
    "MAT-0225": "MAT-0226",
    "EPP-0083": "EPP-0068",
    "EPP-0010": "EPP-0007",
    "EPP-0067": "EPP-0062",
    "EPP-0080": "EPP-0071",
    "EPP-0103": "EPP-0105",
    "EPP-0055": "EPP-0057",
}

# Calzado y guantes vienen por par. Un producto vendido por CAJA no se toca: su precio es por caja
ITEMS_POR_PAR = re.compile(r"^(BOTIN|BOTA|ZAPATO|GUANTE)(?!.*CAJA)")

# Tallas que el texto trae en cantidades ("6/M 6/L", "25M Y 25L") y el detector no puede leer solo
TALLAS_FIJAS = {
    "EPP-0013": "40, 42", "EPP-0026": "39, 41", "EPP-0042": "M, L", "EPP-0043": "M, L, XL", "EPP-0061": "M, L",
    "EPP-0062": "M, L", "EPP-0066": "M, L", "EPP-0067": "M, L", "EPP-0073": "8", "EPP-0074": "8", "EPP-0075": "8",
    "EPP-0081": "9", "EPP-0084": "M, L", "EPP-0117": "", "EPP-0119": "M, L", "EPP-0121": "S, M, L", "EPP-0125": "SMALL",
    "EPP-0126": "MEDIUM", "EPP-0127": "MEDIUM", "EPP-0145": "L, M", "EPP-0146": "M, L", "EPP-0147": "M, L", "EPP-0149": "M, L",
    "EPP-0157": "35, 40",
}

MARCAS_ADICIONALES = ["SANDDER", "FORTEMAX", "SEGPRO", "PORWEST", "PORTWERST", "STEEL PRO", "WERKEN", "OATEY", "PANDUIT",
                      "BASEUS", "MARTELL", "ABC", "UNICON", "SHNEIDER", "SCNEIDER", "WELLPRO", "AFIRENAS"]
MODELOS_CONOCIDOS = ["TURBINE", "HERMES 2 FLOTER", "NAZCA", "HYFLEX 11-840", "HYFLEX 11840", "SOLVEX 37-165", "SIGMASHIELD 880",
                     "SIGMADUR 550", "TEMFLEX 165", "SECUREFIT", "ROGUE ES 305I", "SIKAGROUT 212", "SIKADUR 32", "SIKAFLEX 11 FC",
                     "SIKABOOM", "SIKACEM CURADOR", "SIKA REP 500", "FS-ONE MAX", "HARDTOP FLEXI", "MULTIFLEX CUT-5", "MULTIFLEX CUT-D",
                     "TOUCH N TUFF", "TOUCHNTUFF", "TSG15 BC", "Z-MAX", "Z.-MAX"]

# (tema, códigos del borrador, pregunta). Los datos de cada pregunta salen de las compras.
PREGUNTAS_INICIALES = [
    ("Precio", ["MAT-0085"], "¿Por qué varía más de 50% el precio de la cinta masking?"),
    ("Precio", ["MAT-0094"], "¿Por qué varía más de 50% el precio del concreto embolsado?"),
    ("Precio", ["MAT-0116", "MAT-0117"], 'El disco de desbaste de 4" y el de 7" tienen el mismo precio: ¿es correcto?'),
    ("Unidad", ["MAT-0132"], "La grapa para zuncho tiene unidad GL: ¿es galón, global u otra?"),
    ("Unidad", ["MAT-0039"], "La fila UND cuesta igual que el PAQ: ¿ese UND es en realidad un paquete de 100?"),
    ("Unidad", ["EPP-0029"], "El cartucho por PAR y por UND tiene casi el mismo precio: ¿ese UND es en realidad un par?"),
    ("Unidad", ["EPP-0058"], "El filtro 2097 por PAR y por UND tiene casi el mismo precio: ¿ese UND es en realidad un par?"),
    ("Unidad", ["MAT-0110", "MAT-0111"], 'El disco de corte de 7" se compró por CAJA y por UND: ¿cuántas unidades trae la caja?'),
    ("Item", ["MAT-0009"], "¿Qué es 'ACELERACIÓN'? ¿Es un producto o un servicio?"),
    ("Item", ["MAT-0234"], "¿Qué es 'SPORADE'? No se entiende el producto."),
    ("Item", ["MAT-0240", "MAT-0242"], "Se unieron thinner epóxico y poliuretano: ¿se pueden tratar como un solo item?"),
    ("Item", ["HER-0003", "HER-0004"], "Se unieron cincel plano y cincel punta: ¿se pueden tratar como un solo item?"),
    ("Tipo", ["MAT-0038", "EPP-0011"], "La bolsa de basura 140 L negra está como MATERIAL y como EPP: ¿cuál corresponde?"),
    ("Tipo", ["MAT-0039", "EPP-0012"], "La bolsa de basura 140 L roja está como MATERIAL y como EPP: ¿cuál corresponde?"),
    ("Tipo", ["MAT-0133", "EPP-0070"], "El guante de nitrilo 37-165 está como MATERIAL y como EPP: ¿cuál corresponde?"),
    ("Tipo", ["MAT-0083", "EPP-0045"], "La cinta aislante amarilla 19 mm x 18 m está como MATERIAL y como EPP: ¿cuál corresponde?"),
    ("Tipo", ["MAT-0120", "EPP-0051"], 'La driza de 1/4" está como MATERIAL y como EPP: ¿cuál corresponde?'),
    ("Proveedor", [], "RUC: 20612251569 aparece en PROMICON, WELLPRO, SODIMAC, S&H y GELMANT; 20100205654 en "
                      "Braillard y Grupo Carlitos; PROMICON tiene 2 RUC. ¿Cuál es el RUC correcto de cada uno?"),
]


def quitar_fragmentos(codigo, descripcion):
    """Solo elimina texto del original; si un fragmento no está, se detiene para no adivinar."""
    texto = re.sub(r"\s+", " ", str(descripcion)).strip()
    for fragmento in QUITAR.get(codigo, []):
        # Palabra completa: "3M" no debe coincidir dentro de "13MM"
        inicio = r"(?<![A-Za-z0-9])" if fragmento[0].isalnum() else ""
        fin = r"(?![A-Za-z0-9])" if fragmento[-1].isalnum() else ""
        patron = inicio + re.escape(fragmento) + fin
        texto, cambios = re.subn(patron, " ", texto, count=1, flags=re.IGNORECASE)
        if not cambios:
            raise ValueError(f"{codigo}: '{fragmento}' no está en '{descripcion}'")
    # Solo limpia los separadores que quedaron sueltos al quitar el texto
    texto = re.sub(r"\(\s+", "(", re.sub(r"\s+,", ",", re.sub(r"\s+", " ", texto)))
    return texto.strip(" -/:,.")


def comprobar_sin_inventar(maestro):
    """Cada palabra de la descripción debe existir en el texto original de ese código o de su par unido."""
    palabras = lambda t: set(re.findall(r"[^\s]+", str(t).upper()))
    for codigo, fila in maestro.iterrows():
        origen = maestro.loc[UNIONES.get(codigo, codigo), "Descripción"]
        nuevas = {p.strip("-/:,.()") for p in palabras(fila["generica"])} - {""}
        permitidas = palabras(origen)
        if not all(any(n in p for p in permitidas) for n in nuevas):
            raise ValueError(f"{codigo}: '{fila['generica']}' tiene texto que no está en '{origen}'")


def marca_de_compra(clave, marcas):
    # "3M X 20CM" y "3MX50M" son medidas, no la marca
    return cm.extraer_marca(re.sub(r"\b3M\s*X", " ", clave), marcas)


def modelo_de_compra(clave):
    if clave.startswith("IMPRESION"):  # FR45 es el nombre de un formato, no un modelo
        return ""
    for modelo in MODELOS_CONOCIDOS:
        if modelo in clave:
            return modelo
    return cm.extraer_modelo(clave)


def unir_valores(serie):
    valores = []
    for valor in serie:
        for parte in str(valor).split(", "):
            if parte and parte not in valores:
                valores.append(parte)
    return ", ".join(sorted(valores))


def preparar(borrador):
    maestro = borrador.parse("Maestro").fillna("").set_index("Código")
    compras = borrador.parse("Compras").fillna("")
    faltantes = (set(QUITAR) | set(UNIONES) | set(UNIONES.values()) | set(ELIMINAR) | set(UNIDAD_ELEGIDA) | set(TALLAS_FIJAS)) - set(maestro.index)
    if faltantes:
        raise ValueError(f"Códigos que no están en el borrador: {sorted(faltantes)}")

    eliminados = compras[compras["Código"].isin(ELIMINAR)].copy()
    compras = compras[~compras["Código"].isin(ELIMINAR)].copy()
    maestro = maestro.drop(ELIMINAR)

    maestro["generica"] = [quitar_fragmentos(c, d) for c, d in zip(maestro.index, maestro["Descripción"])]
    for codigo, destino in UNIONES.items():
        maestro.loc[codigo, "generica"] = maestro.loc[destino, "generica"]
    comprobar_sin_inventar(maestro)
    # El tipo es parte de la clave: la clasificación de logística no se cambia al unir items
    maestro["clave"] = maestro["Tipo de existencia"] + "|" + maestro["generica"].map(cm.normalizar)

    marcas = sorted(set(cm.cargar_marcas()) | set(MARCAS_ADICIONALES), key=len, reverse=True)
    original = compras["Descripción original"].map(cm.normalizar)
    compras["marca"] = original.map(lambda c: marca_de_compra(c, marcas))
    compras["modelo"] = original.map(modelo_de_compra)
    compras["talla"] = [TALLAS_FIJAS.get(c, cm.extraer_talla(o)) for c, o in zip(compras["Código"], original)]
    compras["clave"] = compras["Código"].map(maestro["clave"])
    return maestro, compras, eliminados


def construir_items(maestro, compras):
    filas = []
    for clave, grupo in maestro.groupby("clave", sort=False):
        compras_item = compras[compras["clave"] == clave]
        elegidas = [UNIDAD_ELEGIDA[c] for c in grupo.index if c in UNIDAD_ELEGIDA]
        unidades = sorted(compras_item["Unidad"].unique())
        revisar = ""
        if ITEMS_POR_PAR.match(clave.split("|")[1]) and set(unidades) <= {"UND", "PAR"}:
            elegidas = ["PAR"]
        elif not elegidas and "UND" in unidades:
            elegidas = ["UND"]  # regla del usuario: quedarse con la unidad menor
            if len(unidades) > 1:
                revisar = "Precio por " + ", ".join(u for u in unidades if u != "UND") + ": falta saber cuántas UND trae"
        elif not elegidas and len(unidades) > 1:
            revisar = "Comprado en varias unidades: " + ", ".join(unidades)
        filas.append({
            "clave": clave,
            "tipo": grupo["Tipo de existencia"].iloc[0],
            "descripcion": grupo["generica"].iloc[0],
            "marcas": unir_valores(compras_item["marca"]),
            "modelos": unir_valores(compras_item["modelo"]),
            "tallas": unir_valores(compras_item["talla"]),
            # Compras viene de más reciente a más antigua
            "unidad": elegidas[0] if elegidas else compras_item["Unidad"].iloc[0],
            "revisar": revisar,
        })
    items = pd.DataFrame(filas)
    items["orden"] = items["tipo"].map(list(cm.PREFIJOS).index)
    items = items.sort_values(["orden", "clave"]).reset_index(drop=True)
    items["codigo"] = (items["tipo"].map(cm.PREFIJOS) + "-"
                       + (items.groupby("tipo").cumcount() + 1).astype(str).str.zfill(4))
    return items


def escribir_libro(items, maestro, compras, eliminados):
    codigos = dict(zip(items["clave"], items["codigo"]))
    libro = Workbook()

    hoja = libro.active
    hoja.title = "Léeme"
    hoja["A1"] = "Maestro de items genéricos"
    hoja["A1"].font = cm.FUENTE_TITULO
    lineas = [
        f"Fuente: '{ENTRADA}' con las respuestas del usuario.",
        "Items: la descripción no lleva marca ni talla. Marcas, modelos y tallas son las que se han comprado.",
        "Compras: la marca, modelo y talla reales de cada compra, y su precio sin IGV.",
        "Descripciones unidas: qué descripción original quedó dentro de cada item, para revisar las uniones.",
        "Marca, modelo y talla salen del texto: pueden tener errores o estar vacíos.",
    ]
    for i, linea in enumerate(lineas, 3):
        hoja.cell(i, 1, linea).font = cm.FUENTE
    inicio = 3 + len(lineas) + 1
    for col, titulo in enumerate(["Tipo de existencia", "Items", "Filas de compra"], 1):
        celda = hoja.cell(inicio, col, titulo)
        celda.font, celda.fill = cm.FUENTE_ENCABEZADO, cm.COLOR_ENCABEZADO
    for i, tipo in enumerate(cm.PREFIJOS, inicio + 1):
        hoja.cell(i, 1, tipo).font = cm.FUENTE
        hoja.cell(i, 2, f"=COUNTIF(Items!$B$2:$B${len(items) + 1},A{i})").font = cm.FUENTE
        hoja.cell(i, 3, f"=COUNTIF(Compras!$K$2:$K${len(compras) + 1},A{i})").font = cm.FUENTE
    total = inicio + 1 + len(cm.PREFIJOS)
    hoja.cell(total, 1, "Total").font = cm.FUENTE_NEGRITA
    hoja.cell(total, 2, f"=SUM(B{inicio + 1}:B{total - 1})").font = cm.FUENTE_NEGRITA
    hoja.cell(total, 3, f"=SUM(C{inicio + 1}:C{total - 1})").font = cm.FUENTE_NEGRITA
    hoja.column_dimensions["A"].width = 26
    hoja.column_dimensions["B"].width = 10
    hoja.column_dimensions["C"].width = 16

    hoja = libro.create_sheet("Items")
    filas = [[i.codigo, i.tipo, i.descripcion, i.unidad, i.marcas, i.modelos, i.tallas, i.revisar]
             for i in items.itertuples()]
    cm.escribir_tabla(hoja, ["Código", "Tipo de existencia", "Descripción", "Unidad", "Marcas compradas",
                             "Modelos comprados", "Tallas compradas", "Revisar"],
                      filas, [11, 15, 72, 9, 26, 26, 16, 34])
    for f in range(2, len(filas) + 2):
        if hoja.cell(f, 8).value:
            for col in range(1, 9):
                hoja.cell(f, col).fill = cm.COLOR_REVISAR

    hoja = libro.create_sheet("Compras")
    filas = [[codigos[c["clave"]], pd.Timestamp(c["Fecha"]).to_pydatetime(), c["Descripción original"], c["marca"], c["modelo"],
              c["talla"], c["Cant"], c["Unidad"], c["Moneda"], c["Precio unit. sin IGV"], c["Tipo"], c["Proveedor"],
              c["Factura"], c["Código"]] for _, c in compras.sort_values("Fecha", ascending=False, kind="stable").iterrows()]
    cm.escribir_tabla(hoja, ["Código", "Fecha", "Descripción original", "Marca", "Modelo", "Talla", "Cant", "Unidad comprada",
                             "Moneda", "Precio unit. sin IGV", "Tipo", "Proveedor", "Factura", "Código borrador"],
                      filas, [11, 12, 62, 18, 18, 8, 8, 10, 10, 11, 14, 40, 15, 11])
    for f in range(2, len(filas) + 2):
        hoja.cell(f, 2).number_format = "dd/mm/yyyy"
        hoja.cell(f, 10).number_format = "#,##0.00"

    hoja = libro.create_sheet("Descripciones unidas")
    conteo = compras["Código"].value_counts()
    uniones = maestro.reset_index()
    uniones["nuevo"] = uniones["clave"].map(codigos)
    uniones["items_en_grupo"] = uniones.groupby("clave")["Código"].transform("size")
    uniones = uniones[(uniones["items_en_grupo"] > 1) | uniones["Código"].isin(QUITAR)].sort_values(["nuevo", "Código"])
    filas = [[u.nuevo, u.generica, u.Código, u.Descripción, int(conteo.get(u.Código, 0))] for u in uniones.itertuples()]
    cm.escribir_tabla(hoja, ["Código", "Descripción genérica", "Código borrador", "Descripción original", "Compras"],
                      filas, [11, 58, 11, 74, 8])

    hoja = libro.create_sheet("Eliminados")
    filas = [[e["Código"], pd.Timestamp(e["Fecha"]).to_pydatetime(), e["Descripción original"], e["Moneda"], e["Sub total"],
              e["Proveedor"]] for _, e in eliminados.iterrows()]
    cm.escribir_tabla(hoja, ["Código borrador", "Fecha", "Descripción", "Moneda", "Sub total sin IGV", "Proveedor"],
                      filas, [12, 12, 36, 10, 12, 44])
    for f in range(2, len(filas) + 2):
        hoja.cell(f, 2).number_format = "dd/mm/yyyy"
        hoja.cell(f, 5).number_format = "#,##0.00"

    libro.save(SALIDA)


def actualizar_preguntas(compras_borrador):
    """Historial para reconfirmar con logística: agrega las nuevas y nunca toca lo ya respondido."""
    encabezados = ["ID", "Registrada", "Tema", "Pregunta", "Datos de las compras", "Respuesta de logística", "Fecha respuesta"]
    existentes = set()
    if os.path.exists(PREGUNTAS):
        libro = load_workbook(PREGUNTAS)
        hoja = libro.active
        existentes = {fila[3] for fila in hoja.iter_rows(min_row=2, values_only=True)}
    else:
        libro = Workbook()
        hoja = libro.active
        hoja.title = "Preguntas"
        cm.escribir_tabla(hoja, encabezados, [], [8, 12, 11, 62, 70, 40, 12])

    hoy = pd.Timestamp.now().strftime("%d/%m/%Y")
    for tema, codigos, pregunta in PREGUNTAS_INICIALES:
        if pregunta in existentes:
            continue
        filas = compras_borrador[compras_borrador["Código"].isin(codigos)]
        datos = " | ".join(f"{pd.Timestamp(c['Fecha']):%d/%m/%Y} {c['Cant']} {c['Unidad']} a {c['Precio unit. sin IGV']:.2f} {c['Moneda']}"
                           f" ({c['Tipo']}, {c['Proveedor']})" for _, c in filas.iterrows())
        nueva = hoja.max_row + 1
        valores = [f"P-{nueva - 1:03d}", hoy, tema, pregunta, datos, "", ""]
        for col, valor in enumerate(valores, 1):
            celda = hoja.cell(nueva, col, valor)
            celda.font = cm.FUENTE
            celda.alignment = cm.Alignment(wrap_text=True, vertical="top")
    hoja.auto_filter.ref = f"A1:G{hoja.max_row}"
    libro.save(PREGUNTAS)
    return hoja.max_row - 1


def main():
    if os.path.exists(SALIDA) and "--sobrescribir" not in sys.argv:
        sys.exit(f"{SALIDA} ya existe y puede tener cambios tuyos. Para reemplazarlo usa --sobrescribir")
    borrador = pd.ExcelFile(ENTRADA)
    compras_borrador = borrador.parse("Compras").fillna("")
    maestro, compras, eliminados = preparar(borrador)
    items = construir_items(maestro, compras)
    escribir_libro(items, maestro, compras, eliminados)
    n_preguntas = actualizar_preguntas(compras_borrador)

    print(f"Items: {len(maestro)} del borrador -> {len(items)} genéricos | {items['tipo'].value_counts().to_dict()}")
    print(f"Compras: {len(compras)} | eliminadas: {len(eliminados)} ({eliminados['Sub total'].sum():,.2f} sin IGV)")
    print(f"Por revisar: {(items['revisar'] != '').sum()} | preguntas: {n_preguntas} -> {SALIDA}, {PREGUNTAS}")


if __name__ == "__main__":
    main()
