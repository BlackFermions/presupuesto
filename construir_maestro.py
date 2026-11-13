import os
import re
import sys
import unicodedata
from difflib import SequenceMatcher
from itertools import combinations

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ENTRADA = "Inventario maestro.xlsx"
HOJA = "Hoja2"
CATALOGO = "Catalogo.xlsx"   # de aquí salen marcas conocidas
SALIDA = "Maestro_borrador.xlsx"

# Decisiones acordadas con el usuario: la clasificación de logística no se toca salvo aquí
ELIMINAR = ["INTERESES"]
CAMBIOS_TIPO = {
    "SERVICIO": [
        "CALIBRACION DE PINZA AMPERIMETRICA", "CALIBRACION DE NIVEL TOPOGRAFICO",
        "CALIBRACION TELUROMETRO", "CALIBRACION MEGOHMETRO",
        "AMALODARA GWS 18V-10 PC (MANTENIMIENTO)", "AMOLADORA 9557HP (MANTENIMIENTO)",
        "AMOLADORA GWS 9-125 P (MANTENIMIENTO 4 CARBONES)",
        "REPARACION DE 3 AMOLADORAS BOSCH", "RECARGAR EXTINTORES PQS DE 9 KG", "MEMORIA CALCULO",
    ],
    "MOVILIDAD": ["SERVICIO TRANSPORTE ENVIO MATERIALES", "MOVILIDAD MATERIALES"],
    "EPP": ["ARNES HAUK", "LINEA RETRACTIL DE 3 MTRS DE CABLE MARCA K&N AMERICAN HOIST"],
    "HERRAMIENTA": [
        "TALADRO BOSCH", "WINCHA - CINTA METRICA METALICA 8 M (FLEXOMETRO)",
        "CINCEL PUNTA PARA MARTILLO DEMOLEDOR 27 KG", "CINCEL PLANO PARA MARTILLO DEMOLEDOR 27 KG",
        'SIERRA COPA 3/4" CON GUIA MARCO:BAHCO',
        'SIERRA COPA BIMETALICA BAHCO 51MM" Y EJE SIERRA COPA BIMETALICA 32-210MM 10.10MM',
    ],
    "MATERIAL": ["REPUESTO NIVEL TOPOGRAFICO"],
}

PREFIJOS = {"MATERIAL": "MAT", "EPP": "EPP", "HERRAMIENTA": "HER", "EQUIPO": "EQU",
            "SERVICIO": "SER", "MOVILIDAD": "MOV"}

UNIDADES = {"CAJ": "CAJA", "CJ": "CAJA", "CJA": "CAJA", "ROLL": "ROLLO", "RRLL": "ROLLO",
            "GAL": "GALON", "GLN": "GALON", "JGO": "JUEGO", "JUEG": "JUEGO",
            "MILL": "MILLAR", "MTS": "METRO"}
PROVEEDORES = {
    "PROVEEDORES DE MINAS Y CONSTRUCCION S.A.C. (PROMICON)": "PROMICON",
    "WELLPRO": "WELLPRO SAFETY S.A.C.",
    "CORPELIMA": "CORPELIMA S.A.C.",
    "LU.KLEYD E.I.R.L.": "LU.KLEYD PERU E.I.R.L.",
    "SEGUINSA": "SEGURYTEC E INGENIERIA S.A.C. (SEGUINSA)",
}

# Marcas vistas en las compras que no están en los catálogos web
MARCAS_EXTRA = ["STEELPRO", "TECSEG", "SIEMON", "DEWALT", "BAHCO", "ESAB", "ANSELL", "NORSEG", "FORTE",
                "TUMI", "WELDOR", "HAUK", "JOTUN", "PPG", "MOBIL", "RANKOR", "ABRO", "PORTWEST", "HILTI",
                "MAKITA", "CHEMA", "ABLE", "K&N"]
# Nombres del catálogo web que se confunden con palabras comunes ("HEX", "PILA", "TEE"...)
MARCAS_AMBIGUAS = {"OTRO", "PROMELSA", "E.F.", "DF", "TC", "TEC", "TEE", "HEX", "ENC", "GALA", "PILA", "KAP",
                   "KCL", "KESS", "SEW", "JNG", "JSL", "CAPT", "SIRENA", "TALMA", "COMAR", "COEL", "INTELLI",
                   "PRODUIT", "SUNJ", "PLYON", "ICET"}
UNIDADES_EN_TEXTO = re.compile(r"(MM2?|CM|MTS?|MTRS?|KV|VDC|VAC|V|W|AWG|KG|LBS?|ML|L|LTS|GLN|DB|AMPERIOS|PINES|PIES|MCM|HZ|KA|A)\"?$")
NO_ES_MODELO = re.compile(r"^(IP\d+|NTP-|ASTM|A615|RZ1|RJ-?45|CUT-?\d|KN95|\d+(GR|M|MX.*|MTS.*))")
COLORES = {"AZUL", "ROJO", "VERDE", "AMARILLO", "NEGRO", "NEGRA", "BLANCO", "BLANCA", "GRIS", "NARANJA", "VIOLETA"}

COLOR_ENCABEZADO = PatternFill("solid", fgColor="1F3864")
COLOR_REVISAR = PatternFill("solid", fgColor="FFF2CC")
FUENTE = Font(name="Arial", size=10)
FUENTE_ENCABEZADO = Font(name="Arial", size=10, bold=True, color="FFFFFF")
FUENTE_TITULO = Font(name="Arial", size=14, bold=True)
FUENTE_NEGRITA = Font(name="Arial", size=10, bold=True)


def normalizar(texto):
    """Mayúsculas, sin tildes ni comillas raras: dos escrituras iguales dan la misma clave."""
    texto = re.sub(r"[”“″«»¨]", '"', str(texto))
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode().upper()
    return re.sub(r"\s+", " ", texto).strip()


def cargar_marcas():
    catalogo = pd.read_excel(CATALOGO, usecols=["MARCA"])["MARCA"].dropna().astype(str).str.upper().unique()
    marcas = {m for m in catalogo if m not in MARCAS_AMBIGUAS and (len(m) > 3 or m in ("3M", "ABB"))}
    return sorted(marcas | set(MARCAS_EXTRA), key=len, reverse=True)


def extraer_marca(clave, marcas):
    m = re.search(r"\bMARCA\b\s*[-:]?\s*([A-Z0-9&][A-Z0-9&\-\.]+)", clave)
    if m:
        return m.group(1)
    if re.search(r"\bSIKA", clave):
        return "SIKA"
    # "CINTA 3M" describe la cinta reflectiva de la prenda, no la marca de la prenda
    texto = clave.replace("CINTA 3M", "")
    for marca in marcas:
        if re.search(rf"(?<![A-Z0-9]){re.escape(marca)}(?![A-Z0-9])", texto):
            return marca
    return ""


def extraer_modelo(clave):
    m = re.search(r"MODELO\s*:?\s*([A-Z0-9][A-Z0-9\-\.]+)", clave)
    if m:
        return m.group(1)
    # Colores RAL y normas llevan números pero no identifican el modelo
    texto = re.sub(r"\b(RAL|EPOXICA|ISO)\s+\d+", " ", clave)
    for token in re.findall(r"[A-Z0-9][A-Z0-9\-]{3,}", texto):
        if NO_ES_MODELO.search(token):
            continue
        tiene_letra_y_numero = re.search(r"[A-Z]", token) and re.search(r"\d", token)
        solo_numero_largo = re.fullmatch(r"\d{4,6}", token)
        if (tiene_letra_y_numero and not UNIDADES_EN_TEXTO.search(token)
                and not re.match(r"\d+X\d", token)) or solo_numero_largo:
            return token
    return ""


def extraer_talla(clave):
    m = re.search(r'\b(?:TALLA|T)\s*[:\.]\s*"?(XXXL|XXL|XL|L|M|S|\d{2})\b', clave) \
        or re.search(r'\bTALLA\s*"?(XXXL|XXL|XL|L|M|S|\d{2})\b', clave) \
        or re.search(r"\bT/?(\d{2})\b", clave) \
        or re.search(r"\b(XXXL|XXL|XL)\b", clave)
    return m.group(1) if m else ""


def limpiar_compras():
    df = pd.read_excel(ENTRADA, sheet_name=HOJA)
    df["clave"] = df["DESCRIPCIÓN"].map(normalizar)

    eliminados = df[df["clave"].isin(ELIMINAR)]
    df = df[~df["clave"].isin(ELIMINAR)].copy()

    df["tipo_original"] = df["TIPO EXISTENCIA"]
    df["tipo"] = df["TIPO EXISTENCIA"]
    for tipo, claves in CAMBIOS_TIPO.items():
        faltantes = set(claves) - set(df["clave"])
        if faltantes:
            raise ValueError(f"Descripciones no encontradas para {tipo}: {faltantes}")
        df.loc[df["clave"].isin(claves), "tipo"] = tipo

    df["unidad_original"] = df["U/M"]
    df["unidad"] = df["U/M"].str.strip().replace(UNIDADES)
    df["proveedor_original"] = df["PROVEEDOR"]
    df["proveedor"] = df["PROVEEDOR"].str.strip().replace(PROVEEDORES)
    df["ruc"] = df["RUC"].map(lambda r: str(int(r)) if pd.notna(r) else "")
    df["Fecha"] = pd.to_datetime(df["Fecha"])
    # Más reciente primero: la descripción y unidad del maestro salen de la última compra
    return df.sort_values("Fecha", ascending=False, kind="stable"), eliminados


def notas_de_revision(clave, compras):
    notas = []
    if compras["unidad"].nunique() > 1:
        notas.append("Comprado en varias unidades: " + ", ".join(sorted(compras["unidad"].unique())))
    if compras["MONEDA"].nunique() > 1:
        notas.append("Comprado en soles y dólares: mín/máx no son comparables")
    elif compras["PRECIO UNIT"].max() > compras["PRECIO UNIT"].min() * 1.5:
        notas.append("El precio varía más de 50% entre compras")
    if "POR CONFIRMAR" in clave or clave == "MATERIALES":
        notas.append("No describe un producto concreto")
    if (compras["unidad"] == "GL").any():
        notas.append("Unidad GL dudosa: ¿galón o global?")
    if clave == 'DISCO DE DESBASTE DE 7"':
        notas.append('Mismo precio que el disco de 4": confirmar')
    return "; ".join(notas)


def construir_maestro(compras, marcas):
    filas = []
    for clave, grupo in compras.groupby("clave", sort=False):
        ultima = grupo.iloc[0]
        filas.append({
            "clave": clave,
            "tipo": ultima["tipo"],
            "descripcion": re.sub(r"\s+", " ", str(ultima["DESCRIPCIÓN"])).strip(),
            "marca": extraer_marca(clave, marcas),
            "modelo": extraer_modelo(clave),
            "talla": extraer_talla(clave),
            "unidad": ultima["unidad"],
            "revisar": notas_de_revision(clave, grupo),
        })
    maestro = pd.DataFrame(filas)
    maestro["orden"] = maestro["tipo"].map(list(PREFIJOS).index)
    maestro = maestro.sort_values(["orden", "clave"]).reset_index(drop=True)
    maestro["codigo"] = (maestro["tipo"].map(PREFIJOS) + "-"
                         + (maestro.groupby("tipo").cumcount() + 1).astype(str).str.zfill(4))
    return maestro


def posibles_duplicados(maestro, umbral=0.82):
    """Pares parecidos que no se diferencian en medida, talla, color ni variante. Solo propone; no une nada."""
    numeros = lambda c: set(re.findall(r"\d+(?:[\./]\d+)?", c))
    variante = lambda c: set(re.findall(r'\bTIPO "?([A-Z]{1,2})\b', c))  # caja TIPO LB y TIPO LL son distintas
    pares = []
    for a, b in combinations(maestro.itertuples(), 2):
        if a.tipo != b.tipo:
            continue
        na, nb = numeros(a.clave), numeros(b.clave)
        ca, cb = set(a.clave.split()) & COLORES, set(b.clave.split()) & COLORES
        va, vb = variante(a.clave), variante(b.clave)
        if ((na and nb and na != nb) or (a.talla and b.talla and a.talla != b.talla)
                or (ca and cb and ca != cb) or (va and vb and va != vb)):
            continue
        comparador = SequenceMatcher(None, a.clave, b.clave)
        if comparador.quick_ratio() >= umbral and comparador.ratio() >= umbral:
            pares.append((a.codigo, a.descripcion, b.codigo, b.descripcion, round(comparador.ratio(), 2)))
    return sorted(pares, key=lambda p: -p[4])


def escribir_tabla(hoja, encabezados, filas, anchos, fila_inicio=1):
    for col, titulo in enumerate(encabezados, 1):
        celda = hoja.cell(fila_inicio, col, titulo)
        celda.font, celda.fill = FUENTE_ENCABEZADO, COLOR_ENCABEZADO
        celda.alignment = Alignment(wrap_text=True, vertical="center")
    for f, fila in enumerate(filas, fila_inicio + 1):
        for col, valor in enumerate(fila, 1):
            hoja.cell(f, col, valor).font = FUENTE
    for col, ancho in enumerate(anchos, 1):
        hoja.column_dimensions[get_column_letter(col)].width = ancho
    ultima = fila_inicio + len(filas)
    hoja.freeze_panes = f"A{fila_inicio + 1}"
    hoja.auto_filter.ref = f"A{fila_inicio}:{get_column_letter(len(encabezados))}{ultima}"
    return ultima


def hoja_leeme(libro, maestro, n_compras, n_duplicados):
    hoja = libro.active
    hoja.title = "Léeme"
    hoja["A1"] = "Inventario maestro - BORRADOR"
    hoja["A1"].font = FUENTE_TITULO
    lineas = [
        f"Fuente: '{ENTRADA}', hoja {HOJA}. El archivo original no se modifica.",
        "Los códigos son provisionales: quedan fijos cuando se apruebe este borrador.",
        "Los precios de compra están SIN IGV y en la moneda indicada en cada fila.",
        "Maestro = datos propios del item. Compras = precios y proveedores, enlazados por el código.",
        "",
        "Qué revisar:",
        "1. Maestro: filas en amarillo, y las columnas Marca, Modelo y Talla, que son propuestas sacadas del texto.",
        "2. Posibles duplicados: escribir Sí en '¿Mismo item?' si dos filas son el mismo producto.",
        "   Ejemplo: 'GUANTES ANTIIMPACTO' y 'GUANTE ANTI IMPACTO' -> Sí",
        "3. Proveedores: confirmar los nombres unificados. El RUC del Excel original no es confiable.",
    ]
    for i, linea in enumerate(lineas, 3):
        hoja.cell(i, 1, linea).font = FUENTE

    inicio = 3 + len(lineas) + 1
    for col, titulo in enumerate(["Tipo de existencia", "Items", "Filas de compra"], 1):
        celda = hoja.cell(inicio, col, titulo)
        celda.font, celda.fill = FUENTE_ENCABEZADO, COLOR_ENCABEZADO
    for i, tipo in enumerate(PREFIJOS, inicio + 1):
        hoja.cell(i, 1, tipo).font = FUENTE
        hoja.cell(i, 2, f"=COUNTIF(Maestro!$B$2:$B${len(maestro) + 1},A{i})").font = FUENTE
        hoja.cell(i, 3, f"=COUNTIF(Compras!$L$2:$L${n_compras + 1},A{i})").font = FUENTE
    total = inicio + 1 + len(PREFIJOS)
    hoja.cell(total, 1, "Total").font = FUENTE_NEGRITA
    hoja.cell(total, 2, f"=SUM(B{inicio + 1}:B{total - 1})").font = FUENTE_NEGRITA
    hoja.cell(total, 3, f"=SUM(C{inicio + 1}:C{total - 1})").font = FUENTE_NEGRITA
    hoja.cell(total + 2, 1, f"Pares por revisar en 'Posibles duplicados': {n_duplicados}").font = FUENTE
    hoja.column_dimensions["A"].width = 26
    hoja.column_dimensions["B"].width = 10
    hoja.column_dimensions["C"].width = 16


def hoja_maestro(libro, maestro):
    """Solo datos propios del item: precios y proveedores viven en Compras, enlazados por el código."""
    hoja = libro.create_sheet("Maestro")
    filas = [[m.codigo, m.tipo, m.descripcion, m.marca, m.modelo, m.talla, m.unidad, m.revisar]
             for m in maestro.itertuples()]
    encabezados = ["Código", "Tipo de existencia", "Descripción", "Marca (propuesta)", "Modelo (propuesto)", "Talla",
                   "Unidad", "Revisar"]
    escribir_tabla(hoja, encabezados, filas, [11, 15, 70, 16, 18, 7, 9, 50])
    for f in range(2, len(filas) + 2):
        if hoja.cell(f, 8).value:
            for col in range(1, 9):
                hoja.cell(f, col).fill = COLOR_REVISAR


def hoja_compras(libro, compras, codigos):
    hoja = libro.create_sheet("Compras")
    filas = [[codigos[c.clave], c.Fecha.to_pydatetime(), c.descripcion, c.cant, c.unidad_original, c.unidad, c.MONEDA,
              c.precio, c.sub_total, c.total, c.tipo_original, c.tipo, c.proveedor_original, c.proveedor, c.ruc,
              c.factura if pd.notna(c.factura) else ""]
             for c in compras.rename(columns={"DESCRIPCIÓN": "descripcion", "CANT": "cant", "PRECIO UNIT": "precio",
                                              "SUB TOTAL": "sub_total", "TOTAL Incluido IGV": "total",
                                              "FACTURA": "factura"}).itertuples()]
    encabezados = ["Código", "Fecha", "Descripción original", "Cant", "U/M original", "Unidad", "Moneda",
                   "Precio unit. sin IGV", "Sub total", "Total con IGV", "Tipo original", "Tipo", "Proveedor original",
                   "Proveedor", "RUC", "Factura"]
    escribir_tabla(hoja, encabezados, filas, [11, 12, 62, 8, 9, 9, 10, 11, 11, 11, 14, 14, 40, 40, 14, 15])
    for f in range(2, len(filas) + 2):
        hoja.cell(f, 2).number_format = "dd/mm/yyyy"
        for col in (8, 9, 10):
            hoja.cell(f, col).number_format = "#,##0.00"


def hoja_cambios(libro, compras, eliminados, codigos):
    hoja = libro.create_sheet("Cambios de tipo")
    cambios = compras[compras["tipo"] != compras["tipo_original"]]
    filas = [[codigos[c["clave"]], c["DESCRIPCIÓN"], c["tipo_original"], c["tipo"]] for _, c in cambios.iterrows()]
    filas += [["", e["DESCRIPCIÓN"], e["TIPO EXISTENCIA"], "ELIMINADO"] for _, e in eliminados.iterrows()]
    escribir_tabla(hoja, ["Código", "Descripción", "Tipo original", "Tipo nuevo"], filas, [11, 70, 15, 15])


def hoja_duplicados(libro, pares):
    hoja = libro.create_sheet("Posibles duplicados")
    filas = [[*par, ""] for par in pares]
    encabezados = ["Código A", "Descripción A", "Código B", "Descripción B", "Similitud", "¿Mismo item?"]
    escribir_tabla(hoja, encabezados, filas, [11, 55, 11, 55, 10, 10])


def hoja_proveedores(libro, compras):
    hoja = libro.create_sheet("Proveedores")
    resumen = compras.groupby("proveedor").agg(
        nombres=("proveedor_original", lambda s: " | ".join(sorted(s.str.strip().unique()))),
        filas=("proveedor", "size"),
        rucs=("ruc", lambda s: ", ".join(sorted(r for r in s.unique() if r))),
    ).reset_index().sort_values("filas", ascending=False)
    usos_ruc = compras[compras["ruc"] != ""].groupby("ruc")["proveedor"].nunique()
    filas = []
    for r in resumen.itertuples():
        rucs = [x for x in r.rucs.split(", ") if x]
        problemas = []
        if len(rucs) > 1:
            problemas.append("Varios RUC")
        if any(usos_ruc[x] > 1 for x in rucs):
            problemas.append("RUC compartido con otro proveedor")
        filas.append([r.proveedor, r.nombres, r.filas, r.rucs, "; ".join(problemas)])
    escribir_tabla(hoja, ["Proveedor unificado", "Nombres en el Excel", "Filas", "RUC en el Excel", "Revisar"],
                   filas, [44, 70, 7, 26, 32])


def main():
    # El borrador ya tiene las respuestas del usuario: regenerarlo las borraría
    if os.path.exists(SALIDA) and "--sobrescribir" not in sys.argv:
        sys.exit(f"{SALIDA} ya existe y puede tener tus respuestas. Para reemplazarlo usa --sobrescribir")
    compras, eliminados = limpiar_compras()
    maestro = construir_maestro(compras, cargar_marcas())
    codigos = dict(zip(maestro["clave"], maestro["codigo"]))
    pares = posibles_duplicados(maestro)

    libro = Workbook()
    hoja_leeme(libro, maestro, len(compras), len(pares))
    hoja_maestro(libro, maestro)
    hoja_compras(libro, compras, codigos)
    hoja_duplicados(libro, pares)
    hoja_cambios(libro, compras, eliminados, codigos)
    hoja_proveedores(libro, compras)
    libro.save(SALIDA)

    print(f"Compras: {len(compras)} filas | eliminadas: {len(eliminados)}")
    print(f"Maestro: {len(maestro)} items | {maestro['tipo'].value_counts().to_dict()}")
    print(f"Filas por tipo: {compras['tipo'].value_counts().to_dict()}")
    print(f"Por revisar: {(maestro['revisar'] != '').sum()} | posibles duplicados: {len(pares)} -> {SALIDA}")


if __name__ == "__main__":
    main()
