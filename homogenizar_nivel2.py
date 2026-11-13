"""
Homogenización nivel 2 de las descripciones: vocabulario, con reemplazos que aprueba el usuario.

Dos pasos:
  python homogenizar_nivel2.py --borrador   crea "Diccionario nivel 2.xlsx" con las reglas propuestas,
                                             cuántas descripciones toca cada una y un ejemplo.
                                             No pisa un diccionario que ya tenga respuestas.
  python homogenizar_nivel2.py              aplica solo las reglas con aprobar = SI y crea
                                             "Datos consolidados de compras_nivel2.xlsx" (columna descripcion_n2).

Modos de cada regla:
  inicio   solo al comienzo de la descripción (plurales: no toca "JUEGO DE BROCAS" ni "MÁS TUERCAS")
  palabra  palabra completa en cualquier parte
  regex    expresión regular (para unidades y textos pegados)
  tildes   regla global: quita las tildes (la Ñ se mantiene)
"""
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ENTRADA = Path("Datos consolidados de compras_nivel1.xlsx")
DICCIONARIO = Path("Diccionario nivel 2.xlsx")
SALIDA = Path("Datos consolidados de compras_nivel2.xlsx")

LETRA = "A-ZÁÉÍÓÚÜÑ0-9"

# (grupo, modo, buscar, reemplazar, nota)
REGLAS = [
    # Tildes: una sola forma por palabra
    ("TILDES", "tildes", "(todas)", "(sin tilde)",
     "Hoy conviven TUBERIA/TUBERÍA, LINEA/LÍNEA, DIELECTRICO/DIELÉCTRICO y unas 25 parejas más. La Ñ se mantiene."),
    # Textos pegados
    ("PEGADO", "regex", r"PROTECCION(?=\d)", "PROTECCION ", "Va después de quitar tildes."),
    ("PEGADO", "regex", r"CARBONO(?=\d)", "CARBONO ", None),
    ("PEGADO", "regex", r"(?<=[\d ])MT\.X\s*(?=[\d.])", "MT X ", "130MT.X1/4 y 1.00 MT.X .1/2."),
    ("PEGADO", "regex", r"(?<=\d)MX(?=\d)", "M X ", "3MX50M."),
    ("PEGADO", "regex", r"MTSANCHO", "MTS ANCHO", None),
    ("PEGADO", "regex", r"PORCELANATIPO", "PORCELANA TIPO", None),
    ("PEGADO", "regex", r"(?<=\d)PINES", " PINES", "11PINES."),
    ("PEGADO", "regex", r"(?<=\d)AMPERIOS", " AMPERIOS", "10-14AMPERIOS."),
    # Plural -> singular, solo la palabra con la que empieza la descripción (y su adjetivo)
    ("PLURAL", "inicio", "ABRAZADERAS", "ABRAZADERA", None),
    ("PLURAL", "inicio", "ANGULOS", "ANGULO", None),
    ("PLURAL", "inicio", "ARANDELAS PLANAS", "ARANDELA PLANA", None),
    ("PLURAL", "inicio", "ARANDELAS", "ARANDELA", None),
    ("PLURAL", "inicio", "BARBIQUEJOS", "BARBIQUEJO", None),
    ("PLURAL", "inicio", "BISAGRAS", "BISAGRA", None),
    ("PLURAL", "inicio", "BROCHAS", "BROCHA", None),
    ("PLURAL", "inicio", "CINTAS AISLANTES", "CINTA AISLANTE", None),
    ("PLURAL", "inicio", "DISCOS", "DISCO", None),
    ("PLURAL", "inicio", "DRIZAS", "DRIZA", None),
    ("PLURAL", "inicio", "FILTROS", "FILTRO", None),
    ("PLURAL", "inicio", "GRAPAS", "GRAPA", None),
    ("PLURAL", "inicio", "GUANTES", "GUANTE", None),
    ("PLURAL", "inicio", "JEANS", "JEAN", None),
    ("PLURAL", "inicio", "LENTES", "LENTE", "En EPP se dice 'lentes' para un solo par; hoy conviven LENTE y LENTES."),
    ("PLURAL", "inicio", "LUNAS TRANSPARENTES", "LUNA TRANSPARENTE", None),
    ("PLURAL", "inicio", "MANTAS IGNIFUGAS", "MANTA IGNIFUGA", None),
    ("PLURAL", "inicio", "MANTAS", "MANTA", None),
    ("PLURAL", "inicio", "MICAS", "MICA", None),
    ("PLURAL", "inicio", "PERNOS INOXIDABLES", "PERNO INOXIDABLE", None),
    ("PLURAL", "inicio", "PERNOS", "PERNO", None),
    ("PLURAL", "inicio", "PLANCHAS", "PLANCHA", None),
    ("PLURAL", "inicio", "PLATINAS GALVANIZADOS", "PLATINA GALVANIZADA", "También corrige el género (platina es femenino)."),
    ("PLURAL", "inicio", "POLOS", "POLO", None),
    ("PLURAL", "inicio", "RODILLOS", "RODILLO", None),
    ("PLURAL", "inicio", "SOBRELENTES", "SOBRELENTE", None),
    ("PLURAL", "inicio", "TABLEROS", "TABLERO", None),
    ("PLURAL", "inicio", "TACOS", "TACO", None),
    ("PLURAL", "inicio", "TAPAS PROTECTORAS", "TAPA PROTECTORA", None),
    ("PLURAL", "inicio", "TECHOS METALICOS", "TECHO METALICO", "Va después de quitar tildes."),
    ("PLURAL", "inicio", "TORNILLOS AUTOPERFORANTES", "TORNILLO AUTOPERFORANTE", None),
    ("PLURAL", "inicio", "TORNILLOS AUTORROSCANTES", "TORNILLO AUTORROSCANTE", None),
    ("PLURAL", "inicio", "TORNILLOS", "TORNILLO", None),
    ("PLURAL", "inicio", "TUBERIAS", "TUBERIA", None),
    ("PLURAL", "inicio", "TUERCAS", "TUERCA", None),
    ("PLURAL", "inicio", "VIGAS SECUNDARIAS", "VIGA SECUNDARIA", None),
    ("PLURAL", "inicio", "VIGAS", "VIGA", None),
    ("PLURAL", "inicio", "CARETA FACIALES TRANSPARENTES", "CARETA FACIAL TRANSPARENTE", "Concordancia: careta es singular."),
    # Abreviaturas -> palabra completa
    ("ABREVIATURA", "palabra", "INOX.", "INOXIDABLE", None),
    ("ABREVIATURA", "palabra", "INOX", "INOXIDABLE", "Hoy: INOX en 24 descripciones, INOXIDABLE en 10."),
    ("ABREVIATURA", "palabra", "GALV", "GALVANIZADA", "Solo aparece en ARANDELA PLANA GALV (femenino)."),
    ("ABREVIATURA", "palabra", "AC.", "ACERO", "Como en 'AC. INOX. 304'."),
    ("ABREVIATURA", "palabra", "HEX.", "HEXAGONAL", None),
    ("ABREVIATURA", "palabra", "HEX", "HEXAGONAL", None),
    ("ABREVIATURA", "regex", r"(?<![A-Z])P/\s*", "PARA ", "P/ SOLDADOR, P/TUERCA, P/PUESTA."),
    ("ABREVIATURA", "regex", r"(?<![A-Z])C/(?![AB]\b)\s*", "CON ", "C/ DOS VIAS, C/TAPA, C/DOBLE. No toca C/A ni C/B (componente A y B)."),
    ("ABREVIATURA", "regex", r"(?<![A-Z])T/\s+", "TIPO ", "Solo 'SOGA NYLON T/ DRIZA'. No toca T/39 (talla)."),
    ("ABREVIATURA", "palabra", "P.BROCA", "PUNTA BROCA", None),
    ("ABREVIATURA", "palabra", "AUTOPERFORANT", "AUTOPERFORANTE", None),
    ("ABREVIATURA", "palabra", "AUTOPERFOR", "AUTOPERFORANTE", None),
    ("ABREVIATURA", "palabra", "ESTRUCT.", "ESTRUCTURAL", None),
    ("ABREVIATURA", "regex", r"#\s*(?=\d)", "N°", "ALAMBRE #8 y ALAMBRE RECOCIDO N°8 quedan igual."),
    # Unidades de medida dentro del texto
    ("UNIDAD", "regex", r"(?<=\d)\s*(?:METROS|MTRS|MTS|MT)(?![A-Z])", " M",
     "Solo después de un número: '3 MTS', '6MT', '8 METROS' pasan a '3 M', '6 M', '8 M'."),
    ("UNIDAD", "palabra", "LBS", "LB", "COMBA 4 LBS / COMBA 6LBS."),
    # Erratas evidentes
    ("ERRATA", "palabra", "AUDITIVIO", "AUDITIVO", None),
    ("ERRATA", "palabra", "TRAJW", "TRAJE", None),
    ("ERRATA", "palabra", "TYVECK", "TYVEK", None),
    ("ERRATA", "palabra", "TYVEX", "TYVEK", None),
    ("ERRATA", "palabra", "SOLDARUDA", "SOLDADURA", None),
    ("ERRATA", "palabra", "TIRRAFON", "TIRAFON", None),
    ("ERRATA", "palabra", "ILO CORRIDO", "HILO CORRIDO", None),
    ("ERRATA", "palabra", "AMALODARA", "AMOLADORA", None),
    ("ERRATA", "palabra", "MADERIA", "MADERA", None),
    ("ERRATA", "palabra", "PLATFORMA", "PLATAFORMA", None),
    ("ERRATA", "palabra", "HARTOP", "HARDTOP", None),
    ("ERRATA", "palabra", "CERO AL CARBONO", "ACERO AL CARBONO", None),
    ("ERRATA", "palabra", "ANTIMPACTO", "ANTIIMPACTO", "Hoy conviven ANTIMPACTO, ANTIIMPACTO y ANTI IMPACTO."),
    ("ERRATA", "palabra", "ANTI IMPACTO", "ANTIIMPACTO", None),
    ("ERRATA", "palabra", "ENGRANPADOR", "ENGRAPADOR", None),
    ("ERRATA", "palabra", "RIGDA", "RIGIDA", "Va después de quitar tildes."),
    ("ERRATA", "palabra", "METATARZALES", "METATARSALES", None),
    ("ERRATA", "palabra", "NOVALACA", "NOVOLACA", None),
    ("ERRATA", "palabra", "CICLOALINFATICA", "CICLOALIFATICA", "Va después de quitar tildes."),
    ("ERRATA", "palabra", "POLYMERICO", "POLIMERICO", None),
    ("ERRATA", "palabra", "LIJAR DE", "LIJA DE", "'Lijar de 1000': lija de grano 1000."),
    ("ERRATA", "palabra", "PARRAYOS", "PARARRAYOS", None),
    ("ERRATA", "palabra", "TIPO TIPO", "TIPO", "Palabra repetida."),
]

COLUMNAS = ["n", "grupo", "modo", "buscar", "reemplazar", "descripciones", "ejemplo_antes", "ejemplo_despues",
            "nota", "aprobar"]


def sin_tildes(texto):
    texto = texto.replace("Ñ", "\0")
    texto = "".join(c for c in unicodedata.normalize("NFD", texto) if not unicodedata.combining(c))
    return texto.replace("\0", "Ñ")


def patron(modo, buscar):
    if modo == "inicio":
        return re.compile(rf"^{re.escape(buscar)}(?![{LETRA}])")
    if modo == "palabra":
        return re.compile(rf"(?<![{LETRA}]){re.escape(buscar)}(?![{LETRA}])")
    return re.compile(buscar)


def aplicar_regla(texto, modo, buscar, reemplazar):
    if modo == "tildes":
        return sin_tildes(texto)
    return re.sub(r"\s+", " ", patron(modo, buscar).sub(reemplazar, texto)).strip()


def aplicar(texto, reglas):
    for modo, buscar, reemplazar in reglas:
        texto = aplicar_regla(texto, modo, buscar, reemplazar)
    return texto


def estilo(ws, anchos):
    for fila in ws.iter_rows():
        for celda in fila:
            celda.font = Font(name="Arial", size=10)
            celda.alignment = Alignment(vertical="top", wrap_text=True)
    for celda in ws[1]:
        celda.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")
    for i, campo in enumerate([c.value for c in ws[1]], 1):
        ws.column_dimensions[get_column_letter(i)].width = anchos.get(campo, 14)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def borrador():
    if DICCIONARIO.exists():
        sys.exit(f"{DICCIONARIO} ya existe y puede tener respuestas; bórralo a mano si quieres rehacerlo")
    descripciones = sorted(set(pd.read_excel(ENTRADA)["descripcion_n1"]))
    # Cada regla se mide sobre el texto que le llega, con las anteriores ya aplicadas
    actuales = {d: d for d in descripciones}
    filas, sin_uso = [], []
    for grupo, modo, buscar, reemplazar, nota in REGLAS:
        tocadas = []
        for d, texto in actuales.items():
            nuevo = aplicar_regla(texto, modo, buscar, reemplazar)
            if nuevo != texto:
                tocadas.append((d, texto, nuevo))
                actuales[d] = nuevo
        if not tocadas:
            sin_uso.append(buscar)
            continue
        _, antes, despues = tocadas[0]
        filas.append([len(filas) + 1, grupo, modo, buscar, reemplazar, len(tocadas), antes, despues, nota, None])
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reglas"
    ws.append(COLUMNAS)
    for f in filas:
        ws.append(f)
    estilo(ws, {"buscar": 22, "reemplazar": 22, "ejemplo_antes": 50, "ejemplo_despues": 50, "nota": 45,
                "n": 5, "modo": 9, "descripciones": 12, "aprobar": 10})
    amarillo = PatternFill("solid", fgColor="FFF2CC")
    for celda in ws["J"][1:]:
        celda.fill = amarillo
    leeme = wb.create_sheet("Léeme")
    for linea in [
        ["Diccionario de homogenización nivel 2"],
        ["Escribe SI en la columna 'aprobar' de cada regla que quieras aplicar. Lo vacío o NO no se aplica."],
        ["Se puede cambiar el texto de 'reemplazar'. Las reglas se aplican en el orden de la columna 'n'."],
        ["'descripciones' y los ejemplos se calcularon con las reglas anteriores ya aplicadas."],
        ["modo inicio = solo al comienzo de la descripción; palabra = palabra completa; regex = patrón."],
        ["La descripción original y la del nivel 1 se conservan; el resultado va en descripcion_n2."],
    ]:
        leeme.append(linea)
    leeme.column_dimensions["A"].width = 110
    wb.move_sheet("Léeme", offset=-1)
    wb.save(DICCIONARIO)
    print(f"Guardado {DICCIONARIO}: {len(filas)} reglas")
    if sin_uso:
        print("Reglas sin ninguna descripción que tocar (no se incluyeron):", sin_uso)


def aplicar_aprobadas():
    reglas = pd.read_excel(DICCIONARIO, sheet_name="Reglas")
    aprobadas = reglas[reglas["aprobar"].astype(str).str.strip().str.upper() == "SI"].sort_values("n")
    lista = [(r.modo, r.buscar, "" if pd.isna(r.reemplazar) else r.reemplazar) for r in aprobadas.itertuples()]
    tabla = pd.read_excel(ENTRADA)
    tabla.insert(tabla.columns.get_loc("descripcion_n1") + 1, "descripcion_n2",
                 tabla["descripcion_n1"].map(lambda d: aplicar(d, lista)))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Compras"
    ws.append(list(tabla.columns))
    for fila in tabla.itertuples(index=False):
        ws.append([None if pd.isna(v) else (v.date() if isinstance(v, pd.Timestamp) else v) for v in fila])
    estilo(ws, {c: 55 for c in tabla.columns if c.startswith("descripcion")} | {"proveedor": 40})
    for i, campo in enumerate(tabla.columns, 1):
        for celda in ws[get_column_letter(i)][1:]:
            celda.alignment = Alignment(vertical="top")
            if campo in ("precio_total", "precio_unitario"):
                celda.number_format = "#,##0.00"
            elif campo == "fecha":
                celda.number_format = "yyyy-mm-dd"
    cambios = tabla[tabla["descripcion_n1"] != tabla["descripcion_n2"]][["descripcion_n1", "descripcion_n2"]]
    ws2 = wb.create_sheet("Cambios nivel 2")
    ws2.append(["descripcion_n1", "descripcion_n2"])
    for fila in cambios.drop_duplicates().sort_values("descripcion_n2").itertuples(index=False):
        ws2.append(list(fila))
    estilo(ws2, {"descripcion_n1": 70, "descripcion_n2": 70})
    wb.save(SALIDA)
    print(f"Reglas aplicadas: {len(lista)} de {len(reglas)}")
    print(f"Descripciones distintas: nivel 1 {tabla['descripcion_n1'].nunique()} | nivel 2 {tabla['descripcion_n2'].nunique()}")
    print(f"Guardado {SALIDA}")


if __name__ == "__main__":
    borrador() if "--borrador" in sys.argv else aplicar_aprobadas()
