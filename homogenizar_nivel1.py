"""
Homogenización nivel 1 de las descripciones: solo formato, ninguna palabra cambia.

- Todo en MAYÚSCULAS (las tildes se mantienen).
- Un solo espacio entre palabras.
- Pulgadas siempre con comilla doble: ” “ ″ y '' pasan a ".
- Separador de medidas siempre " X ": el * y la x pegada entre dos medidas (2"x2", 3/16x6MT).

La descripción original se conserva; la nueva va en descripcion_n1.
Comprobación: sin espacios, comillas ni separadores, el texto debe ser idéntico al original.

Uso:  python homogenizar_nivel1.py
"""
import re
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ENTRADA = Path("Datos consolidados de compras.xlsx")
SALIDA = Path("Datos consolidados de compras_nivel1.xlsx")

COMILLAS = re.compile(r"[”“″]|''")
# Asterisco suelto (no el '**' de AUTOPERFORANT ... C/A **/)
ASTERISCO = re.compile(r"\s*(?<!\*)\*(?!\*)\s*")
# x pegada o espaciada entre dos medidas: antes un número o pulgada, después un número
X_MEDIDA = re.compile(r"(?<=[\d\"½¾])\s*[xX]\s*(?=[\d.½¾])")


def nivel1(texto):
    texto = COMILLAS.sub('"', str(texto))
    texto = ASTERISCO.sub(" X ", texto)
    texto = X_MEDIDA.sub(" X ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto.upper()


def esqueleto(texto):
    """El texto sin lo que el nivel 1 puede cambiar: espacios, comillas y separadores."""
    texto = COMILLAS.sub('"', str(texto)).upper()
    return re.sub(r"[\s\*X\"]", "", texto)


def guardar(tabla, cambios):
    wb = openpyxl.Workbook()
    hojas = [("Compras", tabla), ("Cambios nivel 1", cambios)]
    for i, (nombre, df) in enumerate(hojas):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = nombre
        ws.append(list(df.columns))
        for fila in df.itertuples(index=False):
            ws.append([None if pd.isna(v) else (v.date() if isinstance(v, pd.Timestamp) else v) for v in fila])
        for fila in ws.iter_rows():
            for celda in fila:
                celda.font = Font(name="Arial", size=10)
        for celda in ws[1]:
            celda.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            celda.fill = PatternFill("solid", fgColor="1F4E78")
            celda.alignment = Alignment(vertical="center")
        for j, campo in enumerate(df.columns, 1):
            letra = get_column_letter(j)
            ancho = 60 if campo.startswith("descripcion") else 45 if campo == "proveedor" else 14
            ws.column_dimensions[letra].width = ancho
            for celda in ws[letra][1:]:
                if campo in ("precio_total", "precio_unitario"):
                    celda.number_format = "#,##0.00"
                elif campo == "fecha":
                    celda.number_format = "yyyy-mm-dd"
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    wb.save(SALIDA)


if __name__ == "__main__":
    tabla = pd.read_excel(ENTRADA)
    tabla.insert(tabla.columns.get_loc("descripcion") + 1, "descripcion_n1", tabla["descripcion"].map(nivel1))
    distintas = tabla[tabla["descripcion"].map(lambda d: re.sub(r"\s+", " ", str(d)).strip())
                      != tabla["descripcion_n1"]]
    malas = [d for d in tabla["descripcion"].unique() if esqueleto(d) != esqueleto(nivel1(d))]
    if malas:
        raise SystemExit(f"El nivel 1 cambió palabras en: {malas}")
    cambios = (distintas[["descripcion", "descripcion_n1"]].drop_duplicates()
               .rename(columns={"descripcion": "descripcion_original"}).sort_values("descripcion_n1"))
    guardar(tabla, cambios)
    print(f"Filas: {len(tabla)} | descripciones distintas antes {tabla['descripcion'].nunique()} "
          f"| después {tabla['descripcion_n1'].nunique()}")
    print(f"Descripciones que cambiaron de forma: {len(cambios)}")
    print(f"Guardado {SALIDA}")
