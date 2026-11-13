"""
Consolidado de compras: junta las filas de las hojas del inventario en una sola tabla, tal cual.

- Lee la copia congelada del inventario (nunca el archivo de Nextcloud).
- Solo toma celdas escritas a mano: las fórmulas (dólares, precio unitario, cotización) se ignoran.
- Moneda: si la fila tiene soles escritos, la compra fue en soles (PEN); si solo tiene dólares
  escritos, fue en dólares (USD). No se convierte nada.
- Precio: total de la fila con IGV y unitario final (total / cantidad), ambos con 2 decimales.
- Sin proveedor ni fecha la fila es una COTIZACION; con ambos, una COMPRA.
- Tipo de existencia: el de la hoja (TIPO_POR_HOJA), no el de la columna TIPO EXISTENCIA.
- Hojas de pedidos (HOJAS_PEDIDO): moneda de su columna MONEDA y precio final con IGV.
- Filas sin fecha en su hoja: FECHA_SIN_DATO, marcadas en fecha_origen como ASIGNADA.

Se regenera completo en cada corrida con las hojas de HOJAS.
Uso:  python consolidar_compras.py
"""
import re
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FUENTE = Path("datos/bronce/inventario/inventario_maestro_2026-11-05_1342.xlsx")
SALIDA = Path("Datos consolidados de compras.xlsx")
FILA_CABECERA = 5

# Hojas ya revisadas, en el orden en que se agregan
HOJAS = ["Herramienta", "Servicio", "Equipo", "Movilidad", "EPP", "Material"]

# En las hojas por categoría el tipo de existencia lo da la hoja, no su columna TIPO EXISTENCIA.
# Presupuesto1 y Presupuesto2 no son categorías: no van aquí.
TIPO_POR_HOJA = {"Material": "MATERIAL", "EPP": "EPP", "Herramienta": "HERRAMIENTA",
                 "Equipo": "EQUIPO", "Servicio": "SERVICIO", "Movilidad": "MOVILIDAD"}

# Hojas de pedidos (columna MONEDA, precio unitario escrito, sin fecha).
# Precio = cantidad x unitario (x 1.18 si el unitario es sin IGV), la misma cuenta del Excel;
# si solo hay total con IGV escrito, se usa ese.
#   tipo: tipo de existencia; None = el de su columna TIPO EXISTENCIA (hojas que no son categoría)
#   desde_fila: Presupuesto2 hasta la fila 183 es copia del listado antiguo, ya incluido por las categorías
HOJAS_PEDIDO = {
    "transporte": {"tipo": "MOVILIDAD", "fila_cabecera": 2},
    "Presupuesto2": {"tipo": None, "fila_cabecera": 1, "desde_fila": 184},
    "Presupuesto1": {"tipo": "MATERIAL", "fila_cabecera": 2, "unitario_con_igv": True},
}
# Encabezados de Presupuesto1 -> los de las otras hojas de pedido
ALIAS_PEDIDO = {"DESCRIPCION": "DESCRIPCIÓN", "CANTIDAD": "CANT", "PRECIO UNITARIO INCL IGV": "PRECIO UNIT",
                "PRECIO TOTAL INCL IGV": "TOTAL Incluido IGV"}
IGV = 1.18
MONEDAS = {"SOLES": "PEN", "DÓLARES": "USD", "DOLARES": "USD"}

# Fecha para las filas cuya hoja no la tiene (decisión del usuario: la del día de la consolidación).
# fecha_origen dice si la fecha vino de la hoja o se asignó, para reemplazarla cuando llegue la real.
FECHA_SIN_DATO = date(2026, 11, 5)

# Excepciones acordadas con el usuario: (hoja, fila) -> tipo. origen_hoja sigue mostrando de dónde vino
CAMBIOS_TIPO = {("Herramienta", 6): "EPP",    # ARNES HAUK
                ("Herramienta", 15): "EPP"}   # LINEA RETRACTIL DE 3 MTRS DE CABLE

# Encabezados de la hoja -> campo (las hojas usan nombres algo distintos)
COLUMNAS = {
    "DESCRIPCIÓN": "descripcion",
    "Cant": "cantidad",
    "U/M": "unidad",
    "PRECIO SOLES": "soles",
    "PRECIO SOLES (SOLES)": "soles",
    "PRECIO DOLARES": "dolares",
    "PRECIO DOLARES (USD)": "dolares",
    "PROVEEDOR": "proveedor",
    "RUC": "ruc",
    "Fecha": "fecha",
}

CAMPOS = ["fuente", "origen_hoja", "origen_fila", "tipo_registro", "fecha", "fecha_origen", "descripcion", "cantidad",
          "unidad", "precio_total", "precio_unitario", "moneda", "tipo_existencia", "proveedor", "ruc"]
ANCHOS = {"fuente": 38, "descripcion": 60, "proveedor": 45, "tipo_existencia": 16, "tipo_registro": 14}


def es_formula(valor):
    return isinstance(valor, str) and valor.startswith("=")


def es_operacion_escrita(valor):
    """Una 'fórmula' que solo tiene números, como =3000/2: es un monto escrito a mano."""
    return es_formula(valor) and re.fullmatch(r"=[\d\s.,+\-*/()]+", valor) is not None


def escrito(formula, resultado):
    """Valor escrito a mano en la celda; None si la celda depende de otras."""
    if es_operacion_escrita(formula):
        return resultado
    return None if es_formula(formula) else formula


def unitario_final(total, cantidad):
    """Precio unitario con IGV, calculado con el total sin redondear."""
    if isinstance(cantidad, (int, float)) and cantidad > 0:
        return round(total / cantidad, 2)
    return None


def vacio(valor):
    return valor is None or (isinstance(valor, str) and valor.strip() in ("", "-"))


def texto(valor):
    if vacio(valor):
        return None
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip()


def extraer(ws, wv, hoja):
    """ws: hoja con fórmulas; wv: la misma hoja con los valores calculados."""
    cabecera = {c.column: COLUMNAS[c.value] for c in ws[FILA_CABECERA] if c.value in COLUMNAS}
    filas, problemas, montos = [], [], []
    for n in range(FILA_CABECERA + 1, ws.max_row + 1):
        celdas = {campo: ws.cell(n, col).value for col, campo in cabecera.items()}
        if all(v is None for v in celdas.values()):
            continue
        for campo, valor in celdas.items():
            if campo not in ("soles", "dolares") and es_formula(valor):
                problemas.append(f"fila {n}: '{campo}' es una fórmula")
        col = {campo: c for c, campo in cabecera.items()}
        soles = escrito(celdas["soles"], wv.cell(n, col["soles"]).value)
        dolares = escrito(celdas.get("dolares"), wv.cell(n, col["dolares"]).value)
        for campo in ("soles", "dolares"):
            if es_operacion_escrita(celdas.get(campo)):
                problemas.append(f"fila {n}: {campo} escrito como operación {celdas[campo]}")
        if not vacio(soles) and not vacio(dolares):
            problemas.append(f"fila {n}: soles y dólares escritos a mano, no se sabe la moneda original")
            continue
        if not vacio(soles):
            moneda, monto = "PEN", float(soles)
        elif not vacio(dolares):
            moneda, monto = "USD", float(dolares)
        else:
            problemas.append(f"fila {n}: sin monto escrito")
            continue
        proveedor, fecha = texto(celdas.get("proveedor")), celdas.get("fecha")
        if (proveedor is None) != vacio(fecha):
            problemas.append(f"fila {n}: tiene proveedor o fecha, pero no ambos")
        cantidad = celdas["cantidad"]
        if isinstance(cantidad, str):
            try:
                cantidad = float(cantidad.strip().replace(",", "."))
                cantidad = int(cantidad) if cantidad.is_integer() else cantidad
                problemas.append(f"fila {n}: cantidad escrita como texto '{celdas['cantidad']}', se guarda como número")
            except ValueError:
                problemas.append(f"fila {n}: cantidad no numérica '{celdas['cantidad']}'")
        montos.append((moneda, monto))
        filas.append({
            "fuente": FUENTE.name, "origen_hoja": hoja, "origen_fila": n,
            "tipo_registro": "COMPRA" if proveedor and not vacio(fecha) else "COTIZACION",
            "fecha": fecha.date() if isinstance(fecha, datetime) else texto(fecha),
            "descripcion": texto(celdas["descripcion"]),
            "cantidad": cantidad,
            "unidad": texto(celdas["unidad"]),
            "precio_total": round(monto, 2), "precio_unitario": unitario_final(monto, cantidad), "moneda": moneda,
            "tipo_existencia": CAMBIOS_TIPO.get((hoja, n), TIPO_POR_HOJA[hoja]),
            "proveedor": proveedor, "ruc": texto(celdas.get("ruc")),
        })
    return filas, problemas, montos


def extraer_pedido(ws, wv, hoja):
    """Hojas con columna MONEDA y sin fecha. Se comprueba contra el total con IGV que calcula el Excel."""
    conf = HOJAS_PEDIDO[hoja]
    fila_cab = conf["fila_cabecera"]
    factor = 1 if conf.get("unitario_con_igv") else IGV
    col = {ALIAS_PEDIDO.get(c.value, c.value): c.column for c in ws[fila_cab] if c.value}
    filas, problemas = [], []
    suma_excel = {"PEN": 0.0, "USD": 0.0}
    for n in range(conf.get("desde_fila", fila_cab + 1), ws.max_row + 1):
        if vacio(ws.cell(n, col["DESCRIPCIÓN"]).value):
            continue
        valor = lambda campo: ws.cell(n, col[campo]).value if campo in col else None
        calculado = lambda campo: wv.cell(n, col[campo]).value
        moneda = MONEDAS.get(texto(valor("MONEDA")) or "")
        cantidad = escrito(valor("CANT"), calculado("CANT"))
        unitario = escrito(valor("PRECIO UNIT"), calculado("PRECIO UNIT"))
        total = escrito(valor("TOTAL Incluido IGV"), calculado("TOTAL Incluido IGV"))
        if moneda is None:
            problemas.append(f"fila {n}: moneda '{valor('MONEDA')}' no reconocida")
            continue
        if cantidad is None or isinstance(cantidad, str):
            problemas.append(f"fila {n}: cantidad '{valor('CANT')}' no se puede leer")
            continue
        if not vacio(unitario):
            precio, precio_unitario = float(cantidad) * float(unitario) * factor, round(float(unitario) * factor, 2)
        elif not vacio(total):
            precio = float(total)
            precio_unitario = unitario_final(precio, cantidad)
        else:
            problemas.append(f"fila {n}: sin precio escrito, no entra")
            continue
        suma_excel[moneda] += float(calculado("TOTAL Incluido IGV") or 0)
        proveedor = texto(valor("PROVEEDOR"))
        filas.append({
            "fuente": FUENTE.name, "origen_hoja": hoja, "origen_fila": n,
            "tipo_registro": "COMPRA" if proveedor else "COTIZACION", "fecha": None,
            "descripcion": texto(valor("DESCRIPCIÓN")), "cantidad": cantidad, "unidad": texto(valor("U/M")),
            "precio_total": round(precio, 2), "precio_unitario": precio_unitario, "moneda": moneda,
            "tipo_existencia": CAMBIOS_TIPO.get((hoja, n), conf["tipo"] or texto(valor("TIPO EXISTENCIA"))),
            "proveedor": proveedor, "ruc": None,
        })
    print(f"{hoja}: filas con descripción {len(filas) + len(problemas)} | consolidado {len(filas)}")
    for m in suma_excel:
        mia = sum(f["precio_total"] for f in filas if f["moneda"] == m)
        print(f"  {m}: total del Excel {suma_excel[m]:,.6f} | consolidado {mia:,.2f}")
        if abs(mia - suma_excel[m]) > 0.01 * max(1, len(filas)):
            problemas.append(f"{m}: el total no cuadra con el del Excel")
    return filas, problemas


def comprobar(ws, hoja, filas, montos):
    """Mismas filas que la hoja y la misma suma por moneda (antes de redondear)."""
    col_desc = next(c.column for c in ws[FILA_CABECERA] if c.value == "DESCRIPCIÓN")
    n_hoja = sum(1 for n in range(FILA_CABECERA + 1, ws.max_row + 1) if not vacio(ws.cell(n, col_desc).value))
    print(f"{hoja}: filas hoja {n_hoja} | consolidado {len(filas)}")
    for m in ("PEN", "USD"):
        exacto = sum(v for mon, v in montos if mon == m)
        redondeado = sum(f["precio_total"] for f in filas if f["moneda"] == m)
        print(f"  {m}: suma exacta {exacto:,.6f} | suma con 2 decimales {redondeado:,.2f}")
    if n_hoja != len(filas):
        sys.exit(f"{hoja}: el consolidado no tiene las mismas filas que la hoja")


def guardar(filas):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Compras"
    ws.append(CAMPOS)
    for f in filas:
        ws.append([f[c] for c in CAMPOS])
    fuente = Font(name="Arial", size=10)
    for fila in ws.iter_rows():
        for celda in fila:
            celda.font = fuente
    for celda in ws[1]:
        celda.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")
        celda.alignment = Alignment(vertical="center")
    for i, campo in enumerate(CAMPOS, 1):
        letra = get_column_letter(i)
        ws.column_dimensions[letra].width = ANCHOS.get(campo, 12)
        for celda in ws[letra][1:]:
            if campo in ("precio_total", "precio_unitario"):
                celda.number_format = "#,##0.00"
            elif campo == "fecha":
                celda.number_format = "yyyy-mm-dd"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(SALIDA)


if __name__ == "__main__":
    libro = openpyxl.load_workbook(FUENTE)
    valores = openpyxl.load_workbook(FUENTE, data_only=True)
    todas = []
    for hoja in HOJAS:
        filas, problemas, montos = extraer(libro[hoja], valores[hoja], hoja)
        if problemas:
            print(f"{hoja} - revisar:", *problemas, sep="\n  ")
        comprobar(libro[hoja], hoja, filas, montos)
        todas += filas
    for hoja in HOJAS_PEDIDO:
        filas, problemas = extraer_pedido(libro[hoja], valores[hoja], hoja)
        if problemas:
            print(f"{hoja} - revisar:", *problemas, sep="\n  ")
            if any("no cuadra" in p or "no se puede leer" in p or "no reconocida" in p for p in problemas):
                sys.exit(f"{hoja}: no se guarda hasta revisar")
        todas += filas
    for f in todas:
        f["fecha_origen"] = "HOJA" if f["fecha"] else "ASIGNADA"
        f["fecha"] = f["fecha"] or FECHA_SIN_DATO
    guardar(todas)
    print(f"Guardado {SALIDA} ({len(todas)} filas)")
