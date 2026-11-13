import re
import time
import html

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# API pública de WooCommerce (la misma que usa la tienda por dentro): devuelve JSON, no HTML
API_URL = "https://promiconperu.com/wp-json/wc/store/v1/products"
POR_PAGINA = 100          # máximo que permite la API
MAX_PAGINAS = None        # None = todo el catálogo; pon un número para pruebas
TIPO_CAMBIO = 3.75
PAUSA_SEG = 1.0           # pausa entre páginas para no saturar el servidor
SALIDA = "Catalogo_Promicon.xlsx"

MARCAS = ["TRUPER", "STANLEY", "3M", "MAKITA", "MILWAUKEE", "BOSCH", "TRAMONTINA", "PRETUL", "CLUTE"]
# Límites de palabra para que "3M" no coincida con "13MM"
MARCAS_RE = re.compile(r"\b(" + "|".join(map(re.escape, MARCAS)) + r")\b", re.IGNORECASE)
# Muchos títulos/descripciones traen "MARCA : LOTO MASTER" o "Marca: 3M"
MARCA_TEXTO_RE = re.compile(r"MARCA\s*:\s*([A-Z0-9][A-Z0-9&\-\. ]{0,30}?)(?=\s+[A-Z0-9\-]*\d|,|<|$)", re.IGNORECASE)
TAGS_RE = re.compile(r"<[^>]+>")


def crear_sesion():
    sesion = requests.Session()
    sesion.headers["User-Agent"] = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    reintentos = Retry(total=5, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
    sesion.mount("https://", HTTPAdapter(max_retries=reintentos))
    return sesion


def limpiar(texto):
    return html.unescape(TAGS_RE.sub(" ", texto or "")).strip()


def detectar_marca(producto):
    if producto.get("brands"):
        return producto["brands"][0]["name"].upper()
    nombre = limpiar(producto["name"])
    descripcion = limpiar(producto.get("description"))
    for texto in (nombre, descripcion):
        m = MARCA_TEXTO_RE.search(texto)
        if m:
            return m.group(1).strip().upper()
    m = MARCAS_RE.search(nombre)
    return m.group(1).upper() if m else "OTRO"


def precio_soles(producto):
    precios = producto["prices"]
    # La API da el precio en céntimos: "3749" con minor_unit 2 = S/ 37.49
    valor = int(precios.get("price") or 0) / 10 ** precios.get("currency_minor_unit", 2)
    return valor or None  # 0 en la tienda significa "consultar precio"


def descargar_catalogo(sesion):
    productos, pagina, total_paginas = [], 1, None
    while True:
        resp = sesion.get(API_URL, params={"per_page": POR_PAGINA, "page": pagina}, timeout=30)
        resp.raise_for_status()
        if total_paginas is None:
            total_paginas = int(resp.headers.get("X-WP-TotalPages", 1))
            if MAX_PAGINAS:
                total_paginas = min(total_paginas, MAX_PAGINAS)
            print(f"Productos en la tienda: {resp.headers.get('X-WP-Total')} | páginas a leer: {total_paginas}")
        lote = resp.json()
        productos.extend(lote)
        print(f"  página {pagina}/{total_paginas}: {len(lote)} productos")
        if not lote or pagina >= total_paginas:
            return productos
        pagina += 1
        time.sleep(PAUSA_SEG)


def a_filas(productos):
    fecha = pd.Timestamp.now().strftime("%Y-%m-%d")
    for p in productos:
        soles = precio_soles(p)
        dolares = round(soles / TIPO_CAMBIO, 2) if soles else None
        yield {
            "DESCRIPCIÓN": limpiar(p["name"]),
            "MARCA": detectar_marca(p),
            "SKU": p.get("sku", ""),
            "CATEGORÍA": " / ".join(limpiar(c["name"]) for c in p.get("categories", [])),
            "Cant": 1,
            "U/M": "UND",
            "PRECIO SOLES": soles,
            "PRECIO DOLARES": dolares,
            "Precio Dolares Unit": dolares,
            "Precio unit. a Cotizar": soles,
            "EN STOCK": "Sí" if p.get("is_in_stock") else "No",
            "TIPO EXISTENCIA": "Mercadería",
            "PROVEEDOR": "PROMICON PERU",
            "FACTURA": "",
            "RUC": "",
            "Fecha": fecha,
            "URL": p.get("permalink", ""),
        }


def main():
    productos = descargar_catalogo(crear_sesion())
    df = pd.DataFrame(a_filas(productos)).drop_duplicates(subset=["URL"])
    df.to_excel(SALIDA, index=False)
    con_precio = df["PRECIO SOLES"].notna().sum()
    print(f"Listo: {len(df)} productos ({con_precio} con precio, {len(df) - con_precio} 'consultar') -> {SALIDA}")


if __name__ == "__main__":
    main()
