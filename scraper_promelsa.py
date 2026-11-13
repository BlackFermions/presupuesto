import time

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# API pública de la tienda (Magento), la misma que consulta el sitio: devuelve JSON, no HTML
API_URL = "https://www.promelsa.com.pe/graphql"
WEB_URL = "https://www.promelsa.com.pe/"
CATEGORIA_RAIZ = "Mw=="   # categoría "PROMELSA", contiene todo el catálogo
POR_PAGINA = 100
TOPE_API = 500            # la API nunca entrega más de 500 productos por consulta
TIPO_CAMBIO = 3.75
PAUSA_SEG = 1.0           # pausa entre consultas para no saturar el servidor
SALIDA = "Catalogo_Promelsa.xlsx"

# Son agrupaciones comerciales, no tipos de producto: se recorren al final
# para que la columna CATEGORÍA muestre la categoría real del producto
AGRUPACIONES = ["Marcas", "Destacados", "Promociones", "Outlet", "Portafolio Sostenible", "FISE 2026"]

CONSULTA_PRODUCTOS = """
query ($categoria: String!, $pagina: Int!, $por_pagina: Int!) {
  products(filter: {category_uid: {eq: $categoria}}, pageSize: $por_pagina,
           currentPage: $pagina, sort: {name: ASC}) {
    total_count
    page_info { total_pages }
    items {
      sku
      name
      url_key
      url_suffix
      stock_status
      marca
      price_range { minimum_price { final_price { value } } }
    }
  }
}
"""

CONSULTA_SUBCATEGORIAS = """
query ($categoria: String!) {
  categories(filters: {parent_category_uid: {eq: $categoria}}, pageSize: 300) {
    items { uid name }
  }
}
"""

CONSULTA_MARCAS = """
{
  customAttributeMetadata(attributes: [{attribute_code: "marca", entity_type: "catalog_product"}]) {
    items { attribute_options { label value } }
  }
}
"""

CONSULTA_TOTAL = """
query ($categoria: String!) {
  categoryList(filters: {category_uid: {eq: $categoria}}) { product_count }
}
"""


def crear_sesion():
    sesion = requests.Session()
    sesion.headers["User-Agent"] = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    # Por defecto no se reintentan los POST; aquí es seguro porque las consultas solo leen datos
    reintentos = Retry(total=5, backoff_factor=2, status_forcelist=[429, 500, 502, 503, 504],
                       allowed_methods=["POST"])
    sesion.mount("https://", HTTPAdapter(max_retries=reintentos))
    return sesion


def consultar(sesion, consulta, **variables):
    """Envía una consulta a la API y devuelve sus datos. Si la API reporta error, se detiene."""
    time.sleep(PAUSA_SEG)
    resp = sesion.post(API_URL, json={"query": consulta, "variables": variables}, timeout=90)
    resp.raise_for_status()
    respuesta = resp.json()
    if respuesta.get("errors"):
        raise RuntimeError(f"Error de la API: {respuesta['errors']}")
    return respuesta["data"]


def obtener_marcas(sesion):
    """La API da la marca como un número (82); esto arma el diccionario {82: '3M'}."""
    datos = consultar(sesion, CONSULTA_MARCAS)
    opciones = datos["customAttributeMetadata"]["items"][0]["attribute_options"]
    return {int(o["value"]): o["label"] for o in opciones}


def obtener_subcategorias(sesion, categoria):
    subcategorias = consultar(sesion, CONSULTA_SUBCATEGORIAS, categoria=categoria)["categories"]["items"]
    return sorted(subcategorias, key=lambda c: c["name"] in AGRUPACIONES)


def recorrer_categoria(sesion, categoria, ruta, productos):
    """Guarda en `productos` lo que haya en la categoría, sin repetir SKU.

    Si tiene más productos que el tope de la API, baja a sus subcategorías.
    """
    pagina = consultar(sesion, CONSULTA_PRODUCTOS, categoria=categoria, pagina=1, por_pagina=POR_PAGINA)["products"]

    if pagina["total_count"] >= TOPE_API:
        subcategorias = obtener_subcategorias(sesion, categoria)
        if subcategorias:
            for sub in subcategorias:
                recorrer_categoria(sesion, sub["uid"], ruta + [sub["name"]], productos)
            return
        print(f"  AVISO: '{' / '.join(ruta)}' tiene más de {TOPE_API} y no se puede dividir: puede quedar incompleta")

    total_paginas = pagina["page_info"]["total_pages"]
    for numero in range(1, total_paginas + 1):
        if numero > 1:
            pagina = consultar(sesion, CONSULTA_PRODUCTOS, categoria=categoria, pagina=numero,
                               por_pagina=POR_PAGINA)["products"]
        for producto in pagina["items"]:
            productos.setdefault(producto["sku"], {**producto, "categoria": " / ".join(ruta)})

    if total_paginas:
        print(f"  {' / '.join(ruta)}: {pagina['total_count']} productos | acumulado {len(productos)}")


def a_filas(productos, marcas):
    fecha = pd.Timestamp.now().strftime("%Y-%m-%d")
    for p in productos:
        # Precio 0 en la tienda significa "Solicitar Cotización": queda vacío, no en 0
        soles = p["price_range"]["minimum_price"]["final_price"]["value"] or None
        dolares = round(soles / TIPO_CAMBIO, 2) if soles else None
        yield {
            "DESCRIPCIÓN": p["name"],
            "MARCA": marcas.get(p["marca"], "OTRO"),
            "SKU": p["sku"],
            "CATEGORÍA": p["categoria"],
            "Cant": 1,
            "U/M": "UND",
            "PRECIO SOLES": soles,
            "PRECIO DOLARES": dolares,
            "Precio Dolares Unit": dolares,
            "Precio unit. a Cotizar": soles,
            "EN STOCK": "Sí" if p["stock_status"] == "IN_STOCK" else "No",
            "TIPO EXISTENCIA": "Mercadería",
            "PROVEEDOR": "PROMELSA",
            "FACTURA": "",
            "RUC": "",
            "Fecha": fecha,
            "URL": f"{WEB_URL}{p['url_key']}{p['url_suffix'] or ''}",
        }


def main():
    sesion = crear_sesion()
    marcas = obtener_marcas(sesion)
    total_web = consultar(sesion, CONSULTA_TOTAL, categoria=CATEGORIA_RAIZ)["categoryList"][0]["product_count"]
    print(f"Productos en la tienda: {total_web} | marcas: {len(marcas)}")

    productos = {}
    recorrer_categoria(sesion, CATEGORIA_RAIZ, [], productos)

    df = pd.DataFrame(a_filas(productos.values(), marcas))
    df.to_excel(SALIDA, index=False)
    con_precio = df["PRECIO SOLES"].notna().sum()
    print(f"Listo: {len(df)} de {total_web} productos "
          f"({con_precio} con precio, {len(df) - con_precio} 'Solicitar Cotización') -> {SALIDA}")


if __name__ == "__main__":
    main()
