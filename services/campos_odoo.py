# campos_odoo.py
from enum import Enum


class ClavesExcel(Enum):
    URL = "URL de orígen"
    CODIGO_FAMILIA = "Código de familia"
    STOCK_BULGARIA = "Stock Bulgaria"
    STOCK_UK = "Stock UK"
    STOCK_ITA = "Stock ITA"
    STOCK_TRANSITO = "Stock en tránsito (buyled + madrid)"
    TRANSITO = "En tránsito a UK/Bulgaria"
    IMAGEN = "Imagen principal"
    NOMBRE = "Nombre"
    DESCRIPCION_WEB = "Descripción Web"
    DESCRIPCION = "Descripción"
    DESCRIPCION_COMPRA = "Descripción Compra"
    DESCRIPCION_VENTAS = "Descripción Ventas"
    DESCRIPCION_SIN_STOCK = "Descripción: No hay Stock"
    CATEGORIA = "Categoría"
    PRECIO = "Precio"
    COSTE = "Coste"
    VOLUMEN = "Volumen"
    PESO = "Peso"
    REFERENCIA = "EAN"
    SKU = "SKU"
    MARCA = "Marca"
    ATRIBUTOS = "Atributos"
    GALERIA = "Galería"
    VIDEOS = "Vídeos"
    DOCUMENTOS = "Documentos"


class _ValoresImport(Enum):  # ← uso interno
    URL = "x_url"
    CODIGO_FAMILIA = "x_codigo_de_familia"
    STOCK_BULGARIA = "x_almacen1_custom"
    STOCK_UK = "x_almacen2_custom"
    STOCK_ITA = "x_almacen3_custom"
    STOCK_TRANSITO = "x_transit_stock_custom"
    TRANSITO = "x_transit"
    IMAGEN = "image_1920"
    NOMBRE = "name"
    DESCRIPCION_WEB = "website_description"
    DESCRIPCION = "description"
    DESCRIPCION_COMPRA = "description_purchase"
    DESCRIPCION_VENTAS = "description_sale"
    DESCRIPCION_SIN_STOCK = "out_of_stock_message"
    CATEGORIA = "public_categ_ids"
    PRECIO = "list_price"
    COSTE = "standard_price"
    VOLUMEN = "volume"
    PESO = "weight"
    REFERENCIA = "barcode"
    SKU = "default_code"
    MARCA = "product_brand_id"
    ATRIBUTOS = "Specifications"
    GALERIA = "Image_Urls"
    VIDEOS = "Video_Urls"
    DOCUMENTOS = "Pdf_Urls"


excel_a_odoo = {
    clave.value.strip(): _ValoresImport[clave.name].value.strip()
    for clave in ClavesExcel
}

__all__ = ["ClavesExcel", "excel_a_odoo"]

FISCAL_POSITION_MAP = {
    "Régimen Nacional": "Domestic",
    "EU privado": "EU private",
    "Régimen Intracomunitario": "Intra-community",
    "Régimen Extracomunitario": "Extra-community",
    "Régimen No sujeto por reglas de localización": "Regime not subject to localization rules",
    "Recargo de Equivalencia": "Equivalence surcharge",
    "Recargo de Equivalencia Revendedor con ISP": "ISP reseller equivalence surcharge",

    "Retención IRPF 1%": "Personal income tax withholding 1%",
    "Retención IRPF 2%": "Personal income tax withholding 2%",
    "Retención IRPF 7%": "Personal income tax withholding 7%",
    "Retención IRPF 9%": "Perosnal income tax withholding 9%",
    "Retención IRPF 15%": "Personal income tax withholding 15%",
    "Retención IRPF 18%": "Personal income tax withholding 18%",
    "Retención IRPF 19%": "Personal income tax withholding 19%",
    "Retención IRPF 20%": "Personal income tax withholding 20%",
    "Retención IRPF 21%": "Personal tax income withholding 21%",
    "Retención IRPF 24%": "Personal income tax withholding 24%",

    "Retención 19% arrendamientos": "Withholding 19% leases",
    "Retención 19,5% arrendamientos": "Withholding 19.5% leases",
    "Retención 20% arrendamientos": "Retention 19% leases",  # revisar si es correcto en tu caso
    "Retención 21% arrendamientos": "Withholding 21% leases",

    "Inversion del Sujeto Pasivo Nacional": "National Reverse charge",

    "REAGYP - Agricultura": "REAGYP - Agriculture",
    "REAGYP - Ganadería y pesca": "REAGYP - animal breeding and fishing",

    "DUA": "DUA",

    "Retención IRPF No residentes UE 24%": "24% Withholding for non-EU residents",
    "Retención IRPF No residentes UE 19%": "19% Withholding for non-EU residents",
    "Retención IRPF residentes UE exento por convenio": "Exempt Withholding EU residents",
    "Retención IRPF No residentes UE exento por convenio": "Exempt Withholding Non EU residents",
}
