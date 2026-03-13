import json
import pandas as pd
import scrapy
import os
from datetime import datetime

SPECIFI_RENAME_MAP = {
    "w": "Energía",
    "lm": "Flujo luminoso",
    "medida": "Dimensiones",
    "angulo": "Ángulo",
    "qtcx": "Cantidad/caja",
    "k": "Temperatura de color",
    "ip": "Protección IP",
    "ik": "Protección IK",
    "mat": "Material",
    "corte": "Corte",
    "cri": "CRI",
    "volt": "Tensión",
    "motor": "Motor de CC",
    "led": "LED",
    "furo": "Agujero",
    "rolo": "Rollo",
    "tempe": "Temperatura",
    "ac": "AC",
    "letras": "Letras",
    "pcb": "PCB",
    "poten_max": "Potencia Máxima",
    "poten_nom": "Potencia Nominal",
    "corrente_ac": "Corriente CA",
    "ampere": "Amperios",
    "corrente_carga": "Corriente Carga",
    "corrente_no": "Corriente Nominal",
    "dc": "DC"
}

SPECIFI_REMOVE_KEYS = {
    "fichaIndex",
    "corteMed",
    "por_metro",
    "cor",
    #"descri",
    "pvp_status",
    "altura",
    "diam",
    "l_W",
    "qtd_minima",
    "ref",
    "ugr",
    "precoL"
}


# -------------------------------------------------------
# UTILS
# -------------------------------------------------------

import re

def clean_illegal_chars(text):
    """Elimina caracteres ilegales que openpyxl no permite."""
    if isinstance(text, str):
        return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)
    return text


def transform_especifi_dict(d):
    """
    - Renombra claves según tu tabla
    - Elimina claves no deseadas
    - Extrae pvp para procesarlo fuera
    """

    cleaned = {}

    for key, value in d.items():
        if value in (None, "", "null"):
            continue

        if key == "pvp":
            # Se ignora aquí → se gestiona fuera
            continue

        if key in SPECIFI_REMOVE_KEYS:
            continue

        new_key = SPECIFI_RENAME_MAP.get(key, key)

        cleaned[new_key] = value

    return cleaned


def slugify(text):
    text = text.lower().strip()
    text = re.sub(r"[^\w\s-]", "", text)
    text = text.replace(" ", "-")
    text = re.sub(r"-+", "-", text)
    return text


def normalize_especifi(especifi_list, do_variantes):
    """
    Normaliza el bloque 'especifi':
    1) crea un bloque común
    2) separa variantes
    3) elimina claves con None
    """

    def clean_dict(d):
        return {k: v for k, v in d.items() if v not in (None, "", "null")}

    if not especifi_list or not isinstance(especifi_list, list):
        return especifi_list

    # 1) Base común = copy del primer elemento
    base = especifi_list[0].copy()

    # Claves que deben excluirse del bloque común
    remove_common_keys = {"ref", "k"}

    # Limpiar nulls del bloque común
    base = clean_dict(base)
    variants = []
    if do_variantes:
        for key in remove_common_keys:
            base.pop(key, None)

        # 2) Variantes → solo ref y clave distintiva (k)
        for item in especifi_list:
            variant = {}
            for key in ["ref", "k"]:  # aquí puedes añadir más si cambian
                if key in item and item[key] not in (None, "", "null"):
                    variant[key] = item[key]
            if variant:
                variants.append(variant)

    # 3) Construir output final
    final = [base]
    if do_variantes: final.extend(variants)

    return final


# -------------------------------------------------------
# INIT
# -------------------------------------------------------

def init(spider):
    spider.allowed_domains = ["backend-theta-gold-17.vercel.app"]
    spider.start_urls = ["https://backend-theta-gold-17.vercel.app/produtos"]

    spider.advance_products = []

    print("🔧 Handler advance inicializado correctamente")


# -------------------------------------------------------
# START
# -------------------------------------------------------

async def start(spider):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/123.0.0.0 Safari/537.36"
        )
    }

    for url in spider.start_urls:
        yield scrapy.Request(url, callback=spider.parse, headers=headers, dont_filter=True)


# -------------------------------------------------------
# PARSE LISTA IDP
# -------------------------------------------------------

def parse(spider, response):
    data = json.loads(response.text)

    for item in data:
        idP = item.get("idP")
        if not idP:
            continue

        url = f"https://backend-theta-gold-17.vercel.app/api/productFull/{idP}"

        yield scrapy.Request(
            url,
            callback=spider.parse_product,
            meta={"idP": idP}
        )


# -------------------------------------------------------
# PARSE PRODUCTO
# -------------------------------------------------------

def parse_product(spider, response):
    data = json.loads(response.text)

    # 1) Filtrar NULLS
    clean = {k: v for k, v in data.items() if v is not None}

    # 2) Campos que NO queremos guardar
    campos_excluir = [
        "oculto",
        "novo",
        "em_campanha",
        "campanha_atual",
        "id_Mar",
        "id_Cat",
        "Tipo",
        "acessorios",
        "data_insercao"
    ]

    for campo in campos_excluir:
        clean.pop(campo, None)

    # -------------------------------------------------------
    # 3) Agrupar documentos
    # -------------------------------------------------------
    documentos = {}

    if clean.get("ficha_tecnica"):
        documentos["Ficha técnica"] = clean["ficha_tecnica"]

    if clean.get("cert_confor"):
        documentos["Certificado de cumplimiento"] = clean["cert_confor"]

    # si hay documentos → añadir en el JSON final
    if documentos:
        clean["Documentos"] = documentos

    # eliminar originales
    clean.pop("ficha_tecnica", None)
    clean.pop("cert_confor", None)
    clean.pop("ficha_tecnica_2", None)
    clean.pop("ficha_tecnica_3", None)

    # -------------------------------------------------------
    # 4) Normalizar variantes (y añadir ip/tempe)
    # -------------------------------------------------------
    especifi = clean.get("especifi")
    skus = clean.get("referencias")
    do_variantes = True if len(skus)>1 else False
    if not do_variantes:
        clean["referencias"] = clean["referencias"][0]

    if especifi:
        especifi_norm = normalize_especifi(especifi, do_variantes)

        # Añadimos ip y tempe al primer bloque común
        if isinstance(especifi_norm, list) and len(especifi_norm) > 0:

            for key in ["ip", "tempe", "volt", "ies", "cri", "dc", "ac", "ik", "ampere", "led", "ugr", "dt", "ip", "tempe", "altura", "diam", "pcb"]:
                if key in clean and clean[key] not in (None, "", "null"):
                    especifi_norm[0][key] = clean[key]
                    clean.pop(key, None)

        clean["especifi"] = especifi_norm
        # Transformar campos de especifi (renombrar y filtrar)
        final_especifi = []
        coste = None
        descWeb = None

        for item in clean["especifi"]:
            # Extraer Coste desde pvp si existe
            if "pvp" in item and item["pvp"] not in (None, "", "null"):
                coste = item["pvp"]
            elif "descri" in item and item["descri"] not in (None, "", "null"):
                descWeb = item["descri"]

            transformed = transform_especifi_dict(item)
            if transformed:
                final_especifi.append(transformed)

        clean["especifi"] = [
            {k: clean_illegal_chars(v) for k, v in item.items()}
            for item in final_especifi
        ]

        # Añadir Coste fuera en su propia columna
        if coste is not None:
            clean["Coste"] = coste
        if descWeb is not None:
            clean["Descripción Web"] = descWeb

    # -------------------------------------------------------
    # 5) URL SEO-friendly
    # -------------------------------------------------------
    nome = clean.get("nome", "")
    idP = clean.get("idP", "")

    slug = slugify(nome)
    clean["url"] = f"https://advance-rr.com/prodetail/{slug}/{idP}"

    # Guardar
    safe_clean = {
        k: clean_illegal_chars(v)
        for k, v in clean.items()
    }

    spider.advance_products.append(safe_clean)


# -------------------------------------------------------
# CLOSED → EXPORTA EXCEL
# -------------------------------------------------------

def closed(spider, reason):
    print("📦 Finalizando scrapeo advance...")

    filename = "advance_products"
    filename += "-" + datetime.now().strftime('%d-%m-%Y') + ".xlsx"

    ruta = os.path.expanduser("~/Documents/SMI Files")
    os.makedirs(ruta, exist_ok=True)
    ruta += "/" + filename

    if not spider.advance_products:
        print("⚠️ No se recogieron productos.")
        return

    final_products = []

    for prod in spider.advance_products:
        p = prod.copy()

        # ❗ Si especifi es una lista con 1 solo elemento → convertir a dict
        if isinstance(p.get("especifi"), list) and len(p["especifi"]) == 1:
            p["especifi"] = p["especifi"][0]

        final_products.append(p)

    df = pd.DataFrame(final_products)
    df.to_excel(ruta, index=False)

    print(f"Archivo generado: {filename}")
