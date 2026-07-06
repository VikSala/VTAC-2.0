import scrapy
from urllib.parse import urljoin
import re
from collections import defaultdict

def init(spider):
    spider.allowed_domains = ["v-tac.es"]
    spider.start_urls = ['https://v-tac.es/nuevos-productos', 'https://v-tac.es/descatalogados.html']
    spider.CATEGORIES_LINKS = (
        'https://v-tac.es/sistemas-solares',
        'https://v-tac.es/iluminaci%C3%B3n.html',
        'https://v-tac.es/smart-digital.html',
        'https://v-tac.es/el%C3%A9ctrico.html'
    )
    spider.start_urls = list(spider.CATEGORIES_LINKS)#comentar si prueba pequeña
    spider.product_map = defaultdict(dict)

def parse(spider, response):
    # Si el GUI/CLI no pasa el flag, mantenemos el comportamiento por defecto
    comercial_scrap = getattr(spider, "comercial_scrap", None)
    if comercial_scrap is None:
        comercial_scrap = False

    if comercial_scrap:
        comercial_links = (
            'https://v-tac.es/nuevos-productos',
            'https://v-tac.es/descatalogados.html'
        )
        for comercial in comercial_links:
            full_url = urljoin(response.url, comercial + "?limit=150&start=0")
            yield scrapy.Request(full_url, callback=spider.parse_category)
    else:
        # Extraer enlaces tipo shoplink
        shoplinks = response.css(
            "section#sp-main-body div.sppb-addon-content h4 a::attr(href)").getall()  # Esto es lo que evita que se vaya al menu: section#sp-main-body
        for shop in shoplinks:
            full_url = urljoin(response.url, shop + "?limit=150&start=0")
            yield scrapy.Request(full_url, callback=spider.parse_category)

        if response.url in spider.CATEGORIES_LINKS and "sistemas-solares" in response.url:
            descatalogados_url = "https://v-tac.es/descatalogados.html?limit=150&start=0"
            yield scrapy.Request(descatalogados_url, callback=spider.parse_category)

def parse_category(spider, response):
    for product in response.css("div.product"):
        url = "https://v-tac.es" + product.css("a::attr(href)").get()
        name = product.css("a::attr(title)").get()
        image = "https://v-tac.es" + product.css("img::attr(src)").get()
        sku_raw = product.css("img::attr(alt)").get()
        sku = re.sub(r"[^\d]", "", sku_raw) if sku_raw else ""

        if url:
            yield response.follow(url, callback=spider.parse_product, meta={
                "x_url_origen": urljoin(response.url, url),
                "Name": name,
                "Image": urljoin(response.url, image) if image else "",
                "default_code": sku,
            })

    if "No hay productos" not in response.text and "product" in response.text and response.css(
            'a[title="Siguiente"]'):
        current = int(response.url.split("start=")[-1])
        next_page = response.url.replace(f"start={current}", f"start={current + 150}")
        yield scrapy.Request(next_page, callback=spider.parse_category)

def parse_product(spider, response):
    from urllib.parse import urlparse, unquote
    # Extraer las categorías a partir de la URL del producto
    cat_name = ""
    path_parts = urlparse(response.url).path.strip("/").split("/")
    if len(path_parts) > 1:
        cat_parts = path_parts[:-1]  # Excluir el último segmento (el nombre del producto)
        cat_name = " / ".join([unquote(p).replace("-", " ").title() for p in cat_parts])

    item = {
        "x_url_origen": response.meta["x_url_origen"],
        "Name": response.meta["Name"],
        "Image": response.meta["Image"],
        "Image_Urls": [],
        "website_description": "",
        "Video_Urls": [],
        "Pdf_Urls": {},
        "Specifications": {},
        "default_code": response.meta["default_code"],
        "barcode": "",
        "weight": "",
        "x_marca": "",
        "x_categoria": [cat_name] if cat_name else [],
    }

    raw_sku = response.css('div.sku-inner::text').get()

    if raw_sku:
        sku = raw_sku.strip().split('|')[0]
        item["default_code"] = sku


    # Galería de imágenes
    gallery_imgs = response.css(".additional-images a::attr(href)").getall()
    item["Image_Urls"] = [urljoin(response.url, src) for src in gallery_imgs]

    # Descripción (sin bloque “Contáctenos”)
    description = response.css("div.product-description").get()
    video_urls = []
    if description:
        description = re.sub(
            r'<div>\s*<a[^>]*cont(?:%C3%A1|á)ctenos[^>]*>Contáctenos<\/a>\s*<\/div>',
            '',
            description,
            flags=re.IGNORECASE
        )
        # Buscar enlaces de YouTube
        video_urls = re.findall(r'https://(?:www\.)?youtube\.com/watch\?v=[\w-]+', description)
    item["website_description"] = description
    item["Video_Urls"] = video_urls

    # PDFs
    for link in response.css("div.downloads div.product-field"):
        name_parts = link.css(".product-fields-title *::text").getall()
        name = " ".join([n.strip() for n in name_parts if n.strip()])

        href = link.css(".product-field-display a::attr(href)").get()
        if name and href:
            item["Pdf_Urls"][name] = urljoin(response.url, href)

    # Especificaciones
    specs = {}
    for row in response.xpath(
        "//div[contains(@class,'product-field')]"                             
        "[not(contains(@class,'product-field-type-G'))]"                      
        "[not(ancestor::div[contains(@class,'downloads')])]"                  
        "[div[contains(@class,'product-field-display')]]"
    ):
        key_parts = row.css(".product-fields-title *::text").getall()
        key = " ".join([k.strip() for k in key_parts if k.strip()])

        val_parts = row.css(".product-field-display *::text").getall()
        val_text = " ".join([v.strip() for v in val_parts if v.strip()])

        link = row.css(".product-field-display a::attr(href)").get()
        if link:
            val = response.urljoin(link)
        elif row.css(".product-field-display img::attr(src)").get():
            val = response.urljoin(row.css(".product-field-display img::attr(src)").get())
        else:
            val = val_text

        if key and val:
            specs[key] = val

            if key.lower() == "ean código":
                item["barcode"] = str(val)
            elif key.lower() == "peso del artículo":
                item["weight"] = val
            elif key.lower() == "marca":
                item["x_marca"] = val

    item["Specifications"] = specs
    spider.logger.info(f"Producto procesado: {response.meta['Name']}")

    key = item["default_code"] if item["default_code"] else item["Name"]

    if key in spider.product_map:
        existing_item = spider.product_map[key]
        # Añadir la categoría actual si no estaba ya
        if cat_name and cat_name not in existing_item["x_categoria"]:
            existing_item["x_categoria"].append(cat_name)
    else:
        spider.product_map[key] = item

# GUARDAR EN EXCEL
def closed(spider, reason):
    save_to_excel(list(spider.product_map.values()))
    excel_read_and_parse()

def save_to_excel(items, filename="VTAC_ES"):
    import pandas as pd
    import os
    from datetime import datetime
    filename += "-" + datetime.now().strftime('%d-%m-%Y') + ".xlsx"

    if not items:
        print("No hay datos para guardar en Excel.")
        return

    df_main = pd.DataFrame(items)

    ruta = os.path.expanduser("~/Documents/SMI Files")
    os.makedirs(ruta, exist_ok=True)
    ruta += "/" + filename
    df_main["barcode"] = df_main["barcode"].apply(lambda x: f"'{x}")

    # Convertir listas de categorías en string plano
    if "x_categoria" in df_main.columns:
        df_main["x_categoria"] = df_main["x_categoria"].apply(lambda x: ", ".join(x) if isinstance(x, list) else x)

    with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
        df_main.to_excel(writer, sheet_name='ES', index=False)

def excel_read_and_parse(filename="VTAC_ES"):
    import pandas as pd
    import string
    import os
    from datetime import datetime
    def clean_float(value):
        import re
        try:
            if value is None or str(value).strip().lower() in ['nan', '', 'none']:
                return 0.0
            value = re.sub(r'[^\d.,-]', '', str(value).strip())
            return float(value.replace(',', '.')) if value else 0.0
        except (ValueError, TypeError):
            return 0.0
    filename += "-" + datetime.now().strftime('%d-%m-%Y') + ".xlsx"
    ruta = os.path.expanduser("~/Documents/SMI Files") + "/" + filename
    df = pd.read_excel(ruta, sheet_name="ES")
    df['weight'] = df['weight'].apply(clean_float)

    barcode_column = 'barcode'
    df[barcode_column] = df[barcode_column].astype(str)
    barcode_count = {}

    for index, row in df.iterrows():
        original_barcode = row[barcode_column]
        if original_barcode in barcode_count and original_barcode != "'":
            count = barcode_count[original_barcode]
            suffix = string.ascii_uppercase[count - 1]
            new_barcode = f"{original_barcode}{suffix}"
            while new_barcode in barcode_count:
                count += 1
                suffix = string.ascii_uppercase[count]
                new_barcode = f"{original_barcode}{suffix}"
            df.at[index, barcode_column] = new_barcode.lstrip("'")
            barcode_count[original_barcode] += 1
            barcode_count[new_barcode] = 1
        else:
            barcode_count[original_barcode] = 1
            df.at[index, barcode_column] = original_barcode.lstrip("'")

    # Renombrar columnas del excel
    column_renames = {
        "x_url_origen": "URL de orígen",
        "Name": "Nombre",
        "Image": "Imagen principal",
        "Image_Urls": "Galería",
        "website_description": "Descripción Web",
        "Video_Urls": "Vídeos",
        "Pdf_Urls": "Documentos",
        "Specifications": "Atributos",
        "default_code": "SKU",
        "barcode": "EAN",
        "weight": "Peso",
        "x_marca": "Marca",
        "x_categoria": "Categoría",
    }

    # Aplicar el renombrado al DataFrame
    df.rename(columns=column_renames, inplace=True)

    # Guardamos resultados
    df.to_excel(ruta, sheet_name="ES", index=False)
    merge_path = os.path.expanduser("~/Documents/SMI Files/Merge")
    os.makedirs(merge_path, exist_ok=True)
    excel_path = merge_path + "/scraped.xlsx"

    if not os.path.exists(excel_path):
        df.to_excel(excel_path, index=False, sheet_name="ES")
    else:
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, index=False, sheet_name="ES")
    normalizar_atributos_excel(excel_path, "es", excel_path)
    print(f"✅ El archivo Excel está listo: {filename}")

def normalizar_atributos_excel(excel_path: str, region: str, output_path: str = None):
    """
    Normaliza la columna 'Atributos' de la hoja 'Products' según la región.
    - Si region != 'es': reemplaza las keys usando MAP_REGION y elimina las que no estén en el mapping.
    - Si region == 'es': mantiene todas las claves (sólo renombra si hay equivalencia).

    :param excel_path: ruta del Excel original
    :param region: 'es', 'it', 'uk', etc.
    :param output_path: ruta opcional para guardar un nuevo Excel
    :return: DataFrame modificado
    """
    import pandas as pd
    import ast

    MAP_ES = {
        "Adicional": "Comentarios",
        "Cantidad de LED": "Cantidad del LED",
        "Código de orden": "SKU",
        "Dimensiones de montaje": "Dimensiones",
        "Jacket Dia": "Jacket Diameter",
        "Modo de carga corriente de entrada": "",
        "Modo de emergencia corriente de salida": "",
        "Product Features": "",
        "Protección IP": "Grado de protección IP",
        "Tipo casquillo": "Casquillo",
        "Ángulo de haz": "Ángulo de Apertura",
        "CRI 90+": "CRI"
    }
    df = pd.read_excel(excel_path, sheet_name=region.upper())

    if "Atributos" not in df.columns:
        raise ValueError("❌ La hoja 'Products' no contiene la columna 'Atributos'.")

    region = region.lower()
    mapping = MAP_ES
    if mapping is None:
        raise ValueError(f"❌ No existe un mapping para la región: {region}")

    def normalize_dict(attr_str):
        try:
            attrs = ast.literal_eval(attr_str) if isinstance(attr_str, str) else {}
            normalized = {}
            for k, v in attrs.items():
                std_key = mapping.get(k)
                if std_key:  # hay equivalencia → usar la clave española
                    normalized[std_key] = v
                else:
                    if region == "es":
                        # España mantiene aunque no esté en el mapping
                        normalized[k] = v
                    # si region != "es" y no hay equivalencia → se descarta
            return normalized
        except Exception:
            return {}

    # Normalizar columna
    df["Atributos"] = df["Atributos"].apply(normalize_dict)

    # Guardar si se pide
    if output_path:
        with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, index=False, sheet_name=region.upper())
        print(f"✅ Archivo normalizado guardado en: {output_path}")
