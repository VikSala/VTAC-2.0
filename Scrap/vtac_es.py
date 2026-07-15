import os
import re
import time
import json
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from services.utils import Utils
from enum import Enum
from services.campos_odoo import ClavesExcel
from urllib.parse import urljoin, urlparse, unquote, parse_qsl, urlencode, urlunparse


class ScrapMode(Enum):
    TEST = "Test"
    MINI = "Mini"
    COMPLETO = "Completo"


# ✨ MAIN FUNCION PUBLICA
def run_vtac_es_scraper(modo_scrap: ScrapMode = ScrapMode.COMPLETO, only_new_products=False,
                        excel_output="vtac_es.xlsx", progress_callback=None):
    file_path = ""
    if only_new_products:
        from tkinter import Tk, filedialog
        Tk().withdraw()
        file_path = filedialog.askopenfilename(
            title="Selecciona el Excel de productos anteriores",
            filetypes=[("Excel Files", "*.xlsx")]
        )

    # Función helper para emitir progreso
    def emit_progress(message):
        if progress_callback:
            progress_callback(message)
        print(message)

    emit_progress(f"MODO {modo_scrap.value} ES ACTIVADO")

    driver = _init_driver()
    product_links = _get_product_links(
        driver,
        modo_scrap,
        only_new_products,
        file_path,
        emit_progress
    )
    productos = _scrape_productos(driver, product_links, emit_progress)

    if productos:
        _guardar_excel(productos, excel_output)
        emit_progress(f"ESPAÑA: 100%. Scraping guardado en {excel_output}")
    else:
        emit_progress("No se han scrapeado productos...")


# ⚙️ Inicializa el navegador con el perfil configurado
def _init_driver():
    profile_path = os.path.expanduser(
        "~/AppData/Local/Google/Chrome/User Data/Profile_4_scrapy"
    )
    options = uc.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument(f"--user-data-dir={profile_path}")
    options.add_argument("--profile-directory=Default")
    return uc.Chrome(options=options)


def _reiniciar_driver(driver):
    try:
        driver.quit()
    except:
        pass

    driver = _init_driver()

    time.sleep(3)

    return driver


# 🔍 Obtiene las URLs de productos según el modo seleccionado
def _get_product_links(driver, modo_scrap, only_new_products, file_path="", progress_callback=None):
    def emit_progress(message):
        if progress_callback:
            progress_callback(message)
        print(message)

    product_links = []
    subcategorias = {}
    categorias_principales = [
        "https://v-tac.es/sistemas-solares",
        "https://v-tac.es/iluminaci%C3%B3n.html",
        "https://v-tac.es/smart-digital.html",
        "https://v-tac.es/el%C3%A9ctrico.html"
    ]

    comercial_scrap = True

    SCRAP_SELECTIVO = False

    if not SCRAP_SELECTIVO:

        # comercial_scrap tiene prioridad incluso sobre el modo TEST.
        if comercial_scrap:
            emit_progress("Explorando nuevos productos y descatalogados...")
            subcategorias = {
                "Nuevos productos": "https://v-tac.es/nuevos-productos",
                "Descatalogados": "https://v-tac.es/descatalogados.html",
            }

        elif modo_scrap == ScrapMode.TEST:
            emit_progress("MODO TEST ACTIVADO")
            product_links = [
                "https://v-tac.es/iluminaci%C3%B3n/oficina-y-comercio/focos-de-carril-magn%C3%A9tico/10w-48v-foco-de-carril-colgante-4000k-cuerpo-negro-detail.html",
                "https://v-tac.es/iluminaci%C3%B3n/casa-y-jard%C3%ADn/bombillas-led/bombilla-led-12-5w-filamento-e27-g125-ambar-2200k-detail.html"
            ]

        else:
            emit_progress("Explorando categorías principales...")

            for i, url_categoria in enumerate(categorias_principales, 1):
                emit_progress(
                    f"📁 Buscando subcategorías ({i}/{len(categorias_principales)}): "
                    f"{url_categoria}"
                )
                driver.get(url_categoria)
                time.sleep(2)

                try:
                    enlaces = driver.find_elements(
                        By.CSS_SELECTOR,
                        "section#sp-main-body div.sppb-addon-content h4 a[href]"
                    )

                    for a in enlaces:
                        href = a.get_attribute("href")
                        nombre = a.text.strip() or href

                        if href and nombre:
                            subcategorias[nombre] = urljoin(url_categoria, href)

                except Exception as e:
                    emit_progress(
                        f"❌ No se encontraron subcategorías en {url_categoria}: {e}"
                    )

                # Se conserva el comportamiento del Scrapy español original.
                if "sistemas-solares" in url_categoria:
                    subcategorias["Descatalogados"] = (
                        "https://v-tac.es/descatalogados.html"
                    )

        # Se recorren las URLs comerciales o las subcategorías normales.
        if subcategorias:
            for idx, (nombre_cat, shop_link) in enumerate(subcategorias.items(), 1):
                emit_progress(
                    f"🔍 Procesando subcategoría "
                    f"({idx}/{len(subcategorias)}): {nombre_cat}"
                )

                parsed = urlparse(shop_link)
                query = dict(parse_qsl(parsed.query, keep_blank_values=True))
                query["limit"] = "150"
                query["start"] = "0"
                current_url = urlunparse(
                    parsed._replace(query=urlencode(query))
                )

                visited_pages = set()
                page_idx = 0

                while current_url and current_url not in visited_pages:
                    visited_pages.add(current_url)
                    page_idx += 1

                    if page_idx > 1:
                        emit_progress(f"  📄 Página {page_idx}")

                    driver.get(current_url)
                    time.sleep(3)

                    product_cards = driver.find_elements(
                        By.CSS_SELECTOR,
                        "div.product"
                    )

                    for card in product_cards:
                        try:
                            links = card.find_elements(By.CSS_SELECTOR, "a[href]")
                            if not links:
                                continue

                            href = links[0].get_attribute("href")
                            href = urljoin(current_url, href)

                            if href and href not in product_links:
                                product_links.append(href)

                            if modo_scrap == ScrapMode.MINI:
                                break
                        except Exception:
                            continue

                    if modo_scrap == ScrapMode.MINI:
                        break

                    next_buttons = driver.find_elements(
                        By.CSS_SELECTOR,
                        'a[title="Siguiente"]'
                    )

                    if next_buttons:
                        next_href = next_buttons[0].get_attribute("href")
                        current_url = (
                            urljoin(current_url, next_href)
                            if next_href else None
                        )
                    else:
                        current_url = None

                if modo_scrap == ScrapMode.MINI:
                    break

        if only_new_products:
            emit_progress("🔍 Filtrando productos ya existentes...")
            try:
                if file_path:
                    df_old = pd.read_excel(file_path)
                    urls_existentes = set(
                        df_old[ClavesExcel.URL.value]
                        .dropna()
                        .astype(str)
                        .str.strip()
                    )
                    original_count = len(product_links)
                    product_links = [
                        url for url in product_links
                        if url not in urls_existentes
                    ]
                    emit_progress(
                        f"Productos totales: {original_count}, "
                        f"Nuevos detectados: {len(product_links)}"
                    )
                    return product_links
                else:
                    emit_progress(
                        "No se seleccionó ningún archivo. "
                        "Se devolverán todos los productos."
                    )
            except Exception as e:
                emit_progress(f"Error al comparar productos previos: {e}")

        emit_progress(
            f"Total de productos para procesar: {len(product_links)}"
        )
        return product_links

    else:
        emit_progress("🚀 MODO SCRAP SELECTIVO ESPAÑA ACTIVADO")

        file_path = Utils.seleccionar_excel()
        SKU_EXCEL_PATH = file_path
        SKU_COLUMN_NAME = "SKU"

        # 1️⃣ Cargar SKUs válidos
        try:
            df_sku = pd.read_excel(
                SKU_EXCEL_PATH,
                usecols=[SKU_COLUMN_NAME]
            )

            allowed_skus = set()

            for value in df_sku[SKU_COLUMN_NAME].dropna():
                sku = str(value).strip()

                if re.fullmatch(r"\d+\.0", sku):
                    sku = sku[:-2]

                sku = re.sub(r"[^\d]", "", sku)

                if sku:
                    allowed_skus.add(sku)

            emit_progress(f"📑 SKUs cargados: {len(allowed_skus)}")
        except Exception as e:
            emit_progress(f"❌ No se pudo leer el Excel de SKUs: {e}")
            return []

        product_links = []
        subcategorias = {}

        # ── Obtener subcategorías ──────────────────────────────────

        emit_progress("Explorando categorías principales…")

        for i, url_categoria in enumerate(categorias_principales, 1):
            emit_progress(
                f"📁 Buscando subcategorías ({i}/{len(categorias_principales)}): "
                f"{url_categoria}"
            )
            driver.get(url_categoria)
            time.sleep(3)

            try:
                enlaces = driver.find_elements(
                    By.CSS_SELECTOR,
                    "section#sp-main-body div.sppb-addon-content h4 a[href]"
                )

                for a in enlaces:
                    href = a.get_attribute("href")
                    nombre = a.text.strip() or href

                    if href and nombre:
                        subcategorias[nombre] = urljoin(url_categoria, href)

            except Exception as e:
                emit_progress(
                    f"❌ No se encontraron subcategorías en {url_categoria}: {e}"
                )

            if "sistemas-solares" in url_categoria:
                subcategorias["Descatalogados"] = (
                    "https://v-tac.es/descatalogados.html"
                )

        # ── Recorrer subcategorías y páginas ───────────────────────

        for idx, (nombre_cat, shop_link) in enumerate(subcategorias.items(), 1):
            emit_progress(
                f"🔍 Procesando subcategoría "
                f"({idx}/{len(subcategorias)}): {nombre_cat}"
            )

            parsed = urlparse(shop_link)
            query = dict(parse_qsl(parsed.query, keep_blank_values=True))
            query["limit"] = "150"
            query["start"] = "0"
            current_url = urlunparse(
                parsed._replace(query=urlencode(query))
            )

            visited_pages = set()
            page_idx = 0

            while current_url and current_url not in visited_pages:
                visited_pages.add(current_url)
                page_idx += 1

                if page_idx > 1:
                    emit_progress(f"  📄 Página {page_idx}")

                driver.get(current_url)
                time.sleep(3)

                try:
                    product_cards = driver.find_elements(
                        By.CSS_SELECTOR,
                        "div.product"
                    )
                    print(f"Total: {len(product_cards)} productos")
                    encontrados_en_pagina = 0

                    for card in product_cards:

                        # 2.1 SKU
                        try:
                            image = card.find_element(By.CSS_SELECTOR, "img")
                            sku_raw = image.get_attribute("alt") or ""
                            sku = re.sub(r"[^\d]", "", sku_raw)
                        except Exception:
                            continue

                        if sku not in allowed_skus:
                            continue

                        # 2.2 Enlace
                        try:
                            link_elem = card.find_element(
                                By.CSS_SELECTOR,
                                "a[href]"
                            )
                            href = urljoin(
                                current_url,
                                link_elem.get_attribute("href")
                            )
                        except Exception:
                            continue

                        if href and href not in product_links:
                            product_links.append(href)
                            encontrados_en_pagina += 1

                    if encontrados_en_pagina:
                        emit_progress(
                            f"  ✅ Añadidos {encontrados_en_pagina} "
                            "productos en esta página"
                        )

                except Exception as e:
                    emit_progress(f"  ❌ Error leyendo productos: {e}")

                next_buttons = driver.find_elements(
                    By.CSS_SELECTOR,
                    'a[title="Siguiente"]'
                )

                if next_buttons:
                    next_href = next_buttons[0].get_attribute("href")
                    current_url = (
                        urljoin(current_url, next_href)
                        if next_href else None
                    )
                else:
                    current_url = None

        emit_progress(
            f"🎯 Productos encontrados tras filtrado por SKU: "
            f"{len(product_links)}"
        )
        return product_links


# 👁️ Scrapea la información de cada producto desde su URL
def _scrape_productos(driver, product_links, progress_callback=None):
    def emit_progress(message):
        if progress_callback:
            progress_callback(message)
        print(message)

    productos = []
    productos_por_clave = {}
    productos_totales = len(product_links)
    producto_procesado = 0

    emit_progress(
        f"🚀 Iniciando scraping de {productos_totales} productos..."
    )

    for i, url in enumerate(product_links, 1):

        if i > 1 and i % 300 == 0:
            emit_progress(
                f"♻️ Reiniciando Chrome ({i}/{productos_totales})"
            )
            driver = _reiniciar_driver(driver)

        # Mostrar progreso cada 10 productos o en hitos importantes
        if i % 10 == 0 or i == 1:
            porcentaje = int((i / productos_totales) * 100)
            emit_progress(
                f"ESPAÑA: {porcentaje}% ({i}/{productos_totales}) "
                "productos procesados..."
            )

        driver.get(url)
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);"
        )
        driver.execute_script("document.body.style.zoom='25%'")
        time.sleep(2)

        # Título
        titulo = "Título no encontrado"

        for selector in (
            "h3",
            "div.product-title h1",
            ".productdetails-view h1"
        ):
            try:
                element = driver.find_element(By.CSS_SELECTOR, selector)
                if element.text.strip():
                    titulo = element.text.strip()
                    break
            except Exception:
                continue

        # SKU
        sku = ""
        try:
            raw_sku = driver.find_element(
                By.CSS_SELECTOR,
                "div.sku-inner"
            ).text.strip()

            if raw_sku:
                sku_only = raw_sku.split("|")[0].strip()
                sku = sku_only.removeprefix("SKU: ")
        except:
            pass

        # Imágenes
        imagenes = []

        try:
            gallery_links = driver.find_elements(
                By.CSS_SELECTOR,
                ".additional-images a[href]"
            )

            for link in gallery_links:
                href = link.get_attribute("href")
                href = urljoin(url, href)

                if href and href not in imagenes:
                    imagenes.append(href)
        except:
            pass

        if not imagenes:
            for selector in (
                ".main-image img",
                ".product-image img",
                "img[itemprop='image']"
            ):
                try:
                    images = driver.find_elements(
                        By.CSS_SELECTOR,
                        selector
                    )

                    for image in images:
                        src = (
                            image.get_attribute("src")
                            or image.get_attribute("data-src")
                        )
                        src = urljoin(url, src)

                        if src and src not in imagenes:
                            imagenes.append(src)

                    if imagenes:
                        break
                except Exception:
                    continue

        # Descripción y vídeos
        desc = ""
        video_urls = []

        try:
            desc_block = driver.find_element(
                By.CSS_SELECTOR,
                "div.product-description"
            )
            desc = desc_block.get_attribute("outerHTML") or ""

            desc = re.sub(
                r'<div[^>]*>\s*<a[^>]*href=["\'][^"\']*'
                r'cont(?:%C3%A1|á)ctenos[^"\']*["\'][^>]*>'
                r"\s*Contáctenos\s*</a>\s*</div>",
                "",
                desc,
                flags=re.IGNORECASE
            )

            video_urls = re.findall(
                r"https?://(?:www\.)?"
                r"(?:youtube\.com/(?:watch\?v=|embed/)|youtu\.be/)"
                r"[\w-]+(?:[^\s\"'<]*)?",
                desc,
                flags=re.IGNORECASE
            )
            video_urls = list(dict.fromkeys(video_urls))

        except:
            desc = ""

        if desc and not desc.lstrip().startswith("<"):
            desc = _construir_html_desde_texto(desc)

        # PDFs
        pdfs = {}

        try:
            pdf_rows = driver.find_elements(
                By.CSS_SELECTOR,
                "div.downloads div.product-field"
            )

            for row in pdf_rows:
                try:
                    name = row.find_element(
                        By.CSS_SELECTOR,
                        ".product-fields-title"
                    ).text.strip()

                    link = row.find_element(
                        By.CSS_SELECTOR,
                        ".product-field-display a[href]"
                    )
                    href = urljoin(url, link.get_attribute("href"))

                    if name and href:
                        pdfs[name] = href
                except Exception:
                    continue
        except:
            pass

        # Especificaciones
        detalles = {}
        barcode = ""
        weight = ""
        marca = ""

        try:
            rows = driver.find_elements(
                By.XPATH,
                "//div[contains(@class,'product-field')]"
                "[not(contains(@class,'product-field-type-G'))]"
                "[not(ancestor::div[contains(@class,'downloads')])]"
                "[div[contains(@class,'product-field-display')]]"
            )

            for row in rows:
                try:
                    key = row.find_element(
                        By.CSS_SELECTOR,
                        ".product-fields-title"
                    ).text.strip()

                    display = row.find_element(
                        By.CSS_SELECTOR,
                        ".product-field-display"
                    )

                    links = display.find_elements(
                        By.CSS_SELECTOR,
                        "a[href]"
                    )
                    images = display.find_elements(
                        By.CSS_SELECTOR,
                        "img"
                    )

                    if links:
                        value = urljoin(
                            url,
                            links[0].get_attribute("href")
                        )
                    elif images:
                        src = (
                            images[0].get_attribute("src")
                            or images[0].get_attribute("data-src")
                        )
                        value = urljoin(url, src)
                    else:
                        value = " ".join(display.text.split()).strip()

                    key = " ".join(key.split()).strip()

                    if key and value:
                        detalles[key] = value

                        if key.casefold() == "ean código".casefold():
                            barcode = str(value)
                        elif key.casefold() == "peso del artículo".casefold():
                            weight = value
                        elif key.casefold() == "marca".casefold():
                            marca = value

                except Exception:
                    continue

        except:
            pass

        # Categoría a partir de la URL
        x_categoria = ""

        try:
            path_parts = urlparse(url).path.strip("/").split("/")

            if len(path_parts) > 1:
                category_parts = path_parts[:-1]
                x_categoria = " / ".join(
                    unquote(part).replace("-", " ").title()
                    for part in category_parts
                    if part
                )
        except:
            pass

        producto = {
            ClavesExcel.URL.value: url,
            ClavesExcel.NOMBRE.value: titulo,
            ClavesExcel.IMAGEN.value: imagenes[0] if imagenes else "",
            ClavesExcel.GALERIA.value: imagenes,
            ClavesExcel.DESCRIPCION_WEB.value: desc,
            ClavesExcel.VIDEOS.value: video_urls,
            ClavesExcel.DOCUMENTOS.value: pdfs,
            ClavesExcel.ATRIBUTOS.value: detalles,
            ClavesExcel.SKU.value: sku,
            ClavesExcel.REFERENCIA.value: barcode,
            ClavesExcel.PESO.value: weight,
            ClavesExcel.MARCA.value: marca,
            ClavesExcel.CATEGORIA.value: x_categoria
        }

        # Igual que en el Scrapy original: un mismo SKU puede aparecer
        # en varias categorías. Se conserva una sola ficha y se unen.
        key = sku or titulo or url

        if key in productos_por_clave:
            existente = productos_por_clave[key]

            categorias_existentes = [
                categoria.strip()
                for categoria in str(
                    existente.get(ClavesExcel.CATEGORIA.value, "")
                ).split(",")
                if categoria.strip()
            ]

            if x_categoria and x_categoria not in categorias_existentes:
                categorias_existentes.append(x_categoria)

            existente[ClavesExcel.CATEGORIA.value] = ", ".join(
                categorias_existentes
            )

            for image in imagenes:
                if image not in existente[ClavesExcel.GALERIA.value]:
                    existente[ClavesExcel.GALERIA.value].append(image)

            for video in video_urls:
                if video not in existente[ClavesExcel.VIDEOS.value]:
                    existente[ClavesExcel.VIDEOS.value].append(video)

            for document_name, document_url in pdfs.items():
                existente[ClavesExcel.DOCUMENTOS.value].setdefault(
                    document_name,
                    document_url
                )

            for attr_name, attr_value in detalles.items():
                existente[ClavesExcel.ATRIBUTOS.value].setdefault(
                    attr_name,
                    attr_value
                )

        else:
            productos_por_clave[key] = producto
            productos.append(producto)

        producto_procesado += 1

    # Progreso final
    emit_progress(
        f"ESPAÑA: 100% ({productos_totales}/{productos_totales}) "
        "productos completados!"
    )

    try:
        driver.quit()
    except:
        pass

    return productos


# 📊 Exporta los datos scrapeados a Excel
def _guardar_excel(productos, excel_path):
    items = []
    region = "es"

    for p in productos:
        row = {
            ClavesExcel.URL.value: p[ClavesExcel.URL.value],
            ClavesExcel.NOMBRE.value: p[ClavesExcel.NOMBRE.value],
            ClavesExcel.IMAGEN.value: p[ClavesExcel.IMAGEN.value],
            ClavesExcel.GALERIA.value: p[ClavesExcel.GALERIA.value],
            ClavesExcel.DESCRIPCION_WEB.value: p[ClavesExcel.DESCRIPCION_WEB.value],
            ClavesExcel.VIDEOS.value: p[ClavesExcel.VIDEOS.value],
            ClavesExcel.DOCUMENTOS.value: p[ClavesExcel.DOCUMENTOS.value],
            ClavesExcel.ATRIBUTOS.value: json.dumps(
                p[ClavesExcel.ATRIBUTOS.value],
                ensure_ascii=False
            ),
            ClavesExcel.SKU.value: p[ClavesExcel.SKU.value],
            ClavesExcel.REFERENCIA.value: p[ClavesExcel.REFERENCIA.value],
            ClavesExcel.PESO.value: p[ClavesExcel.PESO.value],
            ClavesExcel.MARCA.value: p[ClavesExcel.MARCA.value],
            ClavesExcel.CATEGORIA.value: p[ClavesExcel.CATEGORIA.value]
        }
        items.append(row)

    Utils.save_to_excel(items, "vtac_es", region)
    Utils.excel_read_and_parse(
        "vtac_es",
        region
    )  # El parseo de atributos se guarda en scraped.xlsx


def _construir_html_desde_texto(texto):
    lineas = [
        line.strip()
        for line in texto.split("\n")
        if line.strip()
    ]
    html = "<div>" + "".join(
        f"<p>{line}</p>"
        for line in lineas
    ) + "</div>"
    return html


if __name__ == "__main__":
    run_vtac_es_scraper(modo_scrap=ScrapMode.COMPLETO)