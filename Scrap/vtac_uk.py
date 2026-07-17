import os
import time
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from services.utils import Utils
from enum import Enum
from services.campos_odoo import ClavesExcel


class ScrapMode(Enum):
    TEST = "Test"
    MINI = "Mini"
    COMPLETO = "Completo"


# ✨ Funcíon principal para lanzar el scraping de Reino Unido
def run_vtac_uk_scraper(modo_scrap: ScrapMode = ScrapMode.TEST, only_new_products=False, excel_output="vtac_uk.xlsx",
                        progress_callback=None):
    file_path = ""
    if only_new_products:
        from tkinter import Tk, filedialog
        Tk().withdraw()
        file_path = filedialog.askopenfilename(title="Selecciona el Excel de productos anteriores",
                                               filetypes=[("Excel Files", "*.xlsx")])

    # Función helper para emitir progreso
    def emit_progress(message):
        if progress_callback:
            progress_callback(message)
        print(message)  # También imprimir en consola

    emit_progress(f"MODO {modo_scrap.value} UK ACTIVADO")

    driver = _init_driver()
    product_links = _get_product_links(driver, modo_scrap, only_new_products, file_path, emit_progress)
    productos = _scrape_productos(driver, product_links, emit_progress)

    if productos:
        _guardar_excel(productos, excel_output)
        emit_progress(f"UK: 100%. Scraping guardado en {excel_output}")
    else:
        emit_progress("No se han scrapeado productos...")


# 🔧 Inicializa el navegador con perfil Chrome traducido
def _init_driver():
    profile_path = os.path.expanduser("~/AppData/Local/Google/Chrome/User Data/Profile_4_scrapy")
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

# 🔍 Extrae los enlaces de productos UK según el modo de scraping
def _get_product_links(driver, modo_scrap, only_new_products, file_path="", progress_callback=None):
    def emit_progress(message):
        if progress_callback:
            progress_callback(message)
        print(message)

    CATEGORIES_LINKS = [
        'https://www.vtacexports.com/default/led-lighting.html',
        'https://www.vtacexports.com/default/decorative-lighting.html',
        'https://www.vtacexports.com/default/smart-products.html',
        'https://www.vtacexports.com/default/digital-accessories.html',
        'https://www.vtacexports.com/default/electrical.html',
        'https://www.vtacexports.com/default/energy.html',
        'https://www.vtacexports.com/default/uk-new-arrivals.html',
        'https://www.vtacexports.com/default/back-in-stock.html',
        'https://www.vtacexports.com/default/top-products.html'
    ]

    SCRAP_SELECTIVO = False

    if not SCRAP_SELECTIVO:

        if modo_scrap == ScrapMode.TEST:
            emit_progress("MODO TEST UK ACTIVADO")
            return [
                "https://www.vtacexports.com/default/vt-44003-300w-led-floodlight-cree-chip-1m-wire-4000k-black-body-135lm-w-6yrs-wty-ip65.html",
                "https://www.vtacexports.com/default/vt-61024-24w-backlit-recessed-panel-samsung-chip-4000k-rd.html"
            ]

        emit_progress("Explorando categorías principales de Reino Unido...")

        shop_links = set()

        for idx, start_url in enumerate(CATEGORIES_LINKS, 1):
            emit_progress(
                f"🔍 Procesando categoría ({idx}/{len(CATEGORIES_LINKS)}): {start_url.split('/')[-1].replace('.html', '').replace('-', ' ').title()}")

            current_url = start_url
            page_count = 0

            while current_url:
                page_count += 1
                if page_count > 1:
                    emit_progress(f"  📄 Página {page_count} de la categoría")

                driver.get(current_url)
                time.sleep(1)

                try:
                    enlaces = driver.find_elements(By.CSS_SELECTOR, "a[href]")
                    products_found_in_page = 0

                    for a in enlaces:
                        href = a.get_attribute("href")
                        if href and "default/vt-" in href:
                            if href not in shop_links:
                                shop_links.add(href)
                                products_found_in_page += 1

                            if modo_scrap == ScrapMode.MINI:
                                break

                    if products_found_in_page > 0:
                        emit_progress(f"  ✅ Encontrados {products_found_in_page} productos en esta página")

                except Exception as e:
                    emit_progress(f"  ❌ Error al leer enlaces: {e}")

                if modo_scrap == ScrapMode.MINI:
                    break

                # Buscar botón "siguiente"
                try:
                    next_button = driver.find_element(By.CSS_SELECTOR, "a.action.next")
                    next_href = next_button.get_attribute("href")
                    if next_href and next_href != current_url:
                        current_url = next_href
                    else:
                        current_url = None
                except:
                    current_url = None

            emit_progress(f"  📊 Total acumulado hasta ahora: {len(shop_links)} productos únicos")

        emit_progress(f"🎯 Total de productos encontrados en UK: {len(shop_links)}")

        if only_new_products:
            emit_progress("🔍 Filtrando productos ya existentes...")
            try:
                if file_path:
                    df_old = pd.read_excel(file_path)
                    urls_existentes = set(df_old[ClavesExcel.URL.value].dropna().astype(str))
                    original_count = len(shop_links)
                    product_links = [url for url in shop_links if url not in urls_existentes]
                    emit_progress(f"Productos totales: {original_count}, Nuevos detectados: {len(product_links)}")
                    return product_links
                else:
                    emit_progress("No se seleccionó ningún archivo. Se devolverán todos los productos.")
            except Exception as e:
                emit_progress(f"Error al comparar productos previos: {e}")

        emit_progress(f"Total de productos para procesar: {len(shop_links)}")
        return list(shop_links)

    else:
        emit_progress("🚀 MODO SCRAP SELECTIVO ACTIVADO")

        import re
        from services.utils import Utils
        file_path = Utils.seleccionar_excel()
        SKU_EXCEL_PATH = file_path
        SKU_COLUMN_NAME = "SKU"

        # 1) Cargar SKUs válidos desde el Excel
        try:
            df_sku = pd.read_excel(SKU_EXCEL_PATH, usecols=[SKU_COLUMN_NAME])
            allowed_skus = set(df_sku[SKU_COLUMN_NAME].dropna().astype(str).str.strip())
            emit_progress(f"📑 SKUs cargados: {len(allowed_skus)}")
        except Exception as e:
            emit_progress(f"❌ No se pudo leer el Excel de SKUs: {e}")
            return []

        # 2) Recorrer categorías y filtrar por SKU
        shop_links = set()
        SKU_RE = re.compile(r"\d+")

        for idx, start_url in enumerate(CATEGORIES_LINKS, 1):
            emit_progress(
                f"🔍 Procesando categoría ({idx}/{len(CATEGORIES_LINKS)}): "
                f"{start_url.split('/')[-1].replace('.html', '').replace('-', ' ').title()}"
            )

            current_url = start_url
            page_count = 0

            while current_url:
                page_count += 1
                if page_count > 1:
                    emit_progress(f"  📄 Página {page_count} de la categoría")

                driver.get(current_url)
                time.sleep(1)  # pequeño respiro para que cargue

                try:
                    product_cards = driver.find_elements(By.CSS_SELECTOR, "form[method='post']")
                    encontrados_en_pagina = 0
                    print(f"Total: {len(product_cards)} productos")

                    for card in product_cards:
                        # ── SKU ──────────────────────────────────────
                        try:
                            sku_elem = card.find_element(By.CSS_SELECTOR, "div.bg-dark")
                            match = SKU_RE.search(sku_elem.text)
                            if not match:
                                print("No hay SKU")
                                continue
                            sku = match.group(0)
                        except Exception:
                            continue

                        # ── Filtro por lista de SKUs ────────────────
                        if sku not in allowed_skus:
                            continue

                        # ── Enlace ──────────────────────────────────
                        try:
                            link_elem = card.find_element(By.CSS_SELECTOR, "a[href*='default/vt-']")
                            href = link_elem.get_attribute("href")
                        except Exception:
                            continue

                        if href and href not in shop_links:
                            shop_links.add(href)
                            encontrados_en_pagina += 1

                        if modo_scrap == ScrapMode.MINI:
                            break

                    if encontrados_en_pagina > 0:
                        emit_progress(f"  ✅ Añadidos {encontrados_en_pagina} productos en esta página")

                except Exception as e:
                    emit_progress(f"  ❌ Error leyendo productos: {e}")

                # ── Botón siguiente página ─────────────────────────
                try:
                    next_button = driver.find_element(By.CSS_SELECTOR, "a.action.next")
                    next_href = next_button.get_attribute("href")
                    if next_href and next_href != current_url:
                        current_url = next_href
                    else:
                        current_url = None
                except Exception:
                    current_url = None

            emit_progress(f"  📊 Total acumulado hasta ahora: {len(shop_links)} productos")

        emit_progress(f"🎯 Productos encontrados tras filtrado por SKU: {len(shop_links)}")
        return list(shop_links)


# 📊 Procesa cada URL de producto y extrae todos sus campos
def _scrape_productos(driver, product_links, progress_callback=None):
    def emit_progress(message):
        if progress_callback:
            progress_callback(message)
        print(message)

    productos = []
    productos_totales = len(product_links)
    producto_procesado = 0

    emit_progress(f"🚀 Iniciando scraping de {productos_totales} productos de Reino Unido...")

    for i, url in enumerate(product_links, 1):

        if i > 1 and i % 300 == 0:
            driver = _reiniciar_driver(driver)

        # Mostrar progreso cada 10 productos o en hitos importantes
        if i % 10 == 0 or i == 1:
            porcentaje = int((i / productos_totales) * 100)
            emit_progress(f"UK: {porcentaje}% ({i}/{productos_totales}) productos procesados...")

        driver.get(url)
        time.sleep(1)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        driver.execute_script("document.body.style.zoom='25%'")

        # Categoría
        x_categoria = ""
        try:
            breadcrumb_items = driver.find_elements(By.CSS_SELECTOR, "nav.breadcrumbs ol.items li.item")
            categorias = []
            for li in breadcrumb_items:
                # Saltar el último (producto actual)
                if "product" in li.get_attribute("class") and "parent_product" not in li.get_attribute("class"):
                    continue
                try:
                    texto = li.find_element(By.TAG_NAME, "a").text.strip()
                    if texto:
                        categorias.append(texto)
                except:
                    continue
            x_categoria = "/".join(categorias)
        except Exception as e:
            if i <= 5:  # Solo mostrar errores en los primeros productos para no spam
                emit_progress(f"No se pudo extraer x_categoria del producto {i}: {e}")

        # Título
        try:
            titulo = driver.find_element(By.CSS_SELECTOR, "div.title-font").text.strip()
        except:
            titulo = "Título no encontrado"

        # Imágenes
        imagenes = []
        try:
            for img in driver.find_elements(By.CSS_SELECTOR, "#gallery img"):
                src = img.get_attribute("src")
                if src:
                    imagenes.append(src)
        except:
            pass

        # SKU y EAN
        sku = ""
        ean = ""
        try:
            for div in driver.find_elements(By.CSS_SELECTOR, "div.mb-6 > div"):
                t = div.text.strip()
                if "SKU" in t:
                    sku = t.split("SKU")[-1].strip()
                elif "EAN" in t:
                    ean = t.split("EAN")[-1].strip()
        except:
            pass

        # Precio
        precio = ""
        try:
            p = driver.find_element(By.CSS_SELECTOR, "div.text-primary.font-semibold.text-3xl")
            precio = p.text.strip().split(":")[-1].replace("\u00a3", "").strip()
        except:
            pass

        # Especificaciones y almacenamiento
        especificaciones = {}
        peso = ""
        marca = "V-TAC"
        try:
            for ul in driver.find_elements(By.CSS_SELECTOR, "#description ul"):
                for li in ul.find_elements(By.TAG_NAME, "li"):
                    key = li.find_element(By.CSS_SELECTOR, ".product-attribute-label").text.strip()
                    val = li.find_element(By.CSS_SELECTOR, ".product-attribute-value").text.strip()
                    if key == "Peso bruto (kg)":
                        peso = val
                    if key == "Samsung":
                        marca = key
                    especificaciones[key] = val
        except:
            pass

        correcciones = {
            "Ataque": "Casquillo",
            "ANTES DE CRISTO": "AC"
        }

        # Aplicar las correcciones
        detalles_corregido = {}
        for key, val in especificaciones.items():
            nueva_key = key
            for original, reemplazo in correcciones.items():
                if original in nueva_key:
                    nueva_key = nueva_key.replace(original, reemplazo)
            detalles_corregido[nueva_key] = val

        especificaciones = detalles_corregido

        # Storage
        '''storage = {}
        try:
            for li in driver.find_elements(By.CSS_SELECTOR, "div.grid.grid-cols-2 li.flex"):
                label = li.find_element(By.CSS_SELECTOR, "span.font-semibold").text.strip().replace(":", "")
                total = li.find_element(By.CSS_SELECTOR, "div.text-sm").text.strip().replace(label, "").strip()
                storage[label] = total
        except:
            pass'''

        # Descripción
        website_description = ""
        try:
            driver.find_element(By.CSS_SELECTOR, "#tab-label-features a").click()
            time.sleep(.5)
            lis = driver.find_elements(By.CSS_SELECTOR, "#features .product-features-desc li")
            website_description = "<ul>" + "".join(f"<li>{li.text.strip()}</li>" for li in lis) + "</ul>"
        except:
            pass

        # PDFs
        pdfs = {}
        try:
            driver.find_element(By.CSS_SELECTOR, "#tab-label-product\\.downloads a").click()
            time.sleep(1)
            items = driver.find_elements(By.CSS_SELECTOR, "#product-attachments .attachment-item")
            for item in items:
                link = item.find_element(By.TAG_NAME, "a")
                href = link.get_attribute("href")
                name = link.find_element(By.CSS_SELECTOR, "span.font-semibold").text.strip()
                if href and name:
                    pdfs[name] = href
        except:
            pass

        # Datasheet
        '''datasheet = ""
        try:
            for a in driver.find_elements(By.CSS_SELECTOR, "a[href*='qpdf/product/generate']"):
                datasheet = a.get_attribute("href")
                break
        except:
            pass'''

        producto = {
            ClavesExcel.URL.value: url,
            ClavesExcel.NOMBRE.value: titulo,
            ClavesExcel.COSTE.value: precio,
            ClavesExcel.IMAGEN.value: imagenes[0] if imagenes else "",
            ClavesExcel.GALERIA.value: imagenes,
            ClavesExcel.DESCRIPCION_WEB.value: website_description,
            ClavesExcel.VIDEOS.value: "",
            ClavesExcel.DOCUMENTOS.value: pdfs,
            ClavesExcel.ATRIBUTOS.value: especificaciones,
            ClavesExcel.SKU.value: sku,
            ClavesExcel.REFERENCIA.value: ean,
            ClavesExcel.PESO.value: peso,
            ClavesExcel.MARCA.value: marca,
            ClavesExcel.CATEGORIA.value: x_categoria
            #"datasheet": datasheet,
            #"storage": storage,
        }

        productos.append(producto)
        producto_procesado += 1

    # Progreso final
    emit_progress(f"UK: 100% ({productos_totales}/{productos_totales}) productos completados!")
    print(f"UK: 100% ({productos_totales}/{productos_totales}) productos completados!")
    try:
        driver.quit()
    except:
        pass

    return productos


# 📄 Guarda los productos extraídos en un Excel
def _guardar_excel(productos, output_path):
    items = []
    region = "uk"
    for p in productos:
        row = {
            ClavesExcel.URL.value: p[ClavesExcel.URL.value],
            ClavesExcel.NOMBRE.value: p[ClavesExcel.NOMBRE.value],
            ClavesExcel.IMAGEN.value: p[ClavesExcel.IMAGEN.value],
            ClavesExcel.GALERIA.value: p[ClavesExcel.GALERIA.value],
            ClavesExcel.DESCRIPCION_WEB.value: p[ClavesExcel.DESCRIPCION_WEB.value],
            ClavesExcel.VIDEOS.value: p[ClavesExcel.VIDEOS.value],
            ClavesExcel.DOCUMENTOS.value: p[ClavesExcel.DOCUMENTOS.value],
            ClavesExcel.ATRIBUTOS.value: p[ClavesExcel.ATRIBUTOS.value],
            ClavesExcel.SKU.value: p[ClavesExcel.SKU.value],
            ClavesExcel.REFERENCIA.value: p[ClavesExcel.REFERENCIA.value],
            ClavesExcel.PESO.value: p[ClavesExcel.PESO.value],
            ClavesExcel.MARCA.value: p[ClavesExcel.MARCA.value],
            ClavesExcel.CATEGORIA.value: p[ClavesExcel.CATEGORIA.value],
            ClavesExcel.COSTE.value: p[ClavesExcel.COSTE.value]
            #"datasheet": p["datasheet"],
            #"storage": p["storage"],
        }
        items.append(row)

    Utils.save_to_excel(items, "vtac_uk", region)
    Utils.excel_read_and_parse("vtac_uk", region)

# Permite ejecutar directamente desde terminal
if __name__ == "__main__": run_vtac_uk_scraper(modo_scrap=ScrapMode.TEST)
