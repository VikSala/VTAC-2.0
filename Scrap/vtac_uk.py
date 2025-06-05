import os
import time
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from services.utils import Utils
from enum import Enum


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


# 🔍 Extrae los enlaces de productos UK según el modo de scraping
def _get_product_links(driver, modo_scrap, only_new_products, file_path="", progress_callback=None):
    def emit_progress(message):
        if progress_callback:
            progress_callback(message)
        print(message)

    if modo_scrap == ScrapMode.TEST:
        emit_progress("MODO TEST UK ACTIVADO")
        return [
            "https://www.vtacexports.com/default/vt-5361-1m-micro-usb-cable-l-type-gold-diamond-series.html",
            "https://www.vtacexports.com/default/vt-5361-1m-micro-usb-cable-l-type-black-diamond-series.html",
            "https://www.vtacexports.com/default/vt-6043-4-60w-led-ceiling-fan-with-rf-control-4-blades-ac-motor.html",
            "https://www.vtacexports.com/default/vt-3042-3-30w-led-decorative-ceiling-fan-with-rf-control-cct-3-in-1-3-blades-dc-motor.html"
        ]

    emit_progress("Explorando categorías principales de Reino Unido...")

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
                urls_existentes = set(df_old["x_url_origen"].dropna().astype(str))
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

        storage = {}
        try:
            for li in driver.find_elements(By.CSS_SELECTOR, "div.grid.grid-cols-2 li.flex"):
                label = li.find_element(By.CSS_SELECTOR, "span.font-semibold").text.strip().replace(":", "")
                total = li.find_element(By.CSS_SELECTOR, "div.text-sm").text.strip().replace(label, "").strip()
                storage[label] = total
        except:
            pass

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
        datasheet = ""
        try:
            for a in driver.find_elements(By.CSS_SELECTOR, "a[href*='qpdf/product/generate']"):
                datasheet = a.get_attribute("href")
                break
        except:
            pass

        producto = {
            "x_url_origen": url,
            "Name": titulo,
            "standard_price": precio,
            "Image": imagenes[0] if imagenes else "",
            "Image_Urls": imagenes,
            "website_description": website_description,
            "Video_Urls": "",
            "Pdf_Urls": pdfs,
            "Specifications": especificaciones,
            "default_code": sku,
            "barcode": ean,
            "weight": peso,
            "x_marca": marca,
            "datasheet": datasheet,
            "storage": storage,
            "x_categoria": x_categoria
        }

        productos.append(producto)
        producto_procesado += 1

    # Progreso final
    emit_progress(f"UK: 100% ({productos_totales}/{productos_totales}) productos completados!")

    try:
        driver.quit()
    except:
        pass

    return productos


# 📄 Guarda los productos extraídos en un Excel
def _guardar_excel(productos, output_path):
    items = []
    for p in productos:
        row = {
            "x_url_origen": p["x_url_origen"],
            "Name": p["Name"],
            "standard_price": p["standard_price"],
            "Image": p["Image"],
            "Image_Urls": p["Image_Urls"],
            "website_description": p["website_description"],
            "Video_Urls": p["Video_Urls"],
            "Pdf_Urls": p["Pdf_Urls"],
            "Specifications": p["Specifications"],
            "default_code": p["default_code"],
            "barcode": p["barcode"],
            "weight": p["weight"],
            "x_marca": p["x_marca"],
            "datasheet": p["datasheet"],
            "storage": p["storage"],
            "x_categoria": p["x_categoria"]
        }
        items.append(row)

    Utils.save_to_excel(items, "vtac_uk")
    Utils.excel_read_and_parse("vtac_uk")

# Permite ejecutar directamente desde terminal
#if __name__ == "__main__": run_vtac_uk_scraper(modo_scrap=ScrapMode.MINI)
