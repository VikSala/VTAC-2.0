import time
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from collections import namedtuple
import re
from bs4 import BeautifulSoup
from services.utils import Utils

# -------------------------------------------------------------------------------------------------
# CONFIGURACIÓN
# -------------------------------------------------------------------------------------------------
CATEGORIES = [
    "https://www.ledxpress.com/sp_es/lamparas-y-tubos-led.html",
    "https://www.ledxpress.com/sp_es/iluminacion-interior.html",
    "https://www.ledxpress.com/sp_es/iluminacion-exterior.html",
    "https://www.ledxpress.com/sp_es/iluminacion-decorativa.html",
    "https://www.ledxpress.com/sp_es/productos-solares.html",
    "https://www.ledxpress.com/sp_es/iluminacion-profesional.html",
    "https://www.ledxpress.com/sp_es/aspectos-electricos-esenciales.html",
    "https://www.ledxpress.com/sp_es/productos-inteligentes.html",
]

MIN_PRODUCTS_PER_PAGE = 100
OUTPUT_EXCEL = "ledxpress_productos"

# Diccionarios
FINAL_LEVELS = {}
IS_EXCEL = True

# Pares (key, value) para pasar a scrape_product como espera tu función
LinkKV = namedtuple("LinkKV", ["key", "value"])

# -------------------------------------------------------------------------------------------------
# DRIVER
# -------------------------------------------------------------------------------------------------
def init_driver():
    options = uc.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    return uc.Chrome(options=options)

def web_css_selector(driver, selector):
    return WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
    )

# -------------------------------------------------------------------------------------------------
# RECOGER PRODUCTOS DE UNA CATEGORÍA (SIN SUBNIVELES)
# -------------------------------------------------------------------------------------------------
def scrape_products_from_leaf(driver, url):
    """
    Recoge todos los productos de una categoría final (sin subcategorías).
    Respeta el límite de 100 por página y detiene cuando no hay más o hay < MIN_PRODUCTS_PER_PAGE.
    """
    all_links = {}
    page = 1
    while True:
        page_url = f"{url}?p={page}&product_list_limit=100"
        driver.get(page_url)
        time.sleep(2.0)

        # Productos
        items = driver.find_elements(
            By.CSS_SELECTOR,
            "a.product.photo.product-item-photo.relative.flex.justify-center.w-full.mx-auto.mb-2.p-4"
        )
        imgs = driver.find_elements(
            By.CSS_SELECTOR,
            "img.object-contain.product-image-photo"
        )
        num_found = len(items)

        if num_found == 0:
            break

        current_total = len(all_links)
        for a, img in zip(items, imgs):
            href = a.get_attribute("href")
            text = img.get_attribute("title")
            if href and text:
                all_links[href] = text

        added = len(all_links) - current_total
        if added == 0:
            break

        if num_found < MIN_PRODUCTS_PER_PAGE:
            break

        page += 1

    return all_links

# -------------------------------------------------------------------------------------------------
# EXPLORACIÓN RECURSIVA DE SUBCATEGORÍAS
# -------------------------------------------------------------------------------------------------
def explore_category(driver, url, breadcrumb):
    """
    Entra en la URL, si hay subcategorías (div[x-data="subcategoryTile()"]) las recorre recursivamente.
    Si no hay subcategorías, es un último nivel: recoge los productos y guarda en FINAL_LEVELS.
    """
    driver.get(url)
    time.sleep(1.0)

    # Buscar subcategorías: anclas dentro del bloque x-data="subcategoryTile()"
    subcats = driver.find_elements(By.XPATH, '//div[@x-data="subcategoryTile()"]//a[@href]')
    subcat_links = [(a.text.strip(), a.get_attribute("href")) for a in subcats if a.text.strip()]

    if subcat_links:
        # Hay subniveles: bajar recursivamente
        for name, href in subcat_links:
            new_breadcrumb = f"{breadcrumb} / {name}"
            explore_category(driver, href, new_breadcrumb)
    else:
        # Último nivel: recoger productos
        links = scrape_products_from_leaf(driver, url)
        FINAL_LEVELS[breadcrumb] = links
        print(f"🧩 Nivel final: {breadcrumb} → {len(links)} productos")

# -------------------------------------------------------------------------------------------------
# SCRAPEAR DETALLES DE UN PRODUCTO
# -------------------------------------------------------------------------------------------------
def scrape_product(driver, url):
    test = False

    if test:
        print(f"\n🌐 Scrapeando producto: {url.value}")

        driver.get(url.value)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        driver.execute_script("document.body.style.zoom='25%'")
        time.sleep(1)

        item = {"URL de origen": url.value}
        try:
            # 1️⃣ Buscar cualquier div que contenga un atributo x-html con texto tipo "fullText"
            desc_div = driver.find_element(By.XPATH, '//div[@x-html]')

            # 2️⃣ Buscar el <ul> dentro de ese div
            ul_elem = desc_div.find_element(By.TAG_NAME, "ul")
            li_elems = ul_elem.find_elements(By.TAG_NAME, "li")

            # 3️⃣ Construir el HTML <ul> limpio
            li_items = [f"<li>{li.text.strip()}</li>" for li in li_elems if li.text.strip()]
            ul_str = '<ul class="product-features">' + "".join(li_items) + "</ul>"

            item["Resumen"] = ul_str

        except:
            try:
                # 1️⃣ Buscar cualquier div que contenga un atributo x-html con texto tipo "fullText"
                desc_div = driver.find_element(By.CSS_SELECTOR, "div.product-description")
                                                #"div.grid.grid-cols-1.gap-4.sm\\:gap-0.mb-2.md\\:mt-0.mt-4")

                # 2️⃣ Buscar el <ul> dentro de ese div
                ul_elem = desc_div.find_element(By.TAG_NAME, "ul")
                li_elems = ul_elem.find_elements(By.TAG_NAME, "li")

                # 3️⃣ Construir el HTML <ul> limpio
                li_items = [f"<li>{li.text.strip()}</li>" for li in li_elems if li.text.strip()]
                ul_str = '<ul class="product-features">' + "".join(li_items) + "</ul>"

                item["Resumen"] = ul_str
            except Exception as e:
                item["Resumen"] = ""
                print(f"⚠️ No se encontró descripción_ecommerce: {e}")
        return item
    else:
        print(f"\n🌐 Scrapeando producto: {url.value}")

        driver.get(url.value)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        driver.execute_script("document.body.style.zoom='25%'")
        time.sleep(1)

        item = {"URL de orígen": url.value}
        if not IS_EXCEL:
            item["Nombre"] = FINAL_LEVELS[url.key].get(url.value, "")
            item["Categoría"] = url.key

        print("Atributos")# Atributos
        atributos = {}
        try:
            WebDriverWait(driver, 10).until(
                lambda d: d.find_element(By.ID, "product-attributes")
                          and "Material" in d.find_element(By.ID, "product-attributes").text
                          or len(d.find_elements(By.CSS_SELECTOR, "#product-attributes tr")) > 7
            )

            # Una vez renderizada la tabla real, obtenemos el HTML estático
            cont_html = driver.find_element(By.ID, "product-attributes").get_attribute("outerHTML")

            soup = BeautifulSoup(cont_html, "html.parser")

            for idx, tr in enumerate(soup.select("tr"), start=1):
                try:
                    th = tr.select_one("th")
                    td = tr.select_one("td")
                    if not th or not td:
                        continue

                    key = th.get_text(strip=True).rstrip(":")
                    val = td.get_text(strip=True)

                    if key == "Código Ean":
                        item["EAN"] = val
                    elif key == "Peso bruto (Kgs)":
                        item["Peso"] = val
                    elif key == "Marca":
                        item["Marca"] = val
                    elif key == "Modelo":
                        item["Modelo"] = val

                    atributos[key] = val
                except Exception as e:
                    print(f"    ⚠️ Fila {idx} con error: {e}")

        except Exception as e:
            print(f"⛔ Error localizando o leyendo #product-attributes: {e}")

        item["Atributos"] = atributos

        '''print("Almacen")# 2️⃣ Almacén: texto dentro del div con la clase de grid
        try:
            almacen_div = web_css_selector(driver, "div.grid.grid-cols-1.gap-4.sm\\:gap-0.mb-2.md\\:mt-0.mt-4")
            almacen_text = almacen_div.text.strip().split(":")[1]
        except Exception as e:
            almacen_text = "No está disponible"
            print(f"⚠️ No se encontró el bloque de almacén: {e}")

        item["Descripción"] = almacen_text'''

        print("Iconos")# 3️⃣ Iconos (versión nueva con lógica específica)
        iconos = []

        try:
            # 1️⃣ Buscar span con class="warranty-text"
            warranty_spans = driver.find_elements(By.CSS_SELECTOR, "span.warranty-text")
            if warranty_spans:
                iconos.append("http://79.72.55.217:8069/web/image/399898-78bbd986/Garantia5.svg")

            # 2️⃣ Buscar cualquier elemento con href que contenga el icono Samsung
            samsung_href = "https://www.ledxpress.com/static/version1759820377/frontend/Vtac/hyva/es_ES/images/icons/samsung.svg"
            href_elements = driver.find_elements(By.XPATH, "//div[contains(@class, 'product-details')]//*[contains(@src, 'samsung.svg')]")
            if href_elements:
                iconos.append(samsung_href)

        except Exception as e:
            print(f"⚠️ Error detectando iconos: {e}")

        item["Iconos"] = iconos

        print("Precios")# Precios
        try:
            item["Precio final"] = web_css_selector(driver, "div.final-price span.price").text.strip()#driver.find_element(By.CSS_SELECTOR, "div.final-price span.price").text.strip()
            trigger_decimal = item["Precio final"].split(',')[1]
        except:
            item["Precio final"] = web_css_selector(driver, "div.final-price span.price span.price").text.strip()
            print("Doble span...")
        try:
            item["Precio"] = web_css_selector(driver, "div.old-price span.price").text.strip()#driver.find_element(By.CSS_SELECTOR, "div.old-price span.price").text.strip()
            trigger_decimal = item["Precio final"].split(',')[1]
        except:
            item["Precio"] = item["Precio final"]

        print("Galeria")# Galería
        try:
            images = []
            HASH_OLD = "3380650127d143eec657262365bd2ea0"
            HASH_NEW = "207e23213cf636ccdef205098cf3c8a3"

            thumbs = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#thumb-carousel img"))
            )#driver.find_elements(By.CSS_SELECTOR, "#thumb-carousel img")

            for img in thumbs:
                src = img.get_attribute("src")
                if not src:
                    continue

                # Sustituir el hash antiguo por el nuevo
                if HASH_OLD in src:
                    src = src.replace(HASH_OLD, HASH_NEW)

                # Evitar duplicados
                if src not in images:
                    images.append(src)

            # Asignar resultados
            if images:
                item["Imagen principal"] = images[0]
                item["Galería"] = images[1:] if len(images) > 1 else []
            else:
                item["Imagen principal"] = ""
                item["Galería"] = []

        except Exception as e:
            print(f"⚠️ Error al procesar galería: {e}")
            item["Imagen principal"] = ""
            item["Galería"] = []

        print("Features")# Características
        try:
            features_div = driver.find_element(By.ID, "product-features-info")
            item["Descripción Web"] = features_div.get_attribute("outerHTML")
        except:
            item["Descripción Web"] = ""

        try:
            # 1️⃣ Buscar cualquier div que contenga un atributo x-html con texto tipo "fullText"
            desc_div = driver.find_element(By.XPATH, '//div[@x-html]')

            # 2️⃣ Buscar el <ul> dentro de ese div
            ul_elem = desc_div.find_element(By.TAG_NAME, "ul")
            li_elems = ul_elem.find_elements(By.TAG_NAME, "li")

            # 3️⃣ Construir el HTML <ul> limpio
            li_items = [f"<li>{li.text.strip()}</li>" for li in li_elems if li.text.strip()]
            ul_str = '<ul class="product-features">' + "".join(li_items) + "</ul>"

            item["Resumen"] = ul_str

        except:
            try:
                # 1️⃣ Buscar cualquier div que contenga un atributo x-html con texto tipo "fullText"
                desc_div = driver.find_element(By.CSS_SELECTOR, "div.grid.grid-cols-1.gap-4.sm\\:gap-0.mb-2.md\\:mt-0.mt-4")

                # 2️⃣ Buscar el <ul> dentro de ese div
                ul_elem = desc_div.find_element(By.TAG_NAME, "ul")
                li_elems = ul_elem.find_elements(By.TAG_NAME, "li")

                # 3️⃣ Construir el HTML <ul> limpio
                li_items = [f"<li>{li.text.strip()}</li>" for li in li_elems if li.text.strip()]
                ul_str = '<ul class="product-features">' + "".join(li_items) + "</ul>"

                item["Resumen"] = ul_str
            except Exception as e:
                item["Resumen"] = ""
                print(f"⚠️ No se encontró descripción_ecommerce: {e}")

        print("Docs")# Documentos
        documentos = {}

        try:
            # 1️⃣ Clic en la pestaña "Descargas"
            try:
                downloads_tab = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.ID, "tab-label-product.downloads"))
                )
                driver.execute_script("arguments[0].click();", downloads_tab)
            except Exception:
                print("⚠️ No se pudo hacer clic en la pestaña 'Descargas'")

            # 2️⃣ Esperar a que aparezcan los documentos
            WebDriverWait(driver, 2).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.attachment-item"))
            )

            div_links = driver.find_elements(By.CSS_SELECTOR, "div.attachment-item")

            for div in div_links:
                try:
                    a_tag = div.find_element(By.CSS_SELECTOR, "a[href]")
                    href = a_tag.get_attribute("href").strip()

                    # Obtener el título desde el <span> mediante JS (texto renderizado)
                    title = driver.execute_script("""
                        let el = arguments[0].querySelector('span.flex-grow.font-semibold.text-center.text-ssm');
                        return el ? el.innerText.trim() : '';
                    """, a_tag)

                    # Fallbacks por si el título no existe o es "Descarga"
                    if not title or title.lower() == "descarga":
                        title = a_tag.get_attribute("title") or "Documento"

                    # Evita duplicados y valores vacíos
                    if href and title:
                        documentos[title] = href

                except Exception as e:
                    print(f"⚠️ Error procesando documento: {e}")

        except Exception as e:
            print(f"⚠️ Error general obteniendo documentos: {e}")

        item["Documentos"] = documentos

        print("SKU")# SKU
        try:
            sku_elements = WebDriverWait(driver, 10).until(
                lambda d: d.find_elements(By.XPATH, "//*[contains(text(), 'SKU')]")
            )
            if sku_elements:
                sku_text = sku_elements[0].text.strip()
                # Extraer solo el número o código después de "SKU"
                match = re.search(r'SKU\s*([A-Za-z0-9\-]+)', sku_text, re.IGNORECASE)
                if match:
                    item["SKU"] = match.group(1).strip()
                else:
                    item["SKU"] = ""
            else:
                item["SKU"] = ""
        except Exception as e:
            print(f"⚠️ Error extrayendo SKU: {e}")
            item["SKU"] = ""

        print(f"✅ Producto scrapeado: {item['SKU']}...")
        return item

# -------------------------------------------------------------------------------------------------
# MAIN
# -------------------------------------------------------------------------------------------------
def run_ledxpress_full_scraper():
    driver = init_driver()
    productos = []

    try:
        if IS_EXCEL:
            ruta = Utils.seleccionar_excel()
            print(f"📘 Cargando Excel: {ruta}")

            df = pd.read_excel(ruta)

            # Validar columnas necesarias
            required_cols = {"URL de orígen", "Nombre", "Categoría"}
            if not required_cols.issubset(df.columns):
                raise ValueError(f"El Excel debe contener las columnas: {', '.join(required_cols)}")

            total_rows = len(df)
            print(f"🔗 Total de productos en Excel: {total_rows}")

            from collections import namedtuple
            LinkKV = namedtuple("LinkKV", ["key", "value"])

            for idx, row in df.iterrows():
                url = str(row["URL de orígen"]).strip()
                name = str(row["Nombre"]).strip()
                categoria = str(row["Categoría"]).strip()

                if not url or url == "nan":
                    continue

                print(f"\n[{idx + 1}/{total_rows}] Procesando producto desde Excel → {name}")
                kv = LinkKV(key=categoria, value=url)
                producto = scrape_product(driver, kv)
                producto["Nombre"] = name
                producto["Categoría"] = categoria
                producto["URL de orígen"] = url
                productos.append(producto)

        else:
            # 1) Explorar categorías principales recursivamente hasta últimos niveles
            for category_url in CATEGORIES:
                print("\n" + "=" * 90)
                print(f"📂 Categoría principal: {category_url}")
                # Usamos el slug como primer breadcrumb
                breadcrumb_root = category_url.rstrip("/").split("/")[-1].replace(".html", "")
                explore_category(driver, category_url, breadcrumb_root)

            # 2) Imprimir solo últimos niveles (ruta completa y nº de links)
            print("\n📊 RESUMEN DE NIVELES FINALES")
            for path, links in FINAL_LEVELS.items():
                print(f"• {path}: {len(links)} links")

            # 3) Scrappear cada producto pasando (key=bREADCRUMB, value=URL)
            total_leafs = sum(len(v) for v in FINAL_LEVELS.values())
            print(f"\n🔗 Total de enlaces a procesar: {total_leafs}")

            idx = 0
            for breadcrumb, links in FINAL_LEVELS.items():
                for href, name in links.items():
                    idx += 1
                    print(f"\n[{idx}/{total_leafs}] Procesando producto: {name}")
                    kv = LinkKV(key=breadcrumb, value=href)
                    producto = scrape_product(driver, kv)
                    producto["Nombre"] = name  # conserva el texto del enlace
                    productos.append(producto)

    finally:
        driver.quit()
        print("\n🛑 Navegador cerrado.")

    # 4) Guardar resultados
    if productos:
        #df = pd.DataFrame(productos)
        #df.to_excel(OUTPUT_EXCEL, index=False)
        save_to_excel_(productos, OUTPUT_EXCEL, "LEDXPRESS")
        print(f"📦 Resultados guardados en: {OUTPUT_EXCEL}")
    else:
        print("⚠️ No se han obtenido productos.")

def save_to_excel_(items, filename, region):
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

    with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
        df_main.to_excel(writer, sheet_name=region.upper(), index=False)

if __name__ == "__main__":
    run_ledxpress_full_scraper()
