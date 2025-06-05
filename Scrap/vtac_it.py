import os
import time
import json
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from services.utils import Utils
from enum import Enum


class ScrapMode(Enum):
    TEST = "Test"
    MINI = "Mini"
    COMPLETO = "Completo"


# ✨ MAIN FUNCION PUBLICA
def run_vtac_italia_scraper(modo_scrap: ScrapMode = ScrapMode.COMPLETO, only_new_products=False,
                            excel_output="vtac_it.xlsx", progress_callback=None):
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

    emit_progress(f"MODO {modo_scrap.value} IT ACTIVADO")

    driver = _init_driver()
    product_links = _get_product_links(driver, modo_scrap, only_new_products, file_path, emit_progress)
    productos = _scrape_productos(driver, product_links, emit_progress)

    if productos:
        _guardar_excel(productos, excel_output)
        emit_progress(f"ITALIA: 100%. Scraping guardado en {excel_output}")
    else:
        emit_progress("No se han scrapeado productos...")


# ⚙️ Inicializa el navegador con el perfil configurado
def _init_driver():
    profile_path = os.path.expanduser("~/AppData/Local/Google/Chrome/User Data/Profile_4_scrapy")
    options = uc.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument(f"--user-data-dir={profile_path}")
    options.add_argument("--profile-directory=Default")
    return uc.Chrome(options=options)


# 🔍 Obtiene las URLs de productos según el modo seleccionado
def _get_product_links(driver, modo_scrap, only_new_products, file_path="", progress_callback=None):
    def emit_progress(message):
        if progress_callback:
            progress_callback(message)
        print(message)

    product_links = []
    subcategorias = {}

    if modo_scrap == ScrapMode.TEST:
        emit_progress("MODO TEST ACTIVADO")
        product_links = [
            "https://led-italia.it/prodotti/v-tac/fotovoltaico/11868/stazione...",
            "https://led-italia.it/prodotti/v-tac/fotovoltaico/11529/inverter...",
            "https://led-italia.it/prodotti/v-tac/zanzariere-elettriche/11180/...",
        ]
    else:
        emit_progress("Explorando categorías principales...")
        categorias_principales = [
            "https://led-italia.it/prodotti/M4E-fotovoltaico",
            "https://led-italia.it/prodotti/M54-illuminazione-led",
            "https://led-italia.it/prodotti/M68-materiale-elettrico"
        ]

        for i, url_categoria in enumerate(categorias_principales, 1):
            emit_progress(f"📁 Buscando subcategorías ({i}/{len(categorias_principales)}): {url_categoria}")
            driver.get(url_categoria)
            time.sleep(1)

            try:
                contenedor = driver.find_element(By.CSS_SELECTOR,
                                                 "div.grid.grid-cols-2.lg\\:grid-cols-4.xl\\:grid-cols-6")
                enlaces = contenedor.find_elements(By.TAG_NAME, "a")
                for a in enlaces:
                    href = a.get_attribute("href")
                    try:
                        nombre = a.find_element(By.CSS_SELECTOR, "div.text-primary").text.strip()
                        if href and nombre:
                            subcategorias[nombre] = href
                    except:
                        continue
            except Exception as e:
                emit_progress(f"❌ No se encontró el contenedor en {url_categoria}: {e}")

        emit_progress(f"Encontradas {len(subcategorias)} subcategorías")

        for idx, (nombre_cat, shop_link) in enumerate(subcategorias.items(), 1):
            emit_progress(f"🔍 Procesando subcategoría ({idx}/{len(subcategorias)}): {nombre_cat}")

            driver.get(shop_link)
            time.sleep(1)
            try:
                pagination = driver.find_elements(By.CSS_SELECTOR, "div.mt-4 ul.flex li div")
                last_page = max([int(p.text.strip()) for p in pagination if p.text.strip().isdigit()])
            except:
                last_page = 1

            if "sub%5B" in shop_link:
                shop_pages = [f"{shop_link}&page={i}" if i > 0 else shop_link for i in range(last_page)]
            else:
                shop_pages = [f"{shop_link}?page={i}" if i > 0 else shop_link for i in range(last_page)]

            for page_idx, page_url in enumerate(shop_pages, 1):
                if len(shop_pages) > 1:
                    emit_progress(f"  📄 Página {page_idx}/{len(shop_pages)}")

                driver.get(page_url)
                time.sleep(3)
                grid = driver.find_elements(By.CSS_SELECTOR, "div.px-2.grid a")
                for a in grid:
                    href = a.get_attribute("href")
                    if href and "/prodotti/" in href and href not in product_links:
                        product_links.append(href)
                        if modo_scrap == ScrapMode.MINI:
                            break
                if modo_scrap == ScrapMode.MINI:
                    break

    if only_new_products:
        emit_progress("🔍 Filtrando productos ya existentes...")
        try:
            if file_path:
                df_old = pd.read_excel(file_path)
                urls_existentes = set(df_old["x_url_origen"].dropna().astype(str))
                original_count = len(product_links)
                product_links = [url for url in product_links if url not in urls_existentes]
                emit_progress(f"Productos totales: {original_count}, Nuevos detectados: {len(product_links)}")
                return product_links
            else:
                emit_progress("No se seleccionó ningún archivo. Se devolverán todos los productos.")
        except Exception as e:
            emit_progress(f"Error al comparar productos previos: {e}")

    emit_progress(f"Total de productos para procesar: {len(product_links)}")
    return product_links


# 👁️ Scrapea la información de cada producto desde su URL
def _scrape_productos(driver, product_links, progress_callback=None):
    def emit_progress(message):
        if progress_callback:
            progress_callback(message)
        print(message)

    productos = []
    productos_totales = len(product_links)
    producto_procesado = 0

    emit_progress(f"🚀 Iniciando scraping de {productos_totales} productos...")

    for i, url in enumerate(product_links, 1):
        # Mostrar progreso cada 10 productos o en hitos importantes
        if i % 10 == 0 or i == 1:
            porcentaje = int((i / productos_totales) * 100)
            emit_progress(f"ITALIA: {porcentaje}% ({i}/{productos_totales}) productos procesados...")

        driver.get(url)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        driver.execute_script("document.body.style.zoom='25%'")
        time.sleep(3)

        # Título
        try:
            titulo = driver.find_element(By.CSS_SELECTOR, "h2.text-md.text-primary.mt-1").text.strip()
        except:
            titulo = "Título no encontrado"

        # Detalles
        detalles = {}
        try:
            blocks = driver.find_elements(By.CSS_SELECTOR, "div.flex.border-b-thin")
            for div in blocks:
                textos = div.text.strip().split("\n")
                if len(textos) == 2:
                    key, value = textos
                    detalles[key.strip()] = value.strip()
        except:
            pass

        detalles_corregido = {key.replace("Ataque", "Casquillo"): val for key, val in detalles.items()}
        detalles = detalles_corregido

        # Descripción y video
        video_url = ""
        try:
            desc_block = driver.find_element(By.CSS_SELECTOR, "div.space-y-4 div.cms.prose")
            desc = desc_block.text.strip()
            if "Vídeo ilustrativo" in desc:
                iframe = desc_block.find_element(By.CSS_SELECTOR, "iframe[src*='youtube.com']")
                video_url = iframe.get_attribute("src")
                desc += f'<p/><iframe style="aspect-ratio: 16/9" src="{video_url}" allow="accelerometer; encrypted-media; gyroscope" frameborder="0" allowfullscreen=""></iframe>'
        except:
            desc = ""

        # Precio
        try:
            main = driver.find_element(By.TAG_NAME, "main")
            elementos = main.find_elements(By.XPATH, ".//*")
            precio = next((el.text.strip() for el in elementos if "€" in el.text), "")
        except:
            precio = ""

        # Imágenes
        imagenes = []
        try:
            img_tags = driver.find_elements(By.CSS_SELECTOR, "#images-slider-list img")
            for img in img_tags:
                src = img.get_attribute("src")
                if src and src not in imagenes:
                    imagenes.append(src)
        except:
            pass

        # PDFs
        pdfs = {}
        try:
            pdf_links = driver.find_elements(By.CSS_SELECTOR, "div.grid a[href$='.pdf']")
            for a in pdf_links:
                href = a.get_attribute("href")
                if "v-tac-scheda-tecnica" not in href:
                    text = a.text.strip()
                    pdfs[text] = href
        except:
            pass

        # Categoría desde breadcrumbs visibles
        x_categoria = ""
        try:
            breadcrumb_links = driver.find_elements(By.CSS_SELECTOR, "div.lg\\:flex.items-center a")
            if len(breadcrumb_links) >= 3:
                nivel1 = breadcrumb_links[1].text.strip()
                nivel2 = breadcrumb_links[2].text.strip()
                x_categoria = f"{nivel1}/{nivel2}"
        except:
            pass

        producto = {
            "x_url_origen": url,
            "Name": titulo,
            "standard_price": precio.strip().replace("\u00a3", "").strip(),
            "Image": imagenes[0] if imagenes else "",
            "Image_Urls": imagenes,
            "website_description": _construir_html_desde_texto(desc),
            "Video_Urls": video_url,
            "Pdf_Urls": pdfs,
            "Specifications": detalles,
            "default_code": detalles.get("Código SKU") or detalles.get("SKU"),
            "barcode": detalles.get("EAN", ""),
            "weight": detalles.get("Peso", ""),
            "x_marca": detalles.get("Marca", ""),
            "x_categoria": x_categoria
        }

        productos.append(producto)
        producto_procesado += 1

    # Progreso final
    emit_progress(f"ITALIA: 100% ({productos_totales}/{productos_totales}) productos completados!")

    try:
        driver.quit()
    except:
        pass

    return productos


# 📊 Exporta los datos scrapeados a Excel
def _guardar_excel(productos, excel_path):
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
            "Specifications": json.dumps(p["Specifications"], ensure_ascii=False),
            "default_code": p["default_code"],
            "barcode": p["barcode"],
            "weight": p["weight"],
            "x_marca": p["x_marca"],
            "x_categoria": p["x_categoria"]
        }
        items.append(row)

    Utils.save_to_excel(items, "vtac_it")
    Utils.excel_read_and_parse("vtac_it")


def _construir_html_desde_texto(texto):
    lineas = [line.strip() for line in texto.split("\n") if line.strip()]
    html = "<div>" + "".join(f"<p>{line}</p>" for line in lineas) + "</div>"
    return html
