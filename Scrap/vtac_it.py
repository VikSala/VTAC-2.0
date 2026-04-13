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
    categorias_principales = [
        "https://led-italia.it/prodotti/M4E-fotovoltaico",
        "https://led-italia.it/prodotti/M54-illuminazione-led",
        "https://led-italia.it/prodotti/M68-materiale-elettrico"
    ]

    SCRAP_SELECTIVO = False

    if not SCRAP_SELECTIVO:

        if modo_scrap == ScrapMode.TEST:
            emit_progress("MODO TEST ACTIVADO")
            product_links = [
                "https://led-italia.it/prodotti/v-tac/lampadine/21200216/lampadina-led-chip-cree-e27-22w-120lmw-g120-3000k?asq=ac67518b93a2324006af653033dafc74&c=44af3e96a6fe780b105761e98fb9e62b",
                "https://led-italia.it/prodotti/v-tac/lampadine/21200226/lampadina-led-chip-cree-e27-22w-120lmw-g120-4000k?asq=ac67518b93a2324006af653033dafc74&c=f3fb3dec41b1c0dbffb3d74e421582ae",
                "https://led-italia.it/prodotti/v-tac/lampadine/21200236/lampadina-led-chip-cree-e27-22w-120lmw-g120-6500k?asq=ac67518b93a2324006af653033dafc74&c=78b73f73f656060c5c48dafd9ff98df9",
                "https://led-italia.it/prodotti/v-tac/lampadine/212356/lampadina-led-chip-cree-e14-2w-st26-4000k?asq=ac67518b93a2324006af653033dafc74&c=62891e0020cdf5e386ff7175179c4a31",
                "https://led-italia.it/prodotti/v-tac/lampadine/2144566/lampadina-led-chip-cree-e27-17w-110lmw-a65-3000k?asq=ac67518b93a2324006af653033dafc74&c=0f05622b9611833eb61289c43bdd5dbe",
                "https://led-italia.it/prodotti/v-tac/lampadine/2144576/lampadina-led-chip-cree-e27-17w-100lmw-a65-4000k?asq=ac67518b93a2324006af653033dafc74&c=4cda1734021db9312754b5a4604af29d",
                "https://led-italia.it/prodotti/v-tac/tubi/216506/tubo-led-chip-cree-t8-9w-g13-60cm-in-nanoplastica-ruotabile-3000k?asq=8107fefdbe91fa336ef2f09ca6719a64&c=a00e03527ada2113598ae110532c1a83",
                "https://led-italia.it/prodotti/v-tac/tubi/216516/tubo-led-chip-cree-t8-9w-g13-60cm-in-nanoplastica-ruotabile-4000k?asq=8107fefdbe91fa336ef2f09ca6719a64&c=98fe5aff9d54c67adba2b6f6002a4a35",
                "https://led-italia.it/prodotti/v-tac/tubi/216526/tubo-led-chip-cree-t8-9w-g13-60cm-in-nanoplastica-ruotabile-6500k?asq=8107fefdbe91fa336ef2f09ca6719a64&c=f670b6f91a13cd1701ca52351a5e4ed3",
                "https://led-italia.it/prodotti/v-tac/tubi/216556/tubo-led-chip-cree-t8-18w-100lmw-g13-120cm-in-nanoplastica-ruotabile-6500k?asq=8107fefdbe91fa336ef2f09ca6719a64&c=e0391cbda20618565377ed7e5e3a5faf",
                "https://led-italia.it/prodotti/v-tac/tubi/216566/tubo-led-chip-cree-t8-20w-105lmw-g13-150cm-in-nanoplastica-ruotabile-3000k?asq=8107fefdbe91fa336ef2f09ca6719a64&c=d54c914a2b8b1cd33fdb4525479f7341",
                "https://led-italia.it/prodotti/v-tac/illuminazione-di-emergenza/899-6/lampada-led-chip-cree-di-emergenza-38w-20led-montaggio-a-incassoplafone-e-modalita-sase-12h-di-ricarica-6000k-ip20?asq=b1cab57a42c92b551bc2d11eac03f150&c=e0b4a8a06acfce67d4f8077869f3e6cf",
            ]
        else:
            emit_progress("Explorando categorías principales...")

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

            #emit_progress(f"Encontradas {len(subcategorias)} subcategorías")

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
                    grid = driver.find_elements(By.CSS_SELECTOR, "span.px-2.grid a")#No lo detecta
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
                    urls_existentes = set(df_old[ClavesExcel.URL.value].dropna().astype(str))
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

    else:
        emit_progress("🚀 MODO SCRAP SELECTIVO ITALIA ACTIVADO")
        import re
        #from services.utils import Utils
        #file_path = Utils.seleccionar_excel()
        SKU_EXCEL_PATH = file_path
        SKU_COLUMN_NAME = "SKU"

        # 1️⃣  Cargar SKUs válidos
        try:
            df_sku = pd.read_excel(SKU_EXCEL_PATH, usecols=[SKU_COLUMN_NAME])
            allowed_skus = set(df_sku[SKU_COLUMN_NAME].dropna().astype(str).str.strip())
            emit_progress(f"📑 SKUs cargados: {len(allowed_skus)}")
        except Exception as e:
            emit_progress(f"❌ No se pudo leer el Excel de SKUs: {e}")
            return []

        product_links = []  # se reutiliza la variable declarada arriba
        subcategorias = {}
        PAT_SKU = re.compile(r"SKU\s*([0-9]+)", re.I)

        # ── Obtener subcategorías ──────────────────────────────────

        emit_progress("Explorando categorías principales…")

        for i, url_categoria in enumerate(categorias_principales, 1):
            emit_progress(f"📁 Buscando subcategorías ({i}/{len(categorias_principales)}): {url_categoria}")
            driver.get(url_categoria)
            time.sleep(1)

            try:
                contenedor = driver.find_element(
                    By.CSS_SELECTOR,
                    "div.grid.grid-cols-2.lg\\:grid-cols-4.xl\\:grid-cols-6"
                )
                enlaces = contenedor.find_elements(By.TAG_NAME, "a")

                for a in enlaces:
                    href = a.get_attribute("href")
                    try:
                        nombre = a.find_element(By.CSS_SELECTOR, "div.text-primary").text.strip()
                        if href and nombre:
                            subcategorias[nombre] = href
                    except Exception:
                        continue
            except Exception as e:
                emit_progress(f"❌ No se encontró el contenedor en {url_categoria}: {e}")

        # ── Recorrer subcategorías y páginas ───────────────────────

        for idx, (nombre_cat, shop_link) in enumerate(subcategorias.items(), 1):
            emit_progress(f"🔍 Procesando subcategoría ({idx}/{len(subcategorias)}): {nombre_cat}")
            driver.get(shop_link)
            time.sleep(1)
            try:
                pagination = driver.find_elements(By.CSS_SELECTOR, "div.mt-4 ul.flex li div")
                last_page = max([int(p.text.strip()) for p in pagination if p.text.strip().isdigit()])
            except Exception:
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

                try:
                    product_cards = driver.find_elements(
                        By.CSS_SELECTOR,
                        "div.select-none.border-thin"
                    )
                    print(f"Total: {len(product_cards)} productos")
                    encontrados_en_pagina = 0

                    for card in product_cards:

                        # 2.1 SKU
                        try:
                            sku_elem = card.find_element(By.CSS_SELECTOR, "p.text-primary")
                            m = PAT_SKU.search(sku_elem.text)
                            if not m:
                                continue
                            sku = m.group(1)
                        except Exception:
                            continue

                        if sku not in allowed_skus:
                            continue  # filtrado selectivo

                        # 2.2 Enlace
                        try:
                            link_elem = card.find_element(By.CSS_SELECTOR, "a[href*='/prodotti/']")
                            href = link_elem.get_attribute("href")
                        except Exception:
                            continue

                        if href and href not in product_links:
                            product_links.append(href)
                            encontrados_en_pagina += 1

                    if encontrados_en_pagina:
                        emit_progress(f"  ✅ Añadidos {encontrados_en_pagina} productos en esta página")


                except Exception as e:
                    emit_progress(f"  ❌ Error leyendo productos: {e}")

        emit_progress(f"🎯 Productos encontrados tras filtrado por SKU: {len(product_links)}")
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

        #detalles_corregido = {key.replace("Ataque", "Casquillo"): val for key, val in detalles.items()}
        #detalles = detalles_corregido
        correcciones = {
            "Ataque": "Casquillo",
            "ANTES DE CRISTO": "AC"
        }

        # Aplicar las correcciones
        detalles_corregido = {}
        for key, val in detalles.items():
            nueva_key = key
            for original, reemplazo in correcciones.items():
                if original in nueva_key:
                    nueva_key = nueva_key.replace(original, reemplazo)
            detalles_corregido[nueva_key] = val

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

            def extraer_precio(texto):
                # 1. Buscar precio promocional (variaciones: promocional, promoción, promo)
                promo_match = re.search(
                    r'precio\s+(?:promocional|promoción)[^\d€]*€?\s*([\d\.]+,\d{2})',
                    texto,
                    re.IGNORECASE
                )
                if promo_match:
                    return promo_match.group(1)

                # 2. Buscar cualquier otro "precio" con número después (general)
                normal_match = re.search(
                    r'precio[^\d€]*€?\s*([\d\.]+,\d{2})',
                    texto,
                    re.IGNORECASE
                )
                if normal_match:
                    return normal_match.group(1)

                return None

            precio = extraer_precio(precio)

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
            pdf_links = driver.find_elements(
                By.CSS_SELECTOR,
                "div.grid.grid-cols-1.lg\\:flex.lg\\:flex-wrap.gap-2.py-2 a"
            )

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
            # obtenemos todos los enlaces como antes
            breadcrumb_links = driver.find_elements(
                By.XPATH,
                "//div[contains(@class,'gap-2')]/descendant::a"
            )

            categorias = []
            for link in breadcrumb_links:
                texto = link.text.strip()
                if not texto:  # ignora vacíos
                    continue
                if "SKU" in texto:  # cortamos justo antes de la SKU
                    break
                categorias.append(texto)

            x_categoria = " / ".join(categorias)
        except:
            pass

        if precio: precio.strip().replace("\u00a3", "").strip()
        else: precio=0

        producto = {
            ClavesExcel.URL.value: url,
            ClavesExcel.NOMBRE.value: titulo,
            ClavesExcel.PRECIO.value: precio,
            ClavesExcel.IMAGEN.value: imagenes[0] if imagenes else "",
            ClavesExcel.GALERIA.value: imagenes,
            ClavesExcel.DESCRIPCION_WEB.value: _construir_html_desde_texto(desc),
            ClavesExcel.VIDEOS.value: video_url,
            ClavesExcel.DOCUMENTOS.value: pdfs,
            ClavesExcel.ATRIBUTOS.value: detalles,
            ClavesExcel.SKU.value: detalles.get("Código SKU") or detalles.get("SKU"),
            ClavesExcel.REFERENCIA.value: detalles.get("EAN", ""),
            ClavesExcel.PESO.value: detalles.get("Peso", ""),
            ClavesExcel.MARCA.value: detalles.get("Marca", ""),
            ClavesExcel.CATEGORIA.value: x_categoria
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
    region = "it"
    for p in productos:
        row = {
            ClavesExcel.URL.value: p[ClavesExcel.URL.value],
            ClavesExcel.NOMBRE.value: p[ClavesExcel.NOMBRE.value],
            ClavesExcel.IMAGEN.value: p[ClavesExcel.IMAGEN.value],
            ClavesExcel.GALERIA.value: p[ClavesExcel.GALERIA.value],
            ClavesExcel.DESCRIPCION_WEB.value: p[ClavesExcel.DESCRIPCION_WEB.value],
            ClavesExcel.VIDEOS.value: p[ClavesExcel.VIDEOS.value],
            ClavesExcel.DOCUMENTOS.value: p[ClavesExcel.DOCUMENTOS.value],
            ClavesExcel.ATRIBUTOS.value: json.dumps(p[ClavesExcel.ATRIBUTOS.value], ensure_ascii=False),
            ClavesExcel.SKU.value: p[ClavesExcel.SKU.value],
            ClavesExcel.REFERENCIA.value: p[ClavesExcel.REFERENCIA.value],
            ClavesExcel.PESO.value: p[ClavesExcel.PESO.value],
            ClavesExcel.MARCA.value: p[ClavesExcel.MARCA.value],
            ClavesExcel.CATEGORIA.value: p[ClavesExcel.CATEGORIA.value],
            ClavesExcel.PRECIO.value: p[ClavesExcel.PRECIO.value]
        }
        items.append(row)

    Utils.save_to_excel(items, "vtac_it", region)
    Utils.excel_read_and_parse("vtac_it", region) # El parseo de atributos se guarda en scraped.xlsx


def _construir_html_desde_texto(texto):
    lineas = [line.strip() for line in texto.split("\n") if line.strip()]
    html = "<div>" + "".join(f"<p>{line}</p>" for line in lineas) + "</div>"
    return html

if __name__ == "__main__": run_vtac_italia_scraper(modo_scrap=ScrapMode.TEST)