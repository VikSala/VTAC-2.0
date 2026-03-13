import time
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def run_ledxpress_product_test():
    url = "https://www.ledxpress.com/sp_es/lamparas-de-mesa-lamparas-de-carga-inalambrica-ip54-blanco-1-5-vatios-150-lumenes-3000k.html"
    print(f"🌐 Cargando producto: {url}")

    # --- Inicializar navegador ---
    options = uc.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    driver = uc.Chrome(options=options)
    driver.get(url)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    driver.execute_script("document.body.style.zoom='25%'")
    time.sleep(1)

    item = {}

    # 0️⃣ Título del producto
    titulo = ""
    try:
        h1_elem = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h1.title-font"))
        )
        titulo = h1_elem.text.strip()
    except Exception as e:
        print(f"⛔ Error localizando el título: {e}")

    item["Título"] = titulo

    # 1️⃣ SKU y Código de familia
    try:
        # Buscar cualquier elemento con texto que contenga "SKU"
        sku_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'SKU')]")
        if sku_elements:
            item["SKU"] = sku_elements[0].text.strip()
        else:
            item["SKU"] = ""
    except Exception as e:
        print(f"⚠️ Error al leer SKU: {e}")
        item["SKU"] = ""

    # 2️⃣ Almacén: texto dentro del div con la clase de grid
    try:
        almacen_div = driver.find_element(By.CSS_SELECTOR, "div.grid.grid-cols-1.gap-4.sm\\:gap-0.mb-2.md\\:mt-0.mt-4")
        almacen_text = almacen_div.text.strip().split(":")[1]
        print(f"🏬 Texto de almacén detectado:\n{almacen_text}")
    except Exception as e:
        almacen_text = ""
        print(f"⚠️ No se encontró el bloque de almacén: {e}")

    item["Almacén"] = almacen_text

    # 3️⃣ Iconos (versión nueva con lógica específica)
    iconos = []

    try:
        print("🎨 Buscando iconos en la página...")

        # 1️⃣ Buscar span con class="warranty-text"
        warranty_spans = driver.find_elements(By.CSS_SELECTOR, "span.warranty-text")
        if warranty_spans:
            print(f"🟢 Detectado span con warranty-text ({len(warranty_spans)} encontrado/s)")
            iconos.append("http://79.72.55.217:8069/web/image/399898-78bbd986/Garantia5.svg")
        else:
            print("⚪ No se encontró ningún span.warranty-text")

        # 2️⃣ Buscar cualquier elemento con href que contenga el icono Samsung
        samsung_href = "https://www.ledxpress.com/static/version1759820377/frontend/Vtac/hyva/es_ES/images/icons/samsung.svg"
        href_elements = driver.find_elements(By.XPATH, f"//*[@src='{samsung_href}']")
        if href_elements:
            print(f"🟢 Detectado icono de Samsung ({len(href_elements)} encontrado/s)")
            iconos.append(samsung_href)
        else:
            print("⚪ No se encontró el icono de Samsung")

    except Exception as e:
        print(f"⚠️ Error detectando iconos: {e}")

    item["Iconos"] = iconos
    print(f"✅ Iconos guardados: {len(iconos)} → {iconos}")

    # 2️⃣ Precio normal y con descuento
    try:
        price_normal = driver.find_element(By.CSS_SELECTOR, "div.final-price span.price").text.strip()
    except:
        price_normal = ""

    try:
        price_old = driver.find_element(By.CSS_SELECTOR, "div.old-price span.price").text.strip()
    except:
        price_old = ""

    item["Precio final"] = price_normal
    item["Precio"] = price_old

    # 3️⃣ Galería de imágenes
    try:
        images = []
        HASH_OLD = "3380650127d143eec657262365bd2ea0"
        HASH_NEW = "207e23213cf636ccdef205098cf3c8a3"

        thumbs = driver.find_elements(By.CSS_SELECTOR, "#thumb-carousel img")

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

    # Características
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

    except Exception as e:
        item["Resumen"] = ""
        print(f"⚠️ No se encontró descripción_ecommerce: {e}")

    '''# Descripción
    try:
        description_div = driver.find_element(By.ID, "product-features")
        item["Descripción_HTML"] = description_div.get_attribute("outerHTML")
    except:
        item["Descripción_HTML"] = ""'''

    # 6️⃣ Atributos → diccionario {th: td}  (versión con prints de depuración)
    atributos = {}
    try:
        cont = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "product-attributes"))
        )
        html_preview = cont.get_attribute("outerHTML")

        # Buscar filas: prioriza la tabla 'additional-attributes', pero cae a cualquier <tr> si no existe
        rows = cont.find_elements(By.CSS_SELECTOR, "table.additional-attributes > tbody > tr")
        if not rows:
            print("⚠️ No se encontró 'table.additional-attributes', probando selector genérico 'tr'…")
            rows = cont.find_elements(By.CSS_SELECTOR, "tr")

        for idx, row in enumerate(rows, start=1):
            try:
                th_elems = row.find_elements(By.TAG_NAME, "th")
                td_elems = row.find_elements(By.TAG_NAME, "td")

                if not th_elems or not td_elems:
                    # Log de la fila problematica
                    print(f"    ⚠️ Fila {idx} sin th/td válidos. HTML: {row.get_attribute('outerHTML')[:200]}")
                    continue

                # Extrae texto robustamente (text o textContent)
                key = (th_elems[0].text or th_elems[0].get_attribute("textContent") or "").strip()
                val = (td_elems[0].text or td_elems[0].get_attribute("textContent") or "").strip()

                # Limpieza básica
                if key.endswith(":"):
                    key = key[:-1].strip()

                if key:
                    atributos[key] = val
            except Exception as e:
                print(f"    ❌ Error parseando fila {idx}: {e}")

    except Exception as e:
        print(f"⛔ Error localizando o leyendo #product-attributes: {e}")

    item["Atributos"] = atributos

    # 7️⃣ Documentos
    documentos = {}

    try:
        # 1️⃣ Clic en la pestaña "Descargas"
        try:
            downloads_tab = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "tab-label-product.downloads"))
            )
            driver.execute_script("arguments[0].click();", downloads_tab)
            print("✅ Pestaña 'Descargas' abierta correctamente.")
        except Exception as e:
            print(f"⚠️ No se pudo hacer clic en la pestaña 'Descargas': {e}")

        # 2️⃣ Esperar a que aparezcan los documentos
        WebDriverWait(driver, 10).until(
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

    driver.quit()
    print("🛑 Navegador cerrado.")

    # --- Mostrar resultados ---
    print("\n🧩 RESULTADO PRODUCTO LEDXPRESS\n")
    for k, v in item.items():
        print(f"{k}:")
        if isinstance(v, (list, dict)):
            print(v)
        else:
            print(f"  {v}")
        print("-" * 80)

    return item


if __name__ == "__main__":
    run_ledxpress_product_test()
