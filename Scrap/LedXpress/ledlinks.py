import time
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
import pandas as pd

# Categorías a procesar
CATEGORIES = {
    "Bombillas y tubos LED": "https://www.ledxpress.com/sp_es/lamparas-y-tubos-led.html",
    "Iluminación Interior": "https://www.ledxpress.com/sp_es/iluminacion-interior.html",
    "Iluminación Exterior": "https://www.ledxpress.com/sp_es/iluminacion-exterior.html",
    "Iluminación Decorativa": "https://www.ledxpress.com/sp_es/iluminacion-decorativa.html",
    "Productos Solares": "https://www.ledxpress.com/sp_es/productos-solares.html",
    "Iluminación Profesional": "https://www.ledxpress.com/sp_es/iluminacion-profesional.html",
    "Aspectos Eléctricos Esenciales": "https://www.ledxpress.com/sp_es/aspectos-electricos-esenciales.html",
    "Productos Inteligentes": "https://www.ledxpress.com/sp_es/productos-inteligentes.html",
}

# Umbral de detección mínima por página
MIN_PRODUCTS_PER_PAGE = 100

def init_driver():
    """Inicializa Chrome UC con opciones antidetector."""
    options = uc.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    return uc.Chrome(options=options)


def scrape_category(driver, url):
    """Recorre todas las páginas de una categoría hasta que no haya más resultados."""
    product_links = {}
    page = 1
    last_count = 0

    while True:
        page_url = f"{url}?p={page}&product_list_limit=100"
        print(f"\n🌀 Cargando página {page}: {page_url}")
        driver.get(page_url)
        time.sleep(1)

        # Buscar productos
        items = driver.find_elements(
            By.CSS_SELECTOR,
            "a.product.photo.product-item-photo.relative.flex.justify-center.w-full.mx-auto.mb-2.p-4"
        )

        num_found = len(items)
        print(f"📦 Detectados {num_found} productos en página {page}")

        # Criterios de parada
        if num_found == 0:
            print("⚠️ No hay productos en esta página. Fin de la categoría.")
            break

        # Acumular enlaces
        current_total = len(product_links)
        for a in items:
            href = a.get_attribute("href")
            img = a.find_element(By.TAG_NAME, "img")
            title = img.get_attribute("title") or ""
            product_links[href] = title

        added_this_page = len(product_links) - current_total
        print(f"➕ Añadidos {added_this_page} nuevos (total acumulado: {len(product_links)})")

        # Si no se añaden nuevos, paramos
        if added_this_page == 0:
            print("⚠️ No se añadieron nuevos enlaces. Fin de la categoría.")
            break

        if num_found < MIN_PRODUCTS_PER_PAGE:
            print(f"⚠️ Menos de {MIN_PRODUCTS_PER_PAGE} productos ({num_found}). Fin de la categoría.")
            break

        page += 1
        last_count = num_found

    return product_links

def run_ledxpress_scraper():
    """Scrapea los productos de las categorías definidas en CATEGORIES y guarda el resultado en Excel."""
    driver = init_driver()
    productos = []

    try:
        total_categorias = len(CATEGORIES)
        print(f"📦 Procesando {total_categorias} categorías principales...")

        for idx, (categoria_nombre, categoria_url) in enumerate(CATEGORIES.items(), start=1):
            print("\n" + "=" * 90)
            print(f"[{idx}/{total_categorias}] 📂 Categoría: {categoria_nombre}")
            print(f"🌐 URL: {categoria_url}")

            # Scrapeamos los productos de esa categoría (usando tu función existente scrape_category)
            productos_categoria = scrape_category(driver, categoria_url)

            if not productos_categoria:
                print(f"⚠️ No se encontraron productos en: {categoria_nombre}")
                continue

            # Guardar los resultados con la ruta asociada
            for href, nombre in productos_categoria.items():
                productos.append({
                    "URL": href,
                    "Nombre": nombre,
                    "Ruta": categoria_nombre
                })

            print(f"✅ {len(productos_categoria)} productos capturados en '{categoria_nombre}'")

    finally:
        driver.quit()
        print("\n🛑 Navegador cerrado.")

    # Guardar resultados en Excel
    if productos:
        df = pd.DataFrame(productos)
        output_file = "ledxpress_productos_por_categoria.xlsx"
        df.to_excel(output_file, index=False)
        print(f"\n📦 Resultados guardados en: {output_file} ({len(productos)} productos totales)")
    else:
        print("⚠️ No se obtuvieron productos de ninguna categoría.")


if __name__ == "__main__":
    run_ledxpress_scraper()
