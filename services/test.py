import tkinter as tk
from tkinter import messagebox
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.common.keys import Keys  # Asegúrate de que esta importación esté presente
import time
import os

sitios_visitados = [
    "http://smartrekking.com:8069/",
    "http://otro-sitio.com:8080/"
]

def mostrar_alerta(mensaje):
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Configuración necesaria", mensaje)
    root.destroy()

def agregar_sitio_al_script(url_nuevo):
    script_path = os.path.abspath(__file__)
    with open(script_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    nueva_lista = []
    dentro_de_lista = False
    modificada = False

    for line in lines:
        if line.strip().startswith("sitios_visitados = ["):
            dentro_de_lista = True
            nueva_lista.append(line)
            continue

        if dentro_de_lista:
            if line.strip().startswith("]"):
                if url_nuevo not in "".join(nueva_lista):
                    nueva_lista.insert(-1, f'    "{url_nuevo}",\n')
                    modificada = True
                dentro_de_lista = False
                nueva_lista.append(line)
            else:
                nueva_lista.append(line)
        else:
            nueva_lista.append(line)

    if modificada:
        with open(script_path, "w", encoding="utf-8") as f:
            f.writelines(nueva_lista)
        print(f"[INFO] Sitio '{url_nuevo}' añadido al script con éxito.")
    else:
        print(f"[INFO] Sitio ya estaba en la lista o no se modificó.")

def configurar_descarga_si_nueva(url_odoo):
    if url_odoo in sitios_visitados:
        print(f"[INFO] Sitio ya configurado: {url_odoo}")
        return True

    print(f"[INFO] Primera vez accediendo a: {url_odoo}")
    url_codificada = quote(url_odoo, safe='')
    url_chrome_settings = f"chrome://settings/content/siteDetails?site={url_codificada}"

    profile_path = os.path.expanduser("~/AppData/Local/Google/Chrome/User Data/Profile_4_scrapy")
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")
    options.add_argument(f"--user-data-dir={profile_path}")
    options.add_argument("--profile-directory=Default")

    driver = webdriver.Chrome(options=options)

    try:
        driver.get(url_chrome_settings)
        #mostrar_alerta(f"🚫 Es la primera vez que accedes a este sitio.\n\n👉 Por favor, permite las descargas automáticas en:\n{url_chrome_settings}\n\nLuego cierra esta ventana y vuelve a ejecutar el programa.")
        mostrar_alerta(
            f"🛠 Es la primera vez que accedes a este sitio:\n\n{url_odoo}\n\n"
            f"1️⃣ Permite las DESCARGAS AUTOMÁTICAS en:\n\n{url_chrome_settings}\n\n"
            "2️⃣ Luego, en:\n    chrome://settings/content/insecureContent\n"
            "   en 'Puede mostrar contenido no seguro' dale a Añadir esta URL: \n\n{url_odoo}\n\n.\n\n"
            "Cuando termines, espera a que termine de ejecutarse el script y vuelve a ejecutarlo."
        )
        time.sleep(15)
        driver.get("chrome://settings/content/insecureContent")
        time.sleep(15)
    finally:
        driver.quit()
        agregar_sitio_al_script(url_odoo)
    return False

def realizar_exportacion_smartrekking(username="tu_usuario_odoo", password="tu_password_odoo", actualizar_estado=None):
    url_base = "http://smartrekking.com:8069"
    url_login = f"{url_base}/web/login"
    url_objetivo = f"{url_base}/odoo/action-452?view_type=list"

    profile_path = os.path.expanduser("~/AppData/Local/Google/Chrome/User Data/Profile_4_scrapy")
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")
    options.add_argument(f"--user-data-dir={profile_path}")
    options.add_argument("--profile-directory=Default")

    driver = webdriver.Chrome(options=options)
    print("[INFO] WebDriver iniciado.")
    if actualizar_estado:
        actualizar_estado("[INFO] WebDriver iniciado.")

    try:
        print(f"[INFO] Navegando a la página de login: {url_login}")
        if actualizar_estado:
            actualizar_estado(f"[PROGRESO] Intentando login en {url_base}...")
        driver.get(url_login)

        login_input = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "login"))
        )
        login_input.clear()
        login_input.send_keys(username)
        print(f"[INFO] Usuario '{username}' introducido.")

        password_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "password"))
        )
        password_input.clear()
        password_input.send_keys(password)
        print("[INFO] Contraseña introducida.")

        # --- MODIFICACIÓN AQUÍ: Pulsar Enter en el campo de contraseña ---
        print("[INFO] Intentando enviar login presionando Enter en el campo de contraseña...")
        try:
            password_input.send_keys(Keys.RETURN)
            print("[INFO] Tecla Enter presionada en el campo de contraseña.")
        except Exception as e_login_enter:
            print(f"[ERROR_CRITICO] Fallo al intentar enviar Enter en el campo de contraseña: {e_login_enter}")
            driver.save_screenshot("error_login_enter_action.png")
            print("[INFO] Captura de pantalla 'error_login_enter_action.png' guardada.")
            raise
        # --- FIN DE LA MODIFICACIÓN ---

        print("[INFO] Login enviado. Esperando brevemente antes de navegar a la URL objetivo...")
        time.sleep(3)

        print(f"[INFO] Login procesado (asumido). Navegando directamente a: {url_objetivo}")
        if actualizar_estado:
            actualizar_estado(f"[PROGRESO] Navegando a la página de exportación...")
        driver.get(url_objetivo)

        print(f"[INFO] Esperando a que cargue la página objetivo '{url_objetivo}' y aparezca 'checkbox-comp-1'...")
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "checkbox-comp-1"))
        )
        print("[INFO] Página objetivo cargada y 'checkbox-comp-1' presente.")
        if actualizar_estado:
            actualizar_estado("[INFO] Página de exportación cargada.")
        time.sleep(1)

        # 2. Clic en el checkbox para seleccionar todos los elementos (de la página actual)
        print("[INFO] Intentando hacer clic en el checkbox principal 'checkbox-comp-1'...")
        if actualizar_estado:
            actualizar_estado("[PROGRESO] Seleccionando todos los elementos (página actual)...")
        checkbox_todos = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "checkbox-comp-1"))
        )
        driver.execute_script("arguments[0].click();", checkbox_todos)
        print("[INFO] Clic en 'checkbox-comp-1' realizado.")
        if actualizar_estado:
            actualizar_estado("[INFO] Checkbox principal clickeado.")
        time.sleep(1)

        # 3. Clic en el botón "Select all X records"
        print("[INFO] Intentando hacer clic en el botón 'Select all ... records'...")
        if actualizar_estado:
            actualizar_estado("[PROGRESO] Seleccionando todos los registros coincidentes...")
        selector_select_all_domain = "button.o_list_select_domain[title='Select all records matching the search']"
        boton_seleccionar_todo_dominio = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, selector_select_all_domain))
        )
        boton_seleccionar_todo_dominio.click()
        print("[INFO] Clic en 'Select all ... records' realizado.")
        if actualizar_estado:
            actualizar_estado("[INFO] 'Select all ... records' clickeado.")
        time.sleep(1)

        # 4. Clic en el botón/opción de menú "Export"
        print("[INFO] Intentando hacer clic en la opción 'Export' del menú...")
        if actualizar_estado:
            actualizar_estado("[PROGRESO] Accediendo a la opción de Exportar...")

        action_button_selectors = [
            "//button[i[contains(@class, 'fa-cog')]]"
        ]
        action_button_clicked = False
        for i, selector in enumerate(action_button_selectors):
            try:
                action_button = WebDriverWait(driver, 7).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                print(f"[INFO] Botón 'Acción' encontrado con el selector {i + 1}. Haciendo clic...")
                action_button.click()
                action_button_clicked = True
                time.sleep(0.7)
                break
            except TimeoutException:
                print(f"[DEBUG] Botón 'Acción' no encontrado o no clickeable con el selector {i + 1} XPATH: {selector}")
                continue

        if not action_button_clicked:
            print(
                "[WARN] No se encontró un botón 'Acción' clickeable. Se intentará hacer clic directamente en 'Export'. Es posible que este paso falle si 'Export' está en un menú oculto.")

        selector_export_menu_item = "//span[contains(@class, 'o-dropdown-item') and contains(@class, 'o_menu_item') and normalize-space(.)='Export']"
        opcion_exportar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, selector_export_menu_item))
        )
        opcion_exportar.click()
        print("[INFO] Clic en la opción 'Export' del menú realizado.")
        if actualizar_estado:
            actualizar_estado("[INFO] Opción 'Export' del menú clickeada.")
        time.sleep(1.5)

        # 5. Clic en el botón "Export" final (en el modal)
        print("[INFO] Intentando hacer clic en el botón 'Export' final (del modal)...")
        if actualizar_estado:
            actualizar_estado("[PROGRESO] Confirmando la exportación...")
        selector_final_export_button = "button.btn.btn-primary.o_select_button[data-hotkey='v']"
        boton_exportar_final = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, selector_final_export_button))
        )
        boton_exportar_final.click()
        print("[INFO] Clic en el botón 'Export' final realizado. La descarga debería iniciarse.")
        if actualizar_estado:
            actualizar_estado(
                "[INFO] ¡Exportación iniciada! (El archivo se descargará según la configuración del navegador).")

        time.sleep(5)

    # ... (bloques except y finally como estaban) ...
    except ElementNotInteractableException as enie:
        print(f"[ERROR] ElementNotInteractable: El elemento no es interactuable. {enie}")
        driver.save_screenshot("error_element_not_interactable.png")
        if actualizar_estado:
            actualizar_estado(f"[ERROR] Elemento no interactuable: {enie}")
    except TimeoutException as te:
        print(f"[ERROR] Timeout: Un elemento no fue encontrado o no estuvo disponible a tiempo: {te}")
        current_url_on_timeout = driver.current_url
        print(f"[DEBUG] URL actual en el momento del Timeout: {current_url_on_timeout}")
        driver.save_screenshot("error_timeout_general.png")
        if actualizar_estado:
            actualizar_estado(f"[ERROR] Timeout: {te}")
    except NoSuchElementException as nse:
        print(f"[ERROR] NoSuchElement: No se pudo encontrar un elemento: {nse}")
        current_url_on_no_such_element = driver.current_url
        print(f"[DEBUG] URL actual en el momento del NoSuchElement: {current_url_on_no_such_element}")
        driver.save_screenshot("error_no_such_element.png")
        if actualizar_estado:
            actualizar_estado(f"[ERROR] Elemento no encontrado: {nse}")
    except Exception as e:
        print(f"[ERROR] Ocurrió un error inesperado: {e}")
        current_url_on_exception = driver.current_url
        print(f"[DEBUG] URL actual en el momento de la Excepción: {current_url_on_exception}")
        driver.save_screenshot("error_inesperado.png")
        if actualizar_estado:
            actualizar_estado(f"[ERROR] Error inesperado: {e}")
    finally:
        if 'driver' in locals() and driver:
            # driver.quit()
            print(
                "[INFO] Proceso de Selenium finalizado. El navegador permanecerá abierto (descomenta driver.quit() para cerrarlo).")
            if actualizar_estado:
                actualizar_estado("[INFO] Proceso finalizado.")
        else:
            print("[INFO] WebDriver no se inicializó.")

# 🧪 USO
if __name__ == "__main__":
    if configurar_descarga_si_nueva("http://smartrekking.com:8069/"):

        odoo_user = "admin"
        odoo_password = "admin"

        def imprimir_estado(mensaje):
            print(f"ESTADO_GUI: {mensaje}")

        print("Iniciando el bot de exportación de SmartTrekking Odoo...")
        realizar_exportacion_smartrekking(username=odoo_user, password=odoo_password, actualizar_estado=imprimir_estado)
