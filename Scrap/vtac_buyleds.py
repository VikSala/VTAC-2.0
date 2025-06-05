import os
import subprocess
import time
import traceback
from datetime import datetime

from appium import webdriver
from appium.options.android import UiAutomator2Options
from appium.webdriver.common.appiumby import AppiumBy
from selenium.common import NoSuchElementException, StaleElementReferenceException, WebDriverException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# The scraper should be able to:
# 1. Start the scraping process
# 2. Get the stock data for all SKUS present in ODOO16
# 3. Save the stock data to a json named 'buyled_stocks_{date}.json' in the output_dir_path

# Al inicio del archivo vtac_buyleds.py, después de los imports existentes

# ✅ Función principal pública para integración con el controller
def start_scrape(progress_callback=None):
    """
    Función principal para iniciar el scraping de BuyLED
    Compatible con el sistema de callbacks de progreso
    """
    if progress_callback:
        # Configurar el logger para usar el callback de progreso
        import logging
        logger = logging.getLogger("BuyLED")
        logger.setLevel(logging.INFO)

        class ProgressHandler(logging.Handler):
            def emit(self, record):
                progress_callback(self.format(record))

        handler = ProgressHandler()
        handler.setFormatter(logging.Formatter('%(message)s'))
        logger.handlers.clear()  # Limpiar handlers existentes
        logger.addHandler(handler)

        ScraperBuyLedStocks.logger = logger

    # Ejecutar el scraper
    return ScraperBuyLedStocks.start_scrape()


class ScraperBuyLedStocks():
    logger = None

    capabilities = dict(
        platformName='Android',
        automationName='UiAutomator2',
        deviceName='Pixel_8_Pro_API_28_3',
        appPackage='es.buyled.buyledpro',
        appActivity='.MainActivity t24',
        newCommandTimeout=600,
        noReset=False
    )

    BEGIN_SCRAPE_FROM = 3600
    DUMP_FREQUENCY = 50

    appium_server_url = 'http://localhost:4723'
    driver = None

    search_field_locator = 'className("android.widget.EditText")'
    btn_locator = 'className("android.widget.Button")'
    search_btn_locator = 'className("android.widget.Button").index(0)'

    # Buy led stock index = 0 ; ITA stock index = 1
    stock_buyled_locator = 'className("android.view.View").index(7)'
    stock_ita_locator = 'className("android.view.View").index(9)'
    price_locator = 'className("android.view.View").index(11)'
    date_locator = 'className("android.view.View").index(13)'

    login_fields_locator = 'className("android.widget.EditText")'
    email_field_index = 0
    password_field_index = 1

    PRODUCTS_INFO_PATH = 'data/buyled_stocks'

    @classmethod
    def start_scrape(cls):
        import pandas as pd
        productos = cls.cargar_productos_excel()

        cls.start_appium_server()
        cls.start_emulator()
        cls.wait_for_emulator_to_be_ready(cls.get_emulator_device_id(), os.path.join(os.environ['ANDROID_HOME'], 'platform-tools', 'adb'))

        cls.logger.info(f"Empezando extracción de datos para {len(productos)} productos...")
        cls.connect()

        stock_data = []
        acum = 0

        for index, product in enumerate(productos[cls.BEGIN_SCRAPE_FROM:]):
            sku = product.get("default_code")
            if not sku:
                continue

            fetched_data = cls.obtener_datos_seguro(sku)

            if fetched_data:
                stock_data.append(fetched_data)
                acum += 1
                cls.logger.info(f'{acum + cls.BEGIN_SCRAPE_FROM}. {fetched_data}')
            else:
                cls.logger.info(f" SKU omitido por falta de datos: {sku}")

        if stock_data:
            df_result = pd.DataFrame(stock_data)
            timestamp = datetime.now().strftime("%Y%m%d")
            output_file = f"buyled_{timestamp}.xlsx"
            df_result.to_excel(output_file, index=False)
            cls.logger.info(f"Exportado a '{output_file}' con {len(stock_data)} registros.")
        else:
            cls.logger.info("No se obtuvo ningún dato para exportar.")

        cls.end_scrape()

    @classmethod
    def cargar_productos_excel(cls):
        import pandas as pd
        from tkinter import filedialog
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Excel con productos",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if not archivo:
            raise FileNotFoundError("No se seleccionó ningún archivo.")
        df = pd.read_excel(archivo)
        return df[df['x_marca'] == 'V-TAC'].to_dict(orient='records')

    @classmethod
    def obtener_datos_seguro(cls, sku, reintentos=1):
        try:
            if not cls.search_sku(sku):
                cls.restart_android_emulator()
                cls.search_sku(sku)
            return cls.get_stock_data(sku)
        except Exception as e:
            cls.logger.warning(f"⚠️ Error con SKU {sku}: {e}\n{traceback.format_exc()}")
            if reintentos > 0:
                cls.restart_android_emulator()
                return cls.obtener_datos_seguro(sku, reintentos - 1)
            return None

    """def start_scrape(cls):
        cls.start_appium_server()
        cls.start_emulator()

        cls.logger.info('Empezando extracción de datos de ODOO. Espere unos minutos...')
        products_odoo = OdooImport.browse_all_products_in_batches('product_brand_id', '=', OdooImport.VTAC_BRAND_ID)
        cls.logger.info('Finalizada la extración de datos de ODOO.')
        stock_data = []
        # cls.driver = webdriver.Remote(cls.appium_server_url,
        #                               options=UiAutomator2Options().load_capabilities(cls.capabilities))
        # time.sleep(5)
        #
        # cls.login()
        cls.connect()
        acum = 0

        for index, product in enumerate(products_odoo[cls.BEGIN_SCRAPE_FROM:]):
            sku = product.default_code
            if not sku:
                continue

            try:
                no_error = cls.search_sku(sku)
                if not no_error:
                    cls.restart_android_emulator()
                    cls.search_sku(sku)
                fetched_data = cls.get_stock_data(sku)
            except (StaleElementReferenceException, WebDriverException) as e:
                # cls.logger.info(e.__class__, e.__str__())
                cls.restart_android_emulator()
                cls.search_sku(sku)
                fetched_data = cls.get_stock_data(sku)

            # cls.search_sku(sku)
            # fetched_data = cls.get_stock_data(sku)

            if fetched_data:
                stock_data.append(fetched_data)
                acum += 1
                cls.logger.info(f'{acum + cls.BEGIN_SCRAPE_FROM}. {fetched_data}')
            else:
                cls.logger.info(f"Skipping SKU: {sku} due to lack of data.")

            if len(stock_data) >= cls.DUMP_FREQUENCY or index == len(products_odoo) - 1:
                Util.dump_to_json(stock_data, f'{cls.PRODUCTS_INFO_PATH}/buyled_stocks_{acum + cls.BEGIN_SCRAPE_FROM}.json')
                stock_data = []
        cls.end_scrape()"""

    @classmethod
    def end_scrape(cls):
        cls.driver.quit()
        cls.stop_appium_server()
        cls.stop_emulator()

    @classmethod
    def connect(cls):
        cls.driver = webdriver.Remote(cls.appium_server_url,
                                      options=UiAutomator2Options().load_capabilities(cls.capabilities))
        time.sleep(5)

        try:
            cls.login()
        except Exception as e:
            cls.logger.warning(f"⚠️ Error al hacer login: {e}. Intentando reinicio de app.")
            cls.restart_android_emulator()
            cls.login()

    @classmethod
    def get_stock_data(cls, sku):
        try:
            wait = WebDriverWait(cls.driver, 1)  # Espera hasta 5 segundos antes de lanzar una excepción
            stock_buyled_text = wait.until(EC.visibility_of_element_located(
                (AppiumBy.ANDROID_UIAUTOMATOR, cls.stock_buyled_locator))).get_attribute('content-desc')
            stock_ita_text = cls.driver.find_element(AppiumBy.ANDROID_UIAUTOMATOR, cls.stock_ita_locator).get_attribute(
                'content-desc')
            price_text = cls.driver.find_element(AppiumBy.ANDROID_UIAUTOMATOR, cls.price_locator).get_attribute(
                'content-desc')

            try:
                date_text = cls.driver.find_element(AppiumBy.ANDROID_UIAUTOMATOR, cls.date_locator).get_attribute(
                    'content-desc')
            except NoSuchElementException:
                date_text = None

            return {
                'SKU': sku,
                'stock_buyled': int(stock_buyled_text),
                'stock_ita': int(stock_ita_text),
                'price': price_text,
                'date': date_text
            }
        except (NoSuchElementException, TimeoutException):
            cls.logger.info(f'No se encontró el dato de stock para SKU: {sku}. Omitiendo este SKU.')
            return None

    @classmethod
    def search_sku(cls, sku, retries=0):
        try:
            # cls.handle_app_not_responding()
            search_field_element = cls.driver.find_element(AppiumBy.ANDROID_UIAUTOMATOR, cls.search_field_locator)
            search_field_element.click()
            time.sleep(0.2)
            search_field_element.send_keys(sku)
            cls.driver.find_element(AppiumBy.ANDROID_UIAUTOMATOR, cls.search_btn_locator).click()
            time.sleep(1)
            return True
        except (NoSuchElementException, TimeoutException):
            if retries >= 2:
                return False

            cls.logger.info('Search field not found. Retrying...')
            time.sleep(2)
            return cls.search_sku(sku, retries=retries + 1)

    @classmethod
    def handle_app_not_responding(cls):
        try:
            # El texto del botón podría variar según el idioma del dispositivo o la versión del sistema operativo
            wait_button = cls.driver.find_element(AppiumBy.ANDROID_UIAUTOMATOR, 'new UiSelector().text("ESPERAR")')
            if wait_button:
                wait_button.click()
                cls.logger.info("Dialog 'App not responding' found. Clicked WAIT.")
        except NoSuchElementException:
            # Si no se encuentra el diálogo, simplemente continua
            pass

    @classmethod
    def login(cls):
        wait = WebDriverWait(cls.driver, 10)  # Ajusta el tiempo de espera según necesidad

        # Espera explícita para el campo de email
        email_login_element = wait.until(EC.element_to_be_clickable(
            (AppiumBy.ANDROID_UIAUTOMATOR, f'{cls.login_fields_locator}.instance({cls.email_field_index})')))
        email_login_element.click()
        time.sleep(5)  # Espera adicional antes de enviar el texto
        email_login_element.send_keys('compras@optimaluz.com')

        # Espera explícita para el campo de contraseña
        password_login_element = wait.until(EC.element_to_be_clickable(
            (AppiumBy.ANDROID_UIAUTOMATOR, f'{cls.login_fields_locator}.instance({cls.password_field_index})')))
        password_login_element.click()
        time.sleep(5)  # Espera adicional antes de enviar el texto
        password_login_element.send_keys('compras@optimaluz.com')

        # Espera explícita para el botón de login y luego clic
        login_button = wait.until(
            EC.element_to_be_clickable((AppiumBy.ANDROID_UIAUTOMATOR, f'{cls.btn_locator}.description("Entrar")')))
        login_button.click()

        cls.logger.info('Logged in BuyLed app')
        time.sleep(
            5)  # Espera adicional después de hacer clic para asegurar que la app haya tenido tiempo de reaccionar

    @classmethod
    def restart_android_emulator(cls, restart_connection=True):
        cls.logger.info('ERROR INESPERADO PROCEDIENDO A REINICIAR EMULADOR')
        # cls.driver.quit()
        # Rutas completas a adb y emulator
        adb_path = os.path.join(os.environ['ANDROID_HOME'], 'platform-tools', 'adb')
        emulator_path = os.path.join(os.environ['ANDROID_HOME'], 'emulator', 'emulator')

        # Obtener la lista de dispositivos y encontrar el identificador del emulador
        device_id = cls.get_emulator_device_id(adb_path)

        if device_id is None:
            cls.logger.info("Emulador no encontrado")
            return

        # Comando para apagar el emulador
        cls.stop_emulator(adb_path, device_id)
        # Espera a que el emulador se apague completamente
        time.sleep(15)  # Usar time.sleep para Windows

        # Comando para iniciar el emulador
        # subprocess.Popen([emulator_path, "-avd", cls.capabilities['deviceName'], "-no-snapshot-load"])
        cls.start_emulator(emulator_path=emulator_path, adb_path=adb_path)
        cls.wait_for_emulator_to_be_ready(device_id, adb_path)
        cls.logger.info('REINICIO COMPLETO')
        cls.restart_appium_server()
        if restart_connection:
            cls.connect()

    @classmethod
    def stop_emulator(cls, adb_path=None, device_id=None):
        if not adb_path:
            adb_path = cls.get_adb_path()
        if not device_id:
            device_id = cls.get_emulator_device_id(adb_path)
        subprocess.run([adb_path, "-s", device_id, "emu", "kill"], capture_output=True)
        cls.logger.info('EMULADOR APAGADO')

    @classmethod
    def get_emulator_device_id(cls, adb_path=None):
        if not adb_path:
            adb_path = cls.get_adb_path()
        result = subprocess.run([adb_path, "devices"], capture_output=True, text=True)
        print(result.stdout)  # Imprimir dispositivos para verificar
        device_id = None
        for line in result.stdout.splitlines():
            if 'emulator' in line:
                device_id = line.split()[0]  # Asume que el ID del dispositivo está en la primera columna
                break
        return device_id

    @classmethod
    def get_adb_path(cls):
        return os.path.join(os.environ['ANDROID_HOME'], 'platform-tools', 'adb')

    @classmethod
    def start_emulator(cls, emulator_path=None, adb_path=None):
        if not emulator_path:
            emulator_path = os.path.join(os.environ['ANDROID_HOME'], 'emulator', 'emulator')
        if not adb_path:
            adb_path = cls.get_adb_path()
        if cls.get_emulator_device_id(adb_path):
            cls.logger.info('EL EMULADOR ESTABA ENCENDIDO, PROCEDIEDO A REINICIAR')
            cls.restart_android_emulator(restart_connection=False)
        subprocess.Popen([emulator_path, "-avd", cls.capabilities['deviceName'], "-no-snapshot-load"])
        cls.logger.info('EMULADOR ENCENDIDO')
        time.sleep(60)

    @classmethod
    def stop_appium_server(cls, port=4723):
        """ Detiene el servidor de Appium en el puerto especificado. """
        cls.logger.info("Deteniendo el servidor de Appium...")
        try:
            subprocess.run(['taskkill', '/F', '/IM', 'node.exe'], capture_output=True, check=True)
        except subprocess.CalledProcessError:
            cls.logger.info('El servidor no estaba en ejecución')
        cls.logger.info("Servidor de Appium detenido.")

    @classmethod
    def start_appium_server(cls, port=4723):
        """ Inicia el servidor de Appium en el puerto especificado. """
        appium_path = os.path.join(os.environ['USERPROFILE'], 'AppData', 'Roaming', 'npm', 'appium.cmd')
        cls.logger.info("Iniciando el servidor de Appium...")

        process = subprocess.Popen([appium_path, '-p', str(port)], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                   text=True)
        time.sleep(5)
        # Espera un momento para que Appium inicie
        try:
            stdout, stderr = process.communicate(timeout=10)  # Espera hasta 10 segundos para salida inicial
            # cls.logger.info("Salida de Appium:", stdout)
            # if stderr:
            #     cls.logger.info("Error de Appium:", stderr)
        except subprocess.TimeoutExpired:
            cls.logger.info("Appium ha comenzado pero no ha respondido con salida inicial en 10 segundos.")
        cls.logger.info("Servidor de Appium iniciado.")

    @classmethod
    def restart_appium_server(cls, port=4723):
        """ Reinicia el servidor de Appium en el puerto especificado. """
        cls.stop_appium_server(port)
        time.sleep(5)  # Espera adicional para asegurar que el servidor esté completamente detenido
        cls.start_appium_server(port)
        time.sleep(10)

    @classmethod
    def wait_for_emulator_to_be_ready(cls, device_id, adb_path):
        """ Espera hasta que el emulador esté listo para usar """
        while True:
            result = subprocess.run([adb_path, "devices"], capture_output=True, text=True)
            lines = result.stdout.splitlines()
            if any(device_id in line and "device" in line for line in lines):
                cls.logger.info("Emulador listo")
                break
            cls.logger.info("Esperando que el emulador esté listo...")
            time.sleep(5)

# ScraperBuyLedStocks.start_scrape('data/buyled_stocks')
