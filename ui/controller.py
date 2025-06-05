# controller.py
import tkinter as tk
import os
import threading
import time

from enum import Enum, auto
from core.events import EventManager


class ScrapStatus(Enum):
    IDLE = auto()
    RUNNING = auto()
    COMPLETED = auto()
    ERROR = auto()


class Controller:
    def __init__(self, master):
        self.master = master  # Esta es la ventana raíz de Tkinter
        self.scrap_status = ScrapStatus.IDLE
        self.event_manager = EventManager()

        self._current_operation_thread = None
        self._stop_current_operation = threading.Event()

        self.comparison_server_thread = None
        self.comparison_server_stop_event = threading.Event()

        self._initialize_controllers()

    def _initialize_controllers(self):
        class TkVarHolder:
            def __init__(self, master_tk):  # master_tk será self.master (la ventana raíz)
                # Scraper
                self.IF_EXTRACT_ITEM_INFO = tk.IntVar(master=master_tk, value=1)
                self.IF_ONLY_NEW_PRODUCTS = tk.IntVar(master=master_tk)
                self.SPAIN = tk.IntVar(master=master_tk)
                self.UK = tk.IntVar(master=master_tk)
                self.ITALIA = tk.IntVar(master=master_tk)
                self.BUYLED = tk.IntVar(master=master_tk)
                # Excel
                self.MEGA_EXCEL_PATH = tk.StringVar(master=master_tk)
                self.RESTAURAR_BACKUP = tk.BooleanVar(master=master_tk)
                self.GENERAR_NUEVO_BLOQUE = tk.BooleanVar(master=master_tk)
                self.OBTENER_DATOS_PRINCIPAL_HOJAS = tk.BooleanVar(master=master_tk)
                self.GENERAR_EXCEL_NUEVAS_UNIDADES = tk.BooleanVar(master=master_tk)
                self.RELLENAR_NUEVAS_UNIDADES = tk.BooleanVar(master=master_tk)
                self.RELLENAR_STOCKS = tk.BooleanVar(master=master_tk)
                self.CALCULAR_NUM_VENTAS = tk.BooleanVar(master=master_tk)
                self.CALCULAR_VALOR_VENTAS = tk.BooleanVar(master=master_tk)
                self.CALCULAR_VALOR_STOCK = tk.BooleanVar(master=master_tk)
                self.CALCULAR_TOTAL_VALORES_MONETARIOS = tk.BooleanVar(master=master_tk)
                self.RELLENAR_PRINCIPAL = tk.BooleanVar(master=master_tk)
                self.BULGARIA = tk.BooleanVar(master=master_tk)
                self.MADRID = tk.BooleanVar(master=master_tk)
                self.POLONIA = tk.BooleanVar(master=master_tk)
                # Merge
                self.SCRAPPED_DATA_EXCEL_PATH = tk.StringVar(master=master_tk)
                self.USER_SELECTIONS_PATH = tk.StringVar(master=master_tk, value=os.path.join('data', 'common', 'json',
                                                                                              'user_selections.json'))
                self.SKIP_SKUS_JSON_PATH = tk.StringVar(master=master_tk, value=os.path.join('data', 'common', 'json',
                                                                                             'SKUS_TO_SKIP.json'))
                # Nuevas variables para la comparación entre Excels Merge
                self.MERGED_EXCEL_PATH_OLD = tk.StringVar(master=master_tk)
                self.MERGED_EXCEL_PATH_NEW = tk.StringVar(master=master_tk)

                # Import Odoo (Variables actualizadas según solicitud)
                self.IF_IMPORT_PRODUCTS = tk.BooleanVar(master=master_tk, value=1);  # Mantener y por defecto activado
                self.IF_PUBLISH_PRODUCTS = tk.BooleanVar(master=master_tk);  # Nuevo checkbox
                self.IF_IMPORT_COMERCIAL_INFO = tk.BooleanVar(master=master_tk);  # Nuevo checkbox unificado

                # Las siguientes variables ya no son necesarias en la UI,
                # podrían eliminarse si la lógica de _run_import_task ya no las usa individualmente.
                # Por ahora las comento para referencia, pero idealmente se eliminarían si
                # IF_IMPORT_COMERCIAL_INFO las engloba en la lógica del backend.
                # self.IF_IMPORT_FIELDS = tk.BooleanVar(master=master_tk);
                # self.IF_IMPORT_PUBLIC_CATEGORIES = tk.BooleanVar(master=master_tk);
                # self.IF_IMPORT_BRANDS = tk.BooleanVar(master=master_tk);
                # self.IF_SKIP_EXISTING = tk.BooleanVar(master=master_tk);
                # self.IF_FORCE_UPDATE = tk.BooleanVar(master=master_tk);
                # self.IF_IMPORT_ACC = tk.BooleanVar(master=master_tk);
                # self.IF_IMPORT_IMGS_AND_VIDEOS = tk.BooleanVar(master=master_tk);
                # self.SKIP_PRODUCTS_W_MEDIA = tk.BooleanVar(master=master_tk);
                # self.IF_CLEAN_EXISTING_IMGS_AND_VIDEOS = tk.BooleanVar(master=master_tk);
                # self.IF_IMPORT_ICONS = tk.BooleanVar(master=master_tk);
                # self.IF_IMPORT_SPEC_SHEETS = tk.BooleanVar(master=master_tk);
                # self.IF_CLEAN_EXISTING_SPEC_SHEETS = tk.BooleanVar(master=master_tk);
                # self.IF_IMPORT_PDFS = tk.BooleanVar(master=master_tk);
                # self.SKIP_PRODUCTS_W_ATTACHMENTS = tk.BooleanVar(master=master_tk);
                # self.IF_CLEAN_EXISTING_PDFS = tk.BooleanVar(master=master_tk);
                # self.IF_IMPORT_SUPPLIER_INFO_AND_COST = tk.BooleanVar(master=master_tk); # Unificado
                # self.IF_UPDATE_MODE = tk.BooleanVar(master=master_tk); # Eliminado
                # self.IF_IMPORT_AVAILABILITY = tk.BooleanVar(master=master_tk); # Unificado
                # self.IF_IMPORT_LOCAL_STOCK = tk.BooleanVar(master=master_tk); # Unificado
                # self.IF_IMPORT_DESCATALOGADOS_CATALOGO = tk.BooleanVar(master=master_tk);
                # self.IF_GENERATE_MISSING_PRODUCTS_EXCEL = tk.BooleanVar(master=master_tk);
                # self.IF_ARCHIVE_PRODUCTS_FROM_JSON = tk.BooleanVar(master=master_tk)

        # self.master es la ventana raíz de Tkinter pasada al Controller
        self.scraper_controller = TkVarHolder(self.master)
        self.excel_controller = TkVarHolder(self.master)
        self.merge_controller = TkVarHolder(self.master)
        self.import_controller = TkVarHolder(
            self.master)  # Contendrá IF_IMPORT_PRODUCTS, IF_PUBLISH_PRODUCTS, IF_IMPORT_COMERCIAL_INFO

        os.makedirs(os.path.join('data', 'common', 'json'), exist_ok=True)
        os.makedirs(os.path.join('data', 'common', 'html_comparison'), exist_ok=True)

    def _update_status(self, status_enum, message=""):
        if status_enum == ScrapStatus.RUNNING:
            self.scrap_status = ScrapStatus.RUNNING
        elif status_enum == ScrapStatus.COMPLETED:
            self.scrap_status = ScrapStatus.COMPLETED
        elif status_enum == ScrapStatus.ERROR:
            self.scrap_status = ScrapStatus.ERROR
        else:
            self.scrap_status = ScrapStatus.IDLE
        self.event_manager.emit('status_update', self.scrap_status, message)

    def _start_long_operation(self, target_function, *args):
        if self.is_critical_operation_active():
            self.event_manager.emit('operation_error', "Ya hay otra operación crítica en curso.")
            return
        self._stop_current_operation.clear()
        self.comparison_server_stop_event.clear()

        def wrapper_function():
            try:
                target_function(*args)
            except Exception as e:
                if self._stop_current_operation.is_set():
                    self.event_manager.emit('progress_update', "Operación cancelada por el usuario.")
                    self.event_manager.emit('operation_completed', "Operación cancelada.")  # Mensaje genérico
                else:
                    self.event_manager.emit('operation_error', f"Error en la operación: {type(e).__name__} - {str(e)}")
                    import traceback
                    self.event_manager.emit('progress_update', f"Traceback: {traceback.format_exc()}")
            finally:
                self._current_operation_thread = None
                if self.comparison_server_thread and self.comparison_server_thread.is_alive():
                    print("Controller: Operación wrapper finalizada, asegurando parada del servidor Flask.")
                    self._stop_comparison_server()

        self._current_operation_thread = threading.Thread(target=wrapper_function, daemon=True)
        self._current_operation_thread.start()

    def is_critical_operation_active(self):
        return self._current_operation_thread is not None and self._current_operation_thread.is_alive()

    def cancel_current_operation(self):
        if self.is_critical_operation_active():
            self.event_manager.emit('progress_update', "Intentando cancelar la operación actual...")
            self._stop_current_operation.set()
            if self.comparison_server_thread and self.comparison_server_thread.is_alive():
                print("Controller: Cancelación solicitada, activando evento de parada del servidor Flask.")
                self.comparison_server_stop_event.set()
        else:
            self.event_manager.emit('progress_update', "No hay operación activa para cancelar.")

    #SCRAP
    def is_scraping_active(self):
        return self.scrap_status == ScrapStatus.RUNNING

    def execute_scraper_controller_actions(self):
        print("\n🚀 CONTROLLER: EJECUTANDO SCRAPING")
        if self.is_critical_operation_active():
            self.event_manager.emit('scraping_error', "Ya hay una operación en curso.")
            return
        self._start_long_operation(self._run_scraping_task)

    def _run_scraping_task(self):
        """try:
            self._update_status(ScrapStatus.RUNNING, "Iniciando scraping...")
            self.event_manager.emit('progress_update', "Validando configuración de scraping...")
            time.sleep(0.2)  # Simulación
            if not any([self.scraper_controller.ITALIA.get(), self.scraper_controller.UK.get(),
                        self.scraper_controller.SPAIN.get(), self.scraper_controller.BUYLED.get()]):
                raise ValueError("Debe seleccionar al menos una región para el scraping.")
            self.event_manager.emit('progress_update', "Configuración validada.")

            if self.scraper_controller.SPAIN.get():  # Simulación
                self.event_manager.emit('progress_update', "🇪🇸 Iniciando scraping de España...")
                for i in range(3):
                    if self._stop_current_operation.is_set():
                        self.event_manager.emit('progress_update', "Scraping España cancelado.")
                        # self.event_manager.emit('operation_completed', "Scraping cancelado.") # Ya manejado en wrapper
                        return
                    time.sleep(0.5)
                    self.event_manager.emit('progress_update', f"🇪🇸 Progreso España {i + 1}/3")
                self.event_manager.emit('progress_update', "🇪🇸 Scraping de España completado.")

            self._update_status(ScrapStatus.COMPLETED, "Scraping finalizado.")
            self.event_manager.emit('scraping_completed', "Scraping finalizado exitosamente.")
        except ValueError as ve:
            if not self._stop_current_operation.is_set():
                self.event_manager.emit('scraping_error', str(ve))
                self._update_status(ScrapStatus.ERROR, f"Error en scraping: {str(ve)}")
        except Exception as e:
            if not self._stop_current_operation.is_set():
                self.event_manager.emit('scraping_error', f"Error inesperado durante el scraping: {str(e)}")
                self._update_status(ScrapStatus.ERROR, f"Error en scraping: {str(e)}")"""
        """Versión que maneja múltiples regiones según configuración"""
        import threading
        import time

        print("🔥 DEBUG: scrap_it_uk() iniciada - MODO MULTI-REGIÓN")

        # ✅ OBTENER CONFIGURACIONES EN EL HILO PRINCIPAL
        try:
            scrap_italia = bool(self.scraper_controller.ITALIA.get())
            scrap_uk = bool(self.scraper_controller.UK.get())
            scrap_espana = bool(self.scraper_controller.SPAIN.get())
            scrap_buyled = bool(self.scraper_controller.BUYLED.get())
            only_new_products = bool(self.scraper_controller.IF_ONLY_NEW_PRODUCTS.get())

            print(f"🔥 DEBUG: Configuración obtenida:")
            print(f"   - Italia: {scrap_italia}")
            print(f"   - Reino Unido: {scrap_uk}")
            print(f"   - España: {scrap_espana}")
            print(f"   - BuyLED: {scrap_buyled}")
            print(f"   - Solo nuevos productos: {only_new_products}")

        except Exception as e:
            error_msg = f"Error obteniendo configuraciones: {str(e)}"
            print(f"🔥 DEBUG: {error_msg}")
            #if progress_callback: progress_callback(error_msg)
            return None

        # En controller.py, actualiza la función run_scraping():

        def run_scraping():
            try:
                print("🔥 DEBUG: run_scraping() iniciada")
                self._update_status(ScrapStatus.RUNNING, "Iniciando scraping...")

                # ✅ CALLBACK PARA PROGRESO
                def progress_callback_internal(message):
                    self.event_manager.emit('progress_update', message)
                    time.sleep(0.1)

                progress_callback_internal("Configuración validada")
                time.sleep(0.5)

                # Validar que al menos una región esté seleccionada
                if not any([scrap_italia, scrap_uk, scrap_espana, scrap_buyled]):
                    error_msg = "Error: Debe seleccionar al menos una región"
                    print(f"🔥 DEBUG: {error_msg}")
                    progress_callback_internal(error_msg)
                    raise ValueError("No hay regiones seleccionadas")

                progress_callback_internal("Importando módulos de scraping...")
                time.sleep(0.5)

                # ✅ IMPORTACIONES MEJORADAS CON MANEJO INDIVIDUAL
                print("🔥 DEBUG: Iniciando importaciones...")

                vtac_it = None
                vtac_uk = None
                vtac_buyleds = None

                try:
                    if scrap_italia:
                        from Scrap import vtac_it
                        print("🔥 DEBUG: vtac_it importado correctamente")
                except ImportError as e:
                    print(f"🔥 DEBUG: Error importando vtac_it: {e}")
                    if scrap_italia:
                        progress_callback_internal(f"Error importando módulo Italia: {e}")

                try:
                    if scrap_uk:
                        from Scrap import vtac_uk
                        print("🔥 DEBUG: vtac_uk importado correctamente")
                except ImportError as e:
                    print(f"🔥 DEBUG: Error importando vtac_uk: {e}")
                    if scrap_uk:
                        progress_callback_internal(f"Error importando módulo Reino Unido: {e}")

                try:
                    if scrap_buyled:
                        print("🔥 DEBUG: Intentando importar vtac_buyleds directamente...")
                        from Scrap import vtac_buyleds
                        print(f"🔥 DEBUG: vtac_buyleds importado directamente: {vtac_buyleds}")

                except ImportError as e:
                    print(f"🔥 DEBUG: Error importando vtac_buyleds: {e}")
                    if scrap_buyled:
                        progress_callback_internal(f"Error importando módulo BuyLED: {e}")
                        vtac_buyleds = None

                progress_callback_internal("Módulos importados correctamente")
                time.sleep(0.5)

                # ✅ SCRAPING ITALIA
                if scrap_italia and vtac_it:
                    print("🔥 DEBUG: Iniciando proceso REAL Italia")
                    progress_callback_internal("Iniciando scraping de Italia...")
                    time.sleep(1)

                    try:
                        def italia_progress(message):
                            progress_callback_internal(f"{message}")

                        vtac_it.run_vtac_italia_scraper(
                            modo_scrap=vtac_it.ScrapMode.COMPLETO,
                            only_new_products=only_new_products,
                            progress_callback=italia_progress
                        )
                        progress_callback_internal("Scraping de Italia completado")

                    except Exception as e:
                        error_msg = f"Error en scraping de Italia: {e}"
                        print(f"🔥 DEBUG: {error_msg}")
                        progress_callback_internal(error_msg)

                elif scrap_italia:
                    progress_callback_internal("Italia seleccionado pero módulo no disponible")

                # ✅ SCRAPING REINO UNIDO
                if scrap_uk and vtac_uk:
                    print("🔥 DEBUG: Iniciando proceso REAL UK")
                    progress_callback_internal("Iniciando scraping de Reino Unido...")
                    time.sleep(1)

                    try:
                        def uk_progress(message):
                            progress_callback_internal(f"{message}")

                        vtac_uk.run_vtac_uk_scraper(
                            modo_scrap=vtac_uk.ScrapMode.COMPLETO,
                            only_new_products=only_new_products,
                            progress_callback=uk_progress
                        )
                        progress_callback_internal("Scraping de Reino Unido completado")

                    except Exception as e:
                        error_msg = f"Error en scraping de Reino Unido: {e}"
                        print(f"🔥 DEBUG: {error_msg}")
                        progress_callback_internal(error_msg)

                elif scrap_uk:
                    progress_callback_internal("Reino Unido seleccionado pero módulo no disponible")

                # ✅ SCRAPING ESPAÑA - IMPLEMENTADO CON SCRAPY
                # En controller.py, reemplaza la sección de España por esto:

                # ✅ SCRAPING ESPAÑA - IMPLEMENTADO CON SCRAPY
                # En controller.py, reemplaza la sección de España por esto:

                # ✅ SCRAPING ESPAÑA - IMPLEMENTADO CON SCRAPY
                # En controller.py, reemplaza la sección de España por esto:

                # ✅ SCRAPING ESPAÑA - IMPLEMENTADO CON SCRAPY
                # En controller.py, reemplaza toda la sección de España por esto:

                # ✅ SCRAPING ESPAÑA - IMPLEMENTADO CON SCRAPY (VERSIÓN FINAL)
                # En controller.py, reemplaza toda la sección de España por esto:

                # ✅ SCRAPING ESPAÑA - IMPLEMENTADO CON SCRAPY (VERSIÓN FINAL)
                # En controller.py, reemplaza la sección de España por esto:

                # ✅ SCRAPING ESPAÑA - IMPLEMENTADO CON SCRAPY (CON PYTHONPATH CORREGIDO)
                if scrap_espana:
                    print("🔥 DEBUG: Iniciando proceso REAL España")
                    progress_callback_internal("Iniciando scraping de España...")
                    time.sleep(1)

                    try:
                        import subprocess
                        import os
                        import sys

                        def espana_progress(message):
                            progress_callback_internal(f"{message}")

                        espana_progress("Preparando entorno Scrapy...")

                        # ✅ DIRECTORIO SCRAPY CORRECTO
                        current_dir = os.getcwd()
                        parent_dir = os.path.dirname(current_dir)
                        scrapy_project_dir = os.path.join(parent_dir, "Scrap")

                        print(f"🔥 DEBUG: Directorio Scrapy: {scrapy_project_dir}")

                        # ✅ VERIFICACIÓN
                        scrapy_cfg_path = os.path.join(scrapy_project_dir, "scrapy.cfg")
                        settings_path = os.path.join(scrapy_project_dir, "Scrap", "settings.py")

                        if not os.path.exists(scrapy_cfg_path):
                            error_msg = "scrapy.cfg no encontrado"
                            espana_progress(error_msg)
                            return

                        if not os.path.exists(settings_path):
                            error_msg = "settings.py no encontrado"
                            espana_progress(error_msg)
                            return

                        print("🔥 DEBUG: ✅ Archivos verificados")
                        espana_progress("Proyecto Scrapy verificado")

                        # ✅ CONFIGURAR VARIABLES DE ENTORNO CON PYTHONPATH
                        env = os.environ.copy()
                        env['PYTHONUNBUFFERED'] = '1'

                        # ✅ AGREGAR EL DIRECTORIO DEL PROYECTO AL PYTHONPATH
                        current_pythonpath = env.get('PYTHONPATH', '')
                        if current_pythonpath:
                            env['PYTHONPATH'] = f"{scrapy_project_dir}{os.pathsep}{current_pythonpath}"
                        else:
                            env['PYTHONPATH'] = scrapy_project_dir

                        print(f"🔥 DEBUG: PYTHONPATH configurado: {env['PYTHONPATH']}")

                        # ✅ EJECUTAR SCRAPY
                        espana_progress("Iniciando spider 'vtac' para región España...")
                        print("🔥 DEBUG: Ejecutando: scrapy crawl vtac -a region=vtac_es")

                        process = subprocess.Popen(
                            ["scrapy", "crawl", "vtac", "-a", "region=vtac_es"],
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                            text=True,
                            universal_newlines=True,
                            bufsize=1,
                            cwd=scrapy_project_dir,
                            env=env  # ✅ USAR ENTORNO MODIFICADO
                        )

                        # ✅ LEER SALIDA (código igual que antes)
                        def read_output(pipe, label):
                            try:
                                for line in iter(pipe.readline, ''):
                                    if line.strip():
                                        line_clean = line.strip()

                                        if any(keyword in line_clean for keyword in [
                                            'INFO] Scrapy',
                                            'INFO] Crawled',
                                            'INFO] Spider opened',
                                            'INFO] Spider closed',
                                            'DEBUG]',
                                            'Traceback',
                                            'ERROR]'
                                        ]):
                                            if 'ERROR]' in line_clean or 'Traceback' in line_clean:
                                                espana_progress(f"{line_clean}")
                                            elif 'Spider opened' in line_clean:
                                                espana_progress("Spider iniciado")
                                            elif 'Spider closed' in line_clean:
                                                espana_progress("Spider finalizado")
                                            elif 'Crawled' in line_clean and 'pages' in line_clean:
                                                espana_progress(f"{line_clean}")

                                            print(f"🔥 DEBUG España [{label}]: {line_clean}")
                                        else:
                                            print(f"🔥 DEBUG España [{label}]: {line_clean}")

                            except Exception as e:
                                print(f"🔥 DEBUG: Error leyendo {label}: {e}")
                            finally:
                                pipe.close()

                        # ✅ THREADS PARA LECTURA
                        '''import threading
                        stdout_thread = threading.Thread(target=read_output, args=(process.stdout, "OUT"))
                        stderr_thread = threading.Thread(target=read_output, args=(process.stderr, "ERR"))

                        stdout_thread.daemon = True
                        stderr_thread.daemon = True

                        stdout_thread.start()
                        stderr_thread.start()'''

                        # ✅ MONITOREO DEL PROCESO
                        espana_progress("Scraping en progreso...")

                        try:
                            return_code = process.wait(timeout=1800)  # 30 minutos
                        except subprocess.TimeoutExpired:
                            espana_progress("Timeout: Proceso tardó más de 30 minutos")
                            process.kill()
                            return_code = -1

                        #stdout_thread.join(timeout=5)
                        #stderr_thread.join(timeout=5)

                        print(f"🔥 DEBUG: Proceso terminado con código: {return_code}")

                        # ✅ EVALUAR RESULTADO
                        if return_code == 0:
                            espana_progress("✅ Scraping completado exitosamente")

                            try:
                                output_dir = os.path.expanduser("~/Documents/SMI Files")

                                if os.path.exists(output_dir):
                                    espana_progress("Abriendo directorio de resultados...")

                                    import platform
                                    try:
                                        if platform.system() == "Windows":
                                            subprocess.run(["explorer", output_dir], check=False)

                                        espana_progress(f"Resultados en: {output_dir}")

                                    except Exception:
                                        espana_progress(f"Resultados disponibles en: {output_dir}")
                                else:
                                    espana_progress("Directorio de resultados no encontrado")

                            except Exception as e:
                                espana_progress(f"Error manejando resultados: {e}")

                            progress_callback_internal("Scraping de España completado exitosamente")

                        else:
                            error_msg = f"Scraping terminó con errores (código: {return_code})"
                            espana_progress(error_msg)
                            progress_callback_internal(error_msg)

                    except FileNotFoundError:
                        error_msg = "Scrapy no encontrado. Instale con: pip install scrapy"
                        espana_progress(error_msg)
                        progress_callback_internal(error_msg)

                    except Exception as e:
                        error_msg = f"Error inesperado: {e}"
                        espana_progress(error_msg)
                        print(f"🔥 DEBUG: Error España: {e}")
                        import traceback
                        print(f"🔥 DEBUG: Traceback: {traceback.format_exc()}")
                        progress_callback_internal(error_msg)

                # ✅ SCRAPING BUYLED
                if scrap_buyled and vtac_buyleds:
                    print("🔥 DEBUG: Iniciando proceso REAL BuyLED")
                    progress_callback_internal("Iniciando scraping de BuyLED...")
                    time.sleep(1)

                    try:
                        if hasattr(vtac_buyleds, 'ScraperBuyLedStocks'):
                            scraper_class = vtac_buyleds.ScraperBuyLedStocks

                            if hasattr(scraper_class, 'start_scrape'):
                                def buyled_progress(message):
                                    progress_callback_internal(f"{message}")

                                # ✅ CONFIGURAR EL LOGGER PARA BUYLED
                                import logging
                                logger = logging.getLogger("BuyLED")
                                logger.setLevel(logging.INFO)

                                class ProgressHandler(logging.Handler):
                                    def emit(self, record):
                                        buyled_progress(self.format(record))

                                handler = ProgressHandler()
                                handler.setFormatter(logging.Formatter('%(message)s'))
                                logger.handlers.clear()
                                logger.addHandler(handler)

                                scraper_class.logger = logger
                                scraper_class.start_scrape()

                                progress_callback_internal("Scraping de BuyLED completado")

                            else:
                                progress_callback_internal("Método start_scrape no encontrado en ScraperBuyLedStocks")

                        else:
                            progress_callback_internal("Clase ScraperBuyLedStocks no encontrada en vtac_buyleds")

                    except Exception as e:
                        error_msg = f"Error en scraping de BuyLED: {e}"
                        print(f"🔥 DEBUG: {error_msg}")
                        progress_callback_internal(error_msg)

                elif scrap_buyled:
                    progress_callback_internal("BuyLED seleccionado pero módulo no disponible")

                # Finalización
                print("🔥 DEBUG: Finalizando proceso REAL...")
                success_message = "Scraping finalizado exitosamente"
                self._update_status(ScrapStatus.COMPLETED, success_message)

                progress_callback_internal(success_message)
                self.event_manager.emit('scraping_completed', success_message)
                print("🔥 DEBUG: Proceso REAL completado exitosamente")

            except Exception as e:
                error_msg = f"Error durante el scraping: {str(e)}"
                print(f"🔥 DEBUG: ERROR CAPTURADO REAL: {error_msg}")
                import traceback
                print(f"🔥 DEBUG: Traceback: {traceback.format_exc()}")

                self._update_status(ScrapStatus.ERROR, error_msg)
                self.event_manager.emit('progress_update', error_msg)
                self.event_manager.emit('scraping_error', error_msg)
                raise

        # Ejecutar en thread separado
        print("🔥 DEBUG: Creando thread principal...")
        thread = threading.Thread(target=run_scraping)
        thread.daemon = True

        print("🔥 DEBUG: Iniciando thread principal...")
        thread.start()

        print("🔥 DEBUG: Thread iniciado, retornando...")
        return thread

    #MERGE
    def execute_excel_controller_actions(self):
        print("📊 CONTROLLER: Ejecutando acciones de Excel...")
        if self.is_critical_operation_active():
            self.event_manager.emit('operation_error', "Ya hay una operación en curso.")
            return

        # Obtener valores de Tkinter Vars en el hilo principal
        mega_excel_path_value = self.excel_controller.MEGA_EXCEL_PATH.get()
        generar_nuevo_bloque_value = self.excel_controller.GENERAR_NUEVO_BLOQUE.get()
        # ... (obtener otros valores de self.excel_controller si son necesarios en _run_excel_task)

        self._start_long_operation(self._run_excel_task, mega_excel_path_value, generar_nuevo_bloque_value)

    def _run_excel_task(self, mega_excel_path, generar_nuevo_bloque):  # Aceptar valores como args
        try:
            self.event_manager.emit('progress_update', "Iniciando proceso de Excel...")
            # path = self.excel_controller.MEGA_EXCEL_PATH.get() # NO HACER ESTO AQUÍ
            if not mega_excel_path: raise ValueError("Ruta del Mega Excel no especificada.")
            self.event_manager.emit('progress_update', f"Procesando Mega Excel: {mega_excel_path}")
            time.sleep(0.5)  # Simulación
            # if self.excel_controller.GENERAR_NUEVO_BLOQUE.get(): # NO HACER ESTO AQUÍ
            if generar_nuevo_bloque:
                self.event_manager.emit('progress_update', "Generando nuevo bloque...")
                time.sleep(1)  # Simulación
            self.event_manager.emit('operation_completed', "Proceso de Excel completado.")
        except Exception as e:
            if not self._stop_current_operation.is_set():
                self.event_manager.emit('operation_error', f"Error en proceso de Excel: {str(e)}")

    #IMPORT
    def execute_import_controller_action(self):
        print("📥 CONTROLLER: Ejecutando acciones de importación...")
        if self.is_critical_operation_active():
            self.event_manager.emit('operation_error', "Ya hay una operación en curso.")
            return

        # Obtener valores de Tkinter Vars aquí si _run_import_task los necesita
        if_import_products_value = self.import_controller.IF_IMPORT_PRODUCTS.get()
        if_publish_products_value = self.import_controller.IF_PUBLISH_PRODUCTS.get()
        if_import_comercial_info_value = self.import_controller.IF_IMPORT_COMERCIAL_INFO.get()

        # Aquí necesitarás pasar estas variables a _run_import_task y usar sus valores.
        # Por ahora, solo paso if_import_products_value como ejemplo.
        # Deberás actualizar _run_import_task para manejar las nuevas variables.
        self._start_long_operation(self._run_import_task,
                                   if_import_products_value,
                                   if_publish_products_value,
                                   if_import_comercial_info_value)

    def _run_import_task(self, if_import_products, if_publish_products,
                         if_import_comercial_info):
        from core import App_Import
        try:
            if if_import_products:
                self.event_manager.emit('progress_update', "Importando productos...")

                # Aquí va la lógica real para importar productos
                App_Import.main_import()

                if if_publish_products:
                    self.event_manager.emit('progress_update', "ToDO: Publicando productos importados...")
                    # Aquí iría la lógica real para publicar productos
                    time.sleep(0.5)  # Simulación

            if if_import_comercial_info:
                self.event_manager.emit('progress_update', "ToDO: Importando información comercial...")
                # Aquí iría la lógica real para:
                # - Importar información del proveedor y coste
                # - Importar disponibilidad
                # - Importar stock local
                time.sleep(1.5)  # Simulación

            # Simular otras acciones basadas en las variables que se hayan mantenido o la lógica interna
            # Por ejemplo, si IF_IMPORT_FIELDS todavía fuera una variable y estuviera activa:
            # if self.import_controller.IF_IMPORT_FIELDS.get():
            #     self.event_manager.emit('progress_update', "Importando campos estándar...")
            #     time.sleep(1.0)

            self.event_manager.emit('operation_completed', "Importación a Odoo completada.")
        except Exception as e:
            if not self._stop_current_operation.is_set():
                self.event_manager.emit('operation_error', f"Error en importación a Odoo: {str(e)}")
