# controller.py
from core import App_Util
from core import App_Import
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
                self.USER_SELECTIONS_PATH = tk.StringVar(master=master_tk, value=os.path.expanduser(
                    "~/Documents/SMI Files/Data/data/common/json/user_selections.json"))
                self.SKIP_SKUS_JSON_PATH = tk.StringVar(master=master_tk, value=os.path.expanduser(
                    "~/Documents/SMI Files/Data/data/common/json/SKUS_TO_SKIP.json"))
                # Nuevas variables para la comparación entre Excels Merge
                self.MERGED_EXCEL_PATH_OLD = tk.StringVar(master=master_tk)
                self.MERGED_EXCEL_PATH_NEW = tk.StringVar(master=master_tk)

                # Import Odoo (Variables actualizadas según solicitud)
                self.IF_IMPORT_PRODUCTS = tk.BooleanVar(master=master_tk, value=1)
                self.IF_PUBLISH_PRODUCTS = tk.BooleanVar(master=master_tk)
                self.IF_IMPORT_COMERCIAL_INFO = tk.BooleanVar(master=master_tk)
                self.IF_USE_SKU = tk.BooleanVar(master=master_tk)
                self.IF_USE_CATEGORIA = tk.BooleanVar(master=master_tk)

        self.scraper_controller = TkVarHolder(self.master)
        self.excel_controller = TkVarHolder(self.master)
        self.merge_controller = TkVarHolder(self.master)
        self.import_controller = TkVarHolder(self.master)

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
                else:
                    self.event_manager.emit('operation_error', f"Error en la operación: {type(e).__name__} - {str(e)}")
                    import traceback
                    self.event_manager.emit('progress_update', f"Traceback: {traceback.format_exc()}")
            finally:
                self._current_operation_thread = None
                if not self._stop_current_operation.is_set():
                    self.event_manager.emit('operation_completed', "Proceso finalizado.")

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

    # SCRAP
    def is_scraping_active(self):
        return self.scrap_status == ScrapStatus.RUNNING

    def execute_scraper_controller_actions(self):
        print("\n🚀 CONTROLLER: EJECUTANDO SCRAPING")
        if self.is_critical_operation_active():
            self.event_manager.emit('scraping_error', "Ya hay una operación en curso.")
            return
        self._start_long_operation(self._run_scraping_task)

    def abrir_vtac_gui(self, comercial_flag: bool = True, autostart: bool = True):
        import os, sys, subprocess, glob
        base_dir = os.path.expanduser(r"~/Documents/SMI Files")
        # Busca VtacGUI*.exe por si lo renombraste (coge el primero que encuentre)
        matches = sorted(glob.glob(os.path.join(base_dir, "VtacGUI*.exe")))
        if not matches:
            raise FileNotFoundError(f"No se encontró VtacGUI.exe en: {base_dir}")
        exe_path = matches[0]

        env = os.environ.copy()
        env.setdefault("SCRAPY_SETTINGS_MODULE", "Scrap.settings")

        args = [exe_path]
        if autostart: args += ["--autostart"]
        args += ["--comercial", "1" if comercial_flag else "0"]

        # Lanza sin bloquear tu app; cwd asegura rutas relativas correctas
        subprocess.Popen(args, cwd=base_dir, env=env, shell=False)

    def _run_scraping_task(self):
        import time

        print("🔥 DEBUG: SCRAPING - MODO MULTI-REGIÓN")

        try:
            scrap_italia = bool(self.scraper_controller.ITALIA.get())
            scrap_uk = bool(self.scraper_controller.UK.get())
            scrap_espana = bool(self.scraper_controller.SPAIN.get())
            scrap_buyled = bool(self.scraper_controller.BUYLED.get())
            only_new_products = bool(self.scraper_controller.IF_ONLY_NEW_PRODUCTS.get())

        except Exception as e:
            error_msg = f"Error obteniendo configuraciones: {str(e)}"
            self.event_manager.emit('scraping_error', error_msg)
            self._update_status(ScrapStatus.ERROR, error_msg)
            return

        def run_scraping_thread():
            try:
                self._update_status(ScrapStatus.RUNNING, "Iniciando scraping...")

                def progress_callback_internal(message):
                    if self._stop_current_operation.is_set():
                        raise InterruptedError("Scraping cancelado por el usuario.")
                    self.event_manager.emit('progress_update', message)
                    time.sleep(0.01)

                progress_callback_internal("Configuración validada")
                time.sleep(0.2)

                if not any([scrap_italia, scrap_uk, scrap_espana, scrap_buyled]):
                    raise ValueError("Debe seleccionar al menos una región para el scraping.")

                progress_callback_internal("Importando módulos de scraping...")
                time.sleep(0.2)

                vtac_it = vtac_uk = vtac_buyleds = None
                if scrap_italia:
                    try:
                        from Scrap import vtac_it
                    except ImportError as e:
                        progress_callback_internal(f"Advertencia: No se pudo importar vtac_it: {e}")
                if scrap_uk:
                    try:
                        from Scrap import vtac_uk
                    except ImportError as e:
                        progress_callback_internal(f"Advertencia: No se pudo importar vtac_uk: {e}")
                if scrap_buyled:
                    try:
                        from Scrap import vtac_buyleds
                    except ImportError as e:
                        progress_callback_internal(f"Advertencia: No se pudo importar vtac_buyleds: {e}")

                progress_callback_internal("Módulos de scraping listos.")
                time.sleep(0.2)

                if scrap_espana:
                    self.event_manager.emit('progress_update', "Abriendo GUI de scraping (vtac_es)…")
                    self.abrir_vtac_gui(False, True)

                if scrap_italia and vtac_it:
                    progress_callback_internal("🇮🇹 Iniciando scraping de Italia...")
                    vtac_it.run_vtac_italia_scraper(modo_scrap=vtac_it.ScrapMode.COMPLETO,
                                                    only_new_products=only_new_products,
                                                    progress_callback=lambda m: progress_callback_internal(f"🇮🇹 {m}"))
                    progress_callback_internal("🇮🇹 Scraping de Italia completado.")

                if self._stop_current_operation.is_set():
                    raise InterruptedError("Cancelado tras Italia.")

                if scrap_uk and vtac_uk:
                    progress_callback_internal("🇬🇧 Iniciando scraping de Reino Unido...")
                    vtac_uk.run_vtac_uk_scraper(modo_scrap=vtac_uk.ScrapMode.TEST,
                                                only_new_products=only_new_products,
                                                progress_callback=lambda m: progress_callback_internal(f"🇬🇧 {m}"))
                    progress_callback_internal("🇬🇧 Scraping de Reino Unido completado.")

                if self._stop_current_operation.is_set():
                    raise InterruptedError("Cancelado tras UK.")


                if self._stop_current_operation.is_set():
                    raise InterruptedError("Cancelado tras España.")

                if scrap_buyled and vtac_buyleds:
                    progress_callback_internal("💡 Iniciando scraping de BuyLED...")
                    if hasattr(vtac_buyleds, 'ScraperBuyLedStocks') and hasattr(vtac_buyleds.ScraperBuyLedStocks, 'start_scrape'):
                        vtac_buyleds.ScraperBuyLedStocks.start_scrape(
                            progress_callback=lambda m: progress_callback_internal(f"💡 {m}"))
                        progress_callback_internal("💡 Scraping de BuyLED completado.")
                    else:
                        progress_callback_internal("💡 Módulo BuyLED no configurado correctamente.")

                if self._stop_current_operation.is_set():
                    raise InterruptedError("Cancelado tras BuyLED.")

                success_message = "Scraping finalizado exitosamente."
                self._update_status(ScrapStatus.COMPLETED, success_message)
                self.event_manager.emit('scraping_completed', success_message)

            except InterruptedError as ie:
                self.event_manager.emit('progress_update', str(ie))
                self._update_status(ScrapStatus.IDLE, "Scraping cancelado.")
                self.event_manager.emit('scraping_error', "Scraping cancelado por el usuario.")

            except ValueError as ve:
                self._update_status(ScrapStatus.ERROR, f"Error: {str(ve)}")
                self.event_manager.emit('scraping_error', str(ve))
            except Exception as e:
                error_msg = f"Error inesperado: {str(e)}"
                self._update_status(ScrapStatus.ERROR, error_msg)
                self.event_manager.emit('scraping_error', error_msg)
                import traceback
                self.event_manager.emit('progress_update', f"Traceback: {traceback.format_exc()}")

        run_scraping_thread()

    # MERGE
    def execute_excel_controller_actions(self):
        print("📊 CONTROLLER: Ejecutando acciones de Excel...")
        if self.is_critical_operation_active():
            self.event_manager.emit('operation_error', "Ya hay una operación en curso.")
            return

        mega_excel_path_value = self.excel_controller.MEGA_EXCEL_PATH.get()
        generar_nuevo_bloque_value = self.excel_controller.GENERAR_NUEVO_BLOQUE.get()

        self._start_long_operation(self._run_excel_task, mega_excel_path_value, generar_nuevo_bloque_value)

    def _run_excel_task(self, mega_excel_path, generar_nuevo_bloque):
        try:
            self.event_manager.emit('progress_update', "Iniciando proceso de Excel...")
            if not mega_excel_path:
                raise ValueError("Ruta del Mega Excel no especificada.")
            self.event_manager.emit('progress_update', f"Procesando Mega Excel: {mega_excel_path}")

            # Simulación de trabajo
            for i in range(5):
                if self._stop_current_operation.is_set():
                    self.event_manager.emit('progress_update', "Proceso de Excel cancelado.")
                    return
                time.sleep(0.3)
                self.event_manager.emit('progress_update', f"Paso {i + 1} del procesamiento Excel...")

            if generar_nuevo_bloque:
                self.event_manager.emit('progress_update', "Generando nuevo bloque...")
                time.sleep(1)

            self.event_manager.emit('operation_completed', "Proceso de Excel completado.")
        except Exception as e:
            if not self._stop_current_operation.is_set():
                self.event_manager.emit('operation_error', f"Error en proceso de Excel: {str(e)}")

    def _stop_comparison_server(self):
        """Método auxiliar para detener el servidor de comparación si existe"""
        if self.comparison_server_thread and self.comparison_server_thread.is_alive():
            self.comparison_server_stop_event.set()
            try:
                self.comparison_server_thread.join(timeout=5)
                if self.comparison_server_thread.is_alive():
                    print("Advertencia: El servidor de comparación no se detuvo correctamente")
            except Exception as e:
                print(f"Error al detener el servidor de comparación: {e}")
            finally:
                self.comparison_server_thread = None

    # IMPORT
    def execute_import_controller_action(self):
        print("📥 CONTROLLER: Ejecutando acciones de importación...")
        if self.is_critical_operation_active():
            self.event_manager.emit('operation_error', "Ya hay una operación en curso.")
            return

        if_import_products_value = self.import_controller.IF_IMPORT_PRODUCTS.get()
        if_publish_products_value = self.import_controller.IF_PUBLISH_PRODUCTS.get()
        if_import_comercial_info_value = self.import_controller.IF_IMPORT_COMERCIAL_INFO.get()
        if_use_sku_value = self.import_controller.IF_USE_SKU.get()
        if_use_categoria_value = self.import_controller.IF_USE_CATEGORIA.get()

        self._start_long_operation(self._run_import_task,
                                   if_import_products_value,
                                   if_publish_products_value,
                                   if_import_comercial_info_value,
                                   if_use_sku_value,
                                   if_use_categoria_value)

    def _run_import_task(self, if_import_products, if_publish_products, if_import_comercial_info, if_use_sku,
                         if_use_categoria):
        try:
            connection_attempted_or_established = False

            if if_import_products:
                self.event_manager.emit('progress_update', "⚙️ Iniciando importación de productos...")
                # Llamar directamente sin threading adicional
                App_Import.main_import_with_event_manager(self.event_manager)
                connection_attempted_or_established = True

                if self._stop_current_operation.is_set():
                    return
                self.event_manager.emit('progress_update', "✔️ Importación de productos finalizada.")

            if if_publish_products:
                self.event_manager.emit('progress_update', "🚀 Iniciando publicación de productos...")
                time.sleep(0.2)

                if not if_use_sku and not if_use_categoria:
                    self.event_manager.emit('operation_error',
                                            "⚠️ Para publicar, debe seleccionar 'SKU' o 'Categoría'. Publicación omitida.")
                    return
                else:
                    if not connection_attempted_or_established and not App_Import.is_connected:
                        self.event_manager.emit('progress_update',
                                                "🔌 Necesaria conexión a Odoo para publicar. Abriendo ajustes...")
                        App_Import.abrir_ajustes()
                        connection_attempted_or_established = True

                    if self._stop_current_operation.is_set():
                        return

                    if App_Import.is_connected:
                        self.event_manager.emit('progress_update', "🟢 Conexión con Odoo activa.")
                        por_categorias_value = if_use_categoria

                        if por_categorias_value:
                            self.event_manager.emit('progress_update', "⚙️ Publicando por Categoría...")
                        else:
                            self.event_manager.emit('progress_update', "⚙️ Publicando por SKU...")

                        # Llamar directamente sin threading adicional
                        App_Import.main_publish(por_categorias_value, self.event_manager)

                        if not self._stop_current_operation.is_set():
                            self.event_manager.emit('progress_update',
                                                    "✔️ Proceso de publicación finalizado exitosamente.")

                    else:
                        self.event_manager.emit('operation_error',
                                                "❌ No se pudo conectar a Odoo. Publicación de productos omitida.")

            if self._stop_current_operation.is_set():
                return

            if if_import_comercial_info:
                self.event_manager.emit('progress_update', "ℹ️ Iniciando importación de información comercial...")
                time.sleep(0.2)
                if not connection_attempted_or_established and not App_Import.is_connected:
                    self.event_manager.emit('progress_update',
                                            "🔌 Necesaria conexión a Odoo para info. comercial. Abriendo ajustes...")
                    App_Import.abrir_ajustes()
                    connection_attempted_or_established = True

                if self._stop_current_operation.is_set():
                    return

                if App_Import.is_connected:
                    self.event_manager.emit('progress_update', "🟢 Conexión con Odoo activa.")
                    self.event_manager.emit('progress_update', "⚙️ Importando datos comerciales...")

                    App_Util.import_comercial_stock()

                    if self._stop_current_operation.is_set():
                        return
                    self.event_manager.emit('progress_update', "✔️ Información comercial importada.")
                else:
                    self.event_manager.emit('operation_error',
                                            "❌ No se pudo conectar a Odoo. Importación de info. comercial omitida.")

            if self._stop_current_operation.is_set():
                return

            # Solo emitir operation_completed si llegamos hasta aquí sin errores
            self.event_manager.emit('operation_completed', "✅ Tareas de Odoo procesadas.")

        except Exception as e:
            if not self._stop_current_operation.is_set():
                self.event_manager.emit('operation_error', f"❌ Error general en tareas de Odoo: {str(e)}")
            raise

