#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#PS C:\Users\Don Keo\PycharmProjects\V-Tac 2.0\Scrap>  PyInstaller --onefile --windowed --name VtacGUI --hidden-import=twisted.internet.asyncioreactor --hidden-import=twisted.internet.selectreactor --hidden-import=service_identity --hidden-import=certifi --hidden-import=Scrap.spiders.vtac_spider --hidden-import=Scrap.spiders.handlers.vtac_es --add-data "Scrap\spiders\handlers;Scrap\spiders\handlers" vtac_gui_standalone.py

import os
import logging
import queue
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

# --- Instalar reactor Asyncio ANTES de llamar a crochet.setup() ---
import sys, asyncio

# En Windows, usa la política de selector (compatible con Twisted)
if sys.platform.startswith("win"):
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    except Exception:
        pass

from twisted.internet import asyncioreactor
try:
    # Instala el reactor Asyncio usando el loop actual
    asyncioreactor.install(asyncio.get_event_loop())
except Exception:
    # Si ya estuviera instalado, lo ignoramos
    pass
# ------------------------------------------------------------------


# ---------- Configuración de logging global (Scrapy + Twisted) ----------
# (1) Arrancar Crochet (reactor en hilo aparte) una única vez
from crochet import setup, run_in_reactor
setup()

# (2) Redirigir logs de Scrapy a logging estándar
from scrapy.utils.log import configure_logging
configure_logging({})  # usa formato y nivel por defecto de Scrapy

# (3) Twisted -> logging de Python
from twisted.python.log import PythonLoggingObserver
PythonLoggingObserver().start()

# (4) Nivel por defecto (cámbialo a DEBUG si quieres más detalle)
logging.getLogger().setLevel(logging.INFO)

# ---------- Runner global (reutilizable) ----------
from scrapy.crawler import CrawlerRunner
from scrapy.utils.project import get_project_settings
from scrapy.settings import Settings

_runner = None  # se inicializa perezosamente


def _get_runner() -> CrawlerRunner:
    """Crea (una sola vez) y devuelve un CrawlerRunner con settings válidos."""
    global _runner
    if _runner is not None:
        return _runner

    # Intento 1: usar settings del proyecto (si se ejecuta dentro del proyecto Scrapy)
    try:
        s = get_project_settings()
        # Heurística: si no hay atributos, puede venir vacío (entorno fuera del proyecto).
        if getattr(s, "attributes", None):
            _runner = CrawlerRunner(s)
            logging.getLogger(__name__).info("CrawlerRunner inicializado con get_project_settings().")
            return _runner
    except Exception as e:
        logging.getLogger(__name__).warning("get_project_settings() falló: %s", e)

    # Intento 2: forzar un módulo de settings concreto (ajusta si tu paquete es distinto)
    DEFAULT_SCRAPY_SETTINGS_MODULE = os.getenv("SCRAPY_SETTINGS_MODULE", "Scrap.settings")
    os.environ.setdefault("SCRAPY_SETTINGS_MODULE", DEFAULT_SCRAPY_SETTINGS_MODULE)

    s2 = Settings()
    try:
        module = __import__(DEFAULT_SCRAPY_SETTINGS_MODULE, fromlist=["*"])
        s2.setmodule(module)
        logging.getLogger(__name__).info("CrawlerRunner con SCRAPY_SETTINGS_MODULE=%s", DEFAULT_SCRAPY_SETTINGS_MODULE)
    except Exception as e:
        logging.getLogger(__name__).warning(
            "No se pudo importar %s, usando Settings vacíos. Error: %s",
            DEFAULT_SCRAPY_SETTINGS_MODULE, e
        )
    _runner = CrawlerRunner(s2)
    return _runner


def _import_spider():
    """
    Importa y devuelve la clase VtacSpider. Ajusta esta lista si tu estructura cambia.
    """
    SPIDER_IMPORT_CANDIDATES = [
        "Scrap.spiders.vtac_spider:VtacSpider",
        "Scrap.Scrap.spiders.vtac_spider:VtacSpider",
        "vtac_spider:VtacSpider",  # por si está en la misma carpeta del script
    ]
    last_exc = None
    for cand in SPIDER_IMPORT_CANDIDATES:
        modname, clsname = cand.split(":")
        try:
            mod = __import__(modname, fromlist=[clsname])
            return getattr(mod, clsname)
        except Exception as e:
            last_exc = e
    raise ImportError(f"No pude importar VtacSpider. Revisa tus rutas. Último error: {last_exc}")


# ---------- Redirección de logs a la interfaz ----------
class TkQueueHandler(logging.Handler):
    """Handler que empuja mensajes de logging a una cola (consumida por la GUI)."""
    def __init__(self, q: "queue.Queue[str]"):
        super().__init__()
        self.q = q

    def emit(self, record):
        try:
            msg = self.format(record)
            self.q.put(msg)
        except Exception:
            self.handleError(record)


# ---------- Lógica de lanzamiento del spider (no tocar UI desde aquí) ----------
@run_in_reactor
def _launch_spider(region: str, comercial: bool, done_cb, err_cb):
    """
    Lanza el spider dentro del reactor (Crochet).
    done_cb y err_cb son callables thread-safe que se invocan al terminar.
    """
    logger = logging.getLogger(__name__)
    logger.info("Lanzando VtacSpider: region=%s comercial=%s", region, comercial)

    VtacSpider = _import_spider()
    runner = _get_runner()

    # Importante: pasar los argumentos esperados por tu spider (region y comercial_scrap)
    d = runner.crawl(VtacSpider, region=region, comercial_scrap=comercial)

    def _ok(_):
        logger.info("Spider finalizado correctamente.")
        if done_cb:
            done_cb()

    def _err(failure):
        logger.error("Spider falló: %s", failure.getErrorMessage())
        if err_cb:
            err_cb(str(failure))

    d.addCallbacks(_ok, _err)
    return d


# ---------- Interfaz Tkinter ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("V-TAC Scraper GUI (standalone)")
        self.geometry("860x560")

        # Cola de logs
        self.log_queue: "queue.Queue[str]" = queue.Queue()
        self._install_logging_to_gui()

        # --- Controles superiores ---
        frm_top = ttk.Frame(self, padding=(10, 10, 10, 5))
        frm_top.pack(fill="x")

        ttk.Label(frm_top, text="Región:").grid(row=0, column=0, sticky="w")
        self.cmb_region = ttk.Combobox(frm_top, values=["vtac_es", "advance"], state="readonly")
        self.cmb_region.set("vtac_es")
        self.cmb_region.grid(row=0, column=1, padx=6, sticky="w")

        self.var_comercial = tk.BooleanVar(value=False)
        self.chk_comercial = ttk.Checkbutton(frm_top, text="Modo comercial (nuevos + descatalogados)", variable=self.var_comercial)
        self.chk_comercial.grid(row=0, column=2, padx=10, sticky="w")

        self.btn_run = ttk.Button(frm_top, text="Iniciar scraping", command=self._on_run)
        self.btn_run.grid(row=0, column=3, padx=6)

        self.btn_clear = ttk.Button(frm_top, text="Limpiar logs", command=self._clear_logs)
        self.btn_clear.grid(row=0, column=4, padx=6)

        # --- Área de logs ---
        frm_logs = ttk.Frame(self, padding=(10, 5, 10, 10))
        frm_logs.pack(fill="both", expand=True)

        self.txt = ScrolledText(frm_logs, wrap="word", state="disabled")
        self.txt.pack(fill="both", expand=True)

        # Empieza a refrescar los logs
        self.after(80, self._poll_log_queue)

        # Mensaje inicial
        logging.getLogger(__name__).info("Interfaz lista. Pulsa 'Iniciar scraping'.")

    def _install_logging_to_gui(self):
        handler = TkQueueHandler(self.log_queue)
        fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
        handler.setFormatter(fmt)
        root_logger = logging.getLogger()
        root_logger.addHandler(handler)

    def _poll_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self._append_log(msg)
        except queue.Empty:
            pass
        # seguir chequeando
        self.after(80, self._poll_log_queue)

    def _append_log(self, msg: str):
        self.txt.configure(state="normal")
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def _clear_logs(self):
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.configure(state="disabled")

    # Callbacks seguros (se ejecutan en hilo principal) que le pasamos a Crochet
    def _on_spider_done(self):
        logging.getLogger(__name__).info("✔️ Scraping finalizado.")

    def _on_spider_error(self, err_text: str):
        logging.getLogger(__name__).error("❌ Scraping fallido: %s", err_text)

    def _on_run(self):
        region = self.cmb_region.get().strip()
        comercial = bool(self.var_comercial.get())

        # Deshabilita botón mientras se lanza (opcional)
        self.btn_run.config(state="disabled")
        logging.getLogger(__name__).info("⚙️ Lanzando scraping (region=%s, comercial=%s)…", region, comercial)

        def _reenable():
            # Se invoca al terminar (éxito o error)
            self.btn_run.config(state="normal")

        def done_cb():
            # Estos callbacks corren en un hilo crochet; agenda en hilo Tk
            self.after(0, self._on_spider_done)
            self.after(0, _reenable)

        def err_cb(text):
            self.after(0, self._on_spider_error, text)
            self.after(0, _reenable)

        # Lanza el spider (no bloquea)
        _launch_spider(region, comercial, done_cb, err_cb)


def run_gui():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    run_gui()