import importlib
import scrapy
import importlib.util
import os

"""
cd Scrap/Scrap/spiders
scrapy crawl vtac -a region=vtac_es
& "C:/Users/Quimi/AppData/Local/Python/bin/python.exe" -m scrapy crawl vtac -a region=vtac_es
scrapy crawl vtac -a region=advance
"""

class VtacSpider(scrapy.Spider):
    name = "vtac"
    allowed_domains = []
    start_urls = []

    def __init__(self, region=None, comercial_scrap=None, *args, **kwargs):
        super().__init__(*args, **kwargs)

        if region not in ["vtac_es", "vtac_uk", "vtac_it", "buyled", "advance"]:
            raise ValueError(f"Región no válida: {region}")

        # 1) Guardar el flag (soporta "1/0", "true/false", etc.)
        def _to_bool(v):
            if v is None:
                return None
            s = str(v).strip().lower()
            return s in ("1","true","t","yes","y","on","si","sí")
        self.comercial_scrap = _to_bool(comercial_scrap)

        # 2) Cargar handler de forma “EXE-friendly”
        try:
            # Recomendado: importar por paquete (funciona mejor en PyInstaller)
            self.handler = importlib.import_module(f".handlers.{region}", package=__package__ or __name__.rsplit('.',1)[0])
        except Exception:
            # Fallback: por ruta (como tenías antes)
            handler_path = os.path.join(os.path.dirname(__file__), "handlers", f"{region}.py")
            spec = importlib.util.spec_from_file_location("handler_module", handler_path)
            handler_module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(handler_module)
            self.handler = handler_module

        self.handler.init(self)  # carga allowed_domains/start_urls/etc.


    async def start(self):
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/123.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "en-US,en;q=0.9",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
        }

        for url in self.start_urls:
            yield scrapy.Request(url, callback=self.parse, headers=headers, dont_filter=True)

    def parse(self, response):
        return self.handler.parse(self, response)

    def parse_category(self, response):
        return self.handler.parse_category(self, response)

    def parse_product(self, response):
        return self.handler.parse_product(self, response)

    def closed(self, reason):
        self.handler.closed(self, reason)
