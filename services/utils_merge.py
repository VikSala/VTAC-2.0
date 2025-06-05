# services/utils_merge.py
import json
import copy
import threading
import os
from services.utils import Utils


class DataMerger:
    """
    Clase para fusionar datos de productos de múltiples regiones
    según prioridades específicas de campos
    """

    JSON_DUMP_FREQUENCY = 25

    # Rutas por defecto
    MERGED_PRODUCT_INFO_DIR_PATH = 'data/vtac_merged/PRODUCT_INFO'
    MERGED_PRODUCT_MEDIA_DIR_PATH = 'data/vtac_merged/PRODUCT_MEDIA'
    UPLOADED_DATA_DIR_PATH = 'data/vtac_merged/PRODUCT_INFO_UPLOADED'
    UPLOADED_MEDIA_DIR_PATH = 'data/vtac_merged/PRODUCT_MEDIA_UPLOADED'

    # Configuración de scrapers por país
    COUNTRY_SCRAPERS = {
        'es': {'PRODUCTS_INFO_PATH': 'data/es/info', 'PRODUCTS_MEDIA_PATH': 'data/es/media'},
        'uk': {'PRODUCTS_INFO_PATH': 'data/uk/info', 'PRODUCTS_MEDIA_PATH': 'data/uk/media'},
        'ita': {'PRODUCTS_INFO_PATH': 'data/ita/info', 'PRODUCTS_MEDIA_PATH': 'data/ita/media'}
    }

    # Prioridades de campos
    FIELD_PRIORITIES = {
        'default': ('es', 'uk', 'ita'),
        'website_description': ('es', 'uk'),
        'description_purchase': ('es'),
        'accesorios': ('ita'),
        'transit': ('uk'),
        'almacen2_custom': ('uk')
    }

    MEDIA_FIELDS_PRIORITIES = {
        'imgs': ('ita', 'uk', 'es'),
        'videos': ('uk', 'ita', 'es')
    }

    # Campos a eliminar
    FIELDS_TO_DELETE = [
        'Evolución', 'Id eprel', 'Informe emc', 'Informe lvd',
        'Licencia lvd', 'Informe rohs', 'Inmóvil', 'Ordenable en múltiplos de',
        'Piezas bancales', 'Piezas en juego', 'Product information document (eu fiche)',
        'Se puede pedir en múltiplos de', 'Tamaño polo', 'Etiqueta energética ue',
        'Embalaje', 'Piezas por palet', 'Cantidad por palet', 'Cantidad por caja',
        'Unidad de medida', 'SAMSUNG', 'Tipo de enchufe', 'VDE', 'CB Certificate'
    ]

    # Rutas de configuración
    FIELDS_RENAMES_JSON_PATH = 'data/common/json/FIELDS_RENAMES.json'
    VALUES_RENAMES_JSON_PATH = 'data/common/json/VALUES_RENAMES.json'
    SKUS_TO_SKIP = 'data/common/json/SKUS_TO_SKIP.json'

    # Campos que siempre se mantienen de un país
    COUNTRY_FIELDS_ALWAYS_KEEP = []

    # Datos cargados
    merged_data = []
    merged_media = []
    country_data = {'es': [], 'uk': [], 'ita': []}
    country_media = {'es': [], 'uk': [], 'ita': []}

    @classmethod
    def load_data_for_country(cls, country, only_media=False, progress_callback=None):
        """Carga datos de un país específico"""
        if progress_callback:
            progress_callback(f"📦 Cargando datos de {country.upper()}...")

        if country not in cls.COUNTRY_SCRAPERS:
            raise ValueError(f"País {country} no soportado")

        directory_path = cls.COUNTRY_SCRAPERS[country]['PRODUCTS_INFO_PATH']
        if only_media:
            directory_path = cls.COUNTRY_SCRAPERS[country]['PRODUCTS_MEDIA_PATH']

        data = []

        if not os.path.exists(directory_path):
            if progress_callback:
                progress_callback(f"⚠️ Directorio {directory_path} no existe")
            return data

        # Cargar todos los archivos JSON del directorio
        file_list = cls._get_all_json_files(directory_path)

        for i, file_path in enumerate(file_list):
            if progress_callback:
                progress_callback(f"📄 Procesando archivo {i + 1}/{len(file_list)}: {os.path.basename(file_path)}")

            try:
                with open(file_path, "r", encoding='utf-8') as file:
                    file_data = json.load(file)
                    if isinstance(file_data, list):
                        data.extend(file_data)
                    else:
                        data.append(file_data)
            except Exception as e:
                if progress_callback:
                    progress_callback(f"❌ Error cargando {file_path}: {str(e)}")

        # Procesar campos si no es solo media
        if not only_media:
            data = [cls._process_product_fields(p) for p in data if p is not None]
            if progress_callback:
                progress_callback(f"✅ Procesados {len(data)} productos de {country.upper()}")

        return data

    @classmethod
    def _get_all_json_files(cls, directory_path):
        """Obtiene todos los archivos JSON de un directorio"""
        json_files = []
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                if file.endswith('.json'):
                    json_files.append(os.path.join(root, file))
        return json_files

    @classmethod
    def _process_product_fields(cls, product):
        """Procesa campos de un producto (renombrar, eliminar, etc.)"""
        if not product:
            return None

        # Cargar configuraciones de renombrado
        fields_to_rename = cls._load_json_safe(cls.FIELDS_RENAMES_JSON_PATH, {})
        value_renames = cls._load_json_safe(cls.VALUES_RENAMES_JSON_PATH, {})

        # Renombrar campos
        for field, field_renamed in fields_to_rename.items():
            if field in product:
                product[field_renamed] = product[field]
                del product[field]

        # Eliminar campos no deseados
        for field in cls.FIELDS_TO_DELETE:
            if field in product:
                del product[field]

        # Renombrar valores
        for field, renames in value_renames.items():
            if field in product and isinstance(product[field], str):
                for old_value, new_value in renames:
                    product[field] = product[field].replace(old_value, new_value)

        return product

    @classmethod
    def _load_json_safe(cls, file_path, default=None):
        """Carga un archivo JSON de forma segura"""
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                return json.load(file)
        except:
            return default or {}

    @classmethod
    def load_all(cls, if_omit_media, progress_callback=None):
        """Carga todos los datos de todos los países"""
        if progress_callback:
            progress_callback("🚀 Iniciando carga de datos...")

        # Cargar datos de productos
        for country in cls.COUNTRY_SCRAPERS.keys():
            cls.country_data[country] = cls.load_data_for_country(
                country, False, progress_callback
            )

        # Cargar datos de media si no se omite
        if not if_omit_media:
            for country in cls.COUNTRY_SCRAPERS.keys():
                cls.country_media[country] = cls.load_data_for_country(
                    country, True, progress_callback
                )

        if progress_callback:
            progress_callback("✅ Carga de datos completada")

        return cls

    @classmethod
    def get_product_data_from_country_sku(cls, sku, country, only_media=False):
        """Obtiene datos de un producto específico por SKU y país"""
        data = cls.country_data[country]
        if only_media:
            data = cls.country_media[country]

        for product in data:
            if product.get("default_code") == sku:
                return product
        return None

    @classmethod
    def merge_data(cls, if_omit_media, progress_callback=None):
        """Fusiona los datos según las prioridades configuradas"""
        if progress_callback:
            progress_callback("🔗 Iniciando proceso de fusión...")

        # Obtener SKUs únicos
        all_products = cls.country_data['es'] + cls.country_data['uk'] + cls.country_data['ita']
        unique_skus = cls._get_unique_skus(all_products)

        # Cargar SKUs a omitir
        skus_to_skip = cls._load_json_safe(cls.SKUS_TO_SKIP, {'skus': []})

        if progress_callback:
            progress_callback(f"📊 Procesando {len(unique_skus)} productos únicos...")

        cls.merged_data = []
        cls.merged_media = []

        for i, sku in enumerate(unique_skus):
            if sku in skus_to_skip.get('skus', []):
                if progress_callback:
                    progress_callback(f"⏭️ Omitiendo {sku}")
                continue

            # Progreso cada 10 productos
            if i % 10 == 0 and progress_callback:
                progress_callback(f"🔄 Procesando producto {i + 1}/{len(unique_skus)}: {sku}")

            merged_product, merged_media = cls._merge_single_product(sku, if_omit_media)

            if merged_product:
                cls.merged_data.append(merged_product)

            if merged_media and not if_omit_media:
                cls.merged_media.append(merged_media)

        if progress_callback:
            progress_callback(f"✅ Fusión completada: {len(cls.merged_data)} productos procesados")

        return cls.merged_data, cls.merged_media

    @classmethod
    def _get_unique_skus(cls, products):
        """Obtiene SKUs únicos de una lista de productos"""
        skus = set()
        for product in products:
            if product and product.get('default_code'):
                skus.add(product['default_code'])
        return list(skus)

    @classmethod
    def _merge_single_product(cls, sku, if_omit_media):
        """Fusiona un solo producto según las prioridades"""
        # Obtener datos del producto de todos los países
        product_data = {
            'es': cls.get_product_data_from_country_sku(sku, 'es'),
            'uk': cls.get_product_data_from_country_sku(sku, 'uk'),
            'ita': cls.get_product_data_from_country_sku(sku, 'ita')
        }

        product_media = None
        if not if_omit_media:
            product_media = {
                'es': cls.get_product_data_from_country_sku(sku, 'es', True),
                'uk': cls.get_product_data_from_country_sku(sku, 'uk', True),
                'ita': cls.get_product_data_from_country_sku(sku, 'ita', True)
            }

        # Inicializar producto fusionado
        merged_product = {}
        merged_product_media = {"default_code": sku}

        # Copiar datos del primer país disponible según prioridad por defecto
        for country in cls.FIELD_PRIORITIES['default']:
            if product_data[country] is not None:
                merged_product = copy.deepcopy(product_data[country])
                break

        # Fusionar campos específicos según prioridades
        for field, priorities in cls.FIELD_PRIORITIES.items():
            if field == 'default':
                continue

            for country in priorities:
                country_product = product_data.get(country)
                if (country_product and
                        country_product.get(field) and
                        country_product[field]):

                    if isinstance(country_product[field], list):
                        merged_product[field] = copy.deepcopy(country_product[field])
                    else:
                        merged_product[field] = country_product[field]
                    break

        # Fusionar campos de media
        if not if_omit_media and product_media:
            for field, priorities in cls.MEDIA_FIELDS_PRIORITIES.items():
                for country in priorities:
                    country_media = product_media.get(country)
                    if (country_media and
                            country_media.get(field) and
                            country_media[field]):

                        if isinstance(country_media[field], list):
                            merged_product_media[field] = copy.deepcopy(country_media[field])
                        else:
                            merged_product_media[field] = country_media[field]
                        break

        # Aplicar valores por defecto de V-TAC
        cls._apply_default_vtac_values(merged_product)

        return merged_product, merged_product_media if not if_omit_media else None

    @classmethod
    def _apply_default_vtac_values(cls, product):
        """Aplica valores por defecto específicos de V-TAC"""
        defaults = {
            'transit_stock_custom': 0,
            'almacen1_transito_custom': 0,
            'almacen2_transito_custom': 0,
            'almacen1_custom': 0,
            'almacen2_custom': 0,
            'almacen3_custom': 0,
            'transit': 0,
            'invoice_policy': 'delivery',
            'detailed_type': 'product',
            'show_availability': True,
            'allow_out_of_stock_order': True,
            'available_threshold': 100000
        }

        for field, value in defaults.items():
            product[field] = value

        # Procesar volumen
        if 'volume' in product and product['volume']:
            if isinstance(product['volume'], str):
                try:
                    product['volume'] = float(product['volume'].replace(',', '.'))
                except:
                    pass

    @classmethod
    def extract_merged_data(cls, data, media, progress_callback=None):
        """Extrae los datos fusionados a archivos JSON"""
        if progress_callback:
            progress_callback("💾 Iniciando extracción de datos...")

        # Crear directorios si no existen
        os.makedirs(cls.MERGED_PRODUCT_INFO_DIR_PATH, exist_ok=True)
        if media:
            os.makedirs(cls.MERGED_PRODUCT_MEDIA_DIR_PATH, exist_ok=True)

        data_path_template = f'{cls.MERGED_PRODUCT_INFO_DIR_PATH}/MERGED_INFO_{{}}.json'
        media_path_template = f'{cls.MERGED_PRODUCT_MEDIA_DIR_PATH}/MERGED_MEDIA_{{}}.json'

        def extract_data_async(data_to_extract, path_template, data_type_name):
            """Función para extraer datos de forma asíncrona"""
            if not data_to_extract:
                return

            total_files = (len(data_to_extract) + cls.JSON_DUMP_FREQUENCY - 1) // cls.JSON_DUMP_FREQUENCY

            for index in range(0, len(data_to_extract), cls.JSON_DUMP_FREQUENCY):
                file_number = (index // cls.JSON_DUMP_FREQUENCY) + 1
                end_index = min(index + cls.JSON_DUMP_FREQUENCY, len(data_to_extract))

                file_path = path_template.format(f"{file_number:03d}")

                try:
                    with open(file_path, 'w', encoding='utf-8') as file:
                        json.dump(data_to_extract[index:end_index], file,
                                  ensure_ascii=False, indent=2)

                    if progress_callback:
                        progress_callback(f"💾 {data_type_name}: Archivo {file_number}/{total_files} guardado")

                except Exception as e:
                    if progress_callback:
                        progress_callback(f"❌ Error guardando {data_type_name} archivo {file_number}: {str(e)}")

        # Crear threads para extracción paralela
        threads = []

        if data:
            data_thread = threading.Thread(
                target=extract_data_async,
                args=(data, data_path_template, "DATOS"),
                name='data_extraction'
            )
            threads.append(data_thread)

        if media:
            media_thread = threading.Thread(
                target=extract_data_async,
                args=(media, media_path_template, "MEDIA"),
                name='media_extraction'
            )
            threads.append(media_thread)

        # Iniciar threads
        for thread in threads:
            thread.start()

        # Esperar a que terminen
        for thread in threads:
            thread.join()

        if progress_callback:
            progress_callback("✅ Extracción de datos completada")

    @classmethod
    def get_merge_statistics(cls):
        """Obtiene estadísticas del proceso de merge"""
        stats = {
            'countries_loaded': {},
            'total_products': len(cls.merged_data),
            'total_media': len(cls.merged_media)
        }

        # Estadísticas por país
        for country in cls.COUNTRY_SCRAPERS.keys():
            stats['countries_loaded'][country] = {
                'products': len(cls.country_data.get(country, [])),
                'media': len(cls.country_media.get(country, []))
            }

        return stats

    @classmethod
    def reset_data(cls):
        """Reinicia todos los datos cargados"""
        cls.merged_data = []
        cls.merged_media = []
        cls.country_data = {'es': [], 'uk': [], 'ita': []}
        cls.country_media = {'es': [], 'uk': [], 'ita': []}

    @classmethod
    def configure_paths(cls, info_path, media_path, skus_path, country_paths):
        """Configura las rutas de trabajo"""
        cls.MERGED_PRODUCT_INFO_DIR_PATH = info_path
        cls.MERGED_PRODUCT_MEDIA_DIR_PATH = media_path
        cls.SKUS_TO_SKIP = skus_path

        # Configurar rutas por país
        for country, paths in country_paths.items():
            if country in cls.COUNTRY_SCRAPERS:
                cls.COUNTRY_SCRAPERS[country]['PRODUCTS_INFO_PATH'] = paths.get('info', '')
                cls.COUNTRY_SCRAPERS[country]['PRODUCTS_MEDIA_PATH'] = paths.get('media', '')

