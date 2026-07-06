import math
from collections import namedtuple
from collections import defaultdict
from services.campos_odoo import ClavesExcel, excel_a_odoo, FISCAL_POSITION_MAP
from services.utils import Utils
import xmlrpc.client
import pandas as pd
import numpy as np
import ast, re

from services.utils_download_files import ruta_excel

is_connected = False
atributos_cache = {}
valores_cache = {}

#region Importar Productos
def assign_variant_skus(product_template_id, referencias, attribute_variable, params):
    """
    Asigna los SKUs a las variantes generadas por Odoo, respetando el orden.
    """
    models, db, uid, password = params

    # 1) Leer variantes creadas por Odoo
    variants = models.execute_kw(
        db, uid, password,
        'product.product', 'search_read',
        [[('product_tmpl_id', '=', product_template_id)]],
        {'fields': ['id', 'product_template_attribute_value_ids']}
    )

    # 2) Ordenar las variantes por el valor del atributo variable
    # Esto es CLAVE para que coincidan con referencias[0], referencias[1], etc.
    variants_sorted = sorted(
        variants,
        key=lambda v: v['product_template_attribute_value_ids']
    )

    # 3) Asignar SKU por orden
    for variant, ref in zip(variants_sorted, referencias):
        models.execute_kw(
            db, uid, password,
            'product.product', 'write',
            [[variant['id']], {'default_code': ref}]
        )
        print(f"🔧 Asignado SKU {ref} → variante {variant['id']}")

'''def attributes_to_odoo(product_template_id, specifications, params):
    if isinstance(specifications, str):
        try:
            specifications = ast.literal_eval(specifications)
        except Exception as e:
            print(f"❌ Error convirtiendo especificaciones: {e}")
            specifications = {}

    # Agregar atributos
    try:
        if isinstance(specifications, dict):
            for key, value in specifications.items():
                #print(f"🛠️ Key: {key} | Valor: {value}")
                if value:
                    attr_id = Utils.get_or_create_attribute(key, atributos_cache, params)
                    value_id = Utils.get_or_create_attribute_value(attr_id, value, valores_cache, params)
                    Utils.create_attribute_line(product_template_id, attr_id, value_id, params)
    except Exception as e:
        print(f"❌ Error convirtiendo especificaciones: {e}")'''
def attributes_to_odoo(product_template_id, specifications, params, row=None):
    # Convertir string a dict/list si viene del Excel
    if isinstance(specifications, str):
        try:
            specifications = ast.literal_eval(specifications)
        except Exception:
            specifications = {}

    # ---------------------------
    # CASO 1 → sin variantes
    # ---------------------------
    if isinstance(specifications, dict):

        for key, value in specifications.items():
            if not value:
                continue

            attr_id = Utils.get_or_create_attribute(key, atributos_cache, params)
            value_id = Utils.get_or_create_attribute_value(attr_id, value, valores_cache, params)

            Utils.create_attribute_line(product_template_id, attr_id, value_id, params)

        return  # FIN sin variantes

    # ---------------------------
    # CASO 2 → con variantes
    # ---------------------------
    if isinstance(specifications, list):

        if len(specifications) == 0:
            return

        comunes = specifications[0]

        # 1) Crear atributos comunes
        for key, value in comunes.items():
            attr_id = Utils.get_or_create_attribute(key, atributos_cache, params)
            value_id = Utils.get_or_create_attribute_value(attr_id, value, valores_cache, params)

            Utils.create_attribute_line(product_template_id, attr_id, value_id, params)

        # 2) Detectar atributo variable
        atributo_variable = list(specifications[1].keys())[0]

        attr_var_id = Utils.get_or_create_attribute(atributo_variable, atributos_cache, params)

        # 3) Crear valores del atributo variable
        valores_variantes_ids = []
        for variante_dict in specifications[1:]:
            valor = list(variante_dict.values())[0]
            value_id = Utils.get_or_create_attribute_value(attr_var_id, valor, valores_cache, params)
            valores_variantes_ids.append(value_id)

        # 4) Crear línea que generará las variantes
        params.models.execute_kw(
            params.db, params.uid, params.password,
            'product.template.attribute.line', 'create',
            [{
                'product_tmpl_id': product_template_id,
                'attribute_id': attr_var_id,
                'value_ids': [(6, 0, valores_variantes_ids)]
            }]
        )

        # 5) Asignar SKUs a variantes generadas por Odoo
        referencias = row["SKU"] if "SKU" in row else None

        if referencias:
            try:
                referencias = ast.literal_eval(referencias)
                assign_variant_skus(product_template_id, referencias, atributo_variable, params)
            except Exception as e:
                print(f"⚠️ No se pudieron asignar SKUs de variantes: {e}")


# === IMPORT: Leer Excel y crear Productos en Odoo ===
def import_to_odoo(excel_path, limite = 0, import_all = False, event_manager=None):
    from App_Connection import models, db, uid, password

    NOVEDADES = True
    LEDXPRESS = False
    ADVANCE = False
    ODOO16 = False #True

    ConnectionParams = namedtuple("ConnectionParams", ["models", "db", "uid", "password"])
    params = ConnectionParams(models, db, uid, password)

    def emit_progress(message):
        if event_manager:
            event_manager.emit('progress_update', message)
            print(message)
        else:
            print(message)

    if import_all:
        df = pd.read_excel(excel_path, sheet_name="Products")
    else:
        df = pd.read_excel(excel_path, sheet_name="Import")

    total = len(df)
    emit_progress(f"📦 Encontrados {total} productos para importar")

    # === Iterar filas y crear productos ===
    for index, row in df.iloc[limite:].iterrows():
        try:
            current_product = index + 1
            product_name = row.get(ClavesExcel.NOMBRE.value, f"Producto {current_product}")

            emit_progress(f"📦 Importando producto {current_product}/{total}: {product_name}")

            name = row[ClavesExcel.NOMBRE.value]
            product_data = Utils.build_product_data_from_row(row, Utils, df)

            print("1. Categorías")# Categorías
            #if not ODOO16: Utils.preparar_categorias_para_producto(row, product_data, params)

            print("2. Marca")#Marca
            Utils.preparar_marca_para_producto(row, product_data, params)

            #Iconos Novedades
            if NOVEDADES:
                print("3. Iconos")
                atributos = ast.literal_eval(row["Atributos"])
                product_data['x_icono1'] = Utils.image_url_to_base64("http://158.179.220.107:8069/web/image/1011-1dd7909f/V-TAC.png")
                icons_detected = Utils.ico_match(product_data['name'], False)
                if isinstance(icons_detected, str):
                    icons_detected = icons_detected.split(',')
                    ico_count = 2
                    for icon in icons_detected:
                        if icon != '':
                            product_data['x_icono' + str(ico_count)] = Utils.image_url_to_base64(icon)
                            ico_count += 1
                    for attr_name, val_name in atributos.items():
                        if not attr_name or not val_name:
                            continue

                        ico_str = Utils.ico_match(attr_name, val_name)  # Devuelve ico_url o None
                        if ico_str and ico_count < 9:
                            product_data['x_icono' + str(ico_count)] = Utils.image_url_to_base64(ico_str)
                            ico_count += 1
            elif LEDXPRESS:
                product_data['x_icono1'] = Utils.image_url_to_base64("http://79.72.55.217:8069/web/image/1011-1dd7909f/V-TAC.png")
                led_icons = ast.literal_eval(row["Iconos"])
                ico_count = 2
                for icon in led_icons:
                    product_data['x_icono' + str(ico_count)] = Utils.image_url_to_base64(icon)
                    ico_count += 1
            elif ADVANCE:
                product_data['x_icono1'] = Utils.image_url_to_base64(row["Logo"])

            if LEDXPRESS: product_data['image_1920'] = Utils.image_url_to_base64(row["Imagen principal"])

            print("3,5. Crear producto base...")
            product_template_id = models.execute_kw(
                db, uid, password,
                'product.template', 'create',
                [product_data]
            )

            # 3.6 Crear la regla de precio (pricelist item) con el descuento
            if LEDXPRESS:
                descuento = row["Descuento"]
                if descuento and not (isinstance(descuento, float) and math.isnan(descuento)):
                    try:
                        print("3,6. Crear descuento...")
                        # ID de la tarifa a la que quieres aplicar el descuento
                        # Puedes usar 1 si es la 'Public Pricelist' (tarifa pública)
                        pricelist_id = 1

                        # Leer el valor del descuento desde tu Excel
                        discount_value = float(descuento)

                        # Crear la regla de descuento
                        pricelist_item_id = models.execute_kw(
                            db, uid, password,
                            'product.pricelist.item', 'create',
                            [{
                                'pricelist_id': pricelist_id,
                                'applied_on': '1_product',  # Aplicar a un producto concreto
                                'product_tmpl_id': product_template_id,  # ID del producto recién creado
                                'compute_price': 'percentage',  # Tipo de cálculo: porcentaje
                                'percent_price': discount_value,  # Ej: -20 para un -20%
                            }]
                        )

                        print(f"✅ Pricelist creada (ID: {pricelist_item_id}) con descuento {discount_value}%")

                    except Exception as e:
                        print(f"⚠️ Error creando pricelist para producto {product_template_id}: {e}")

            print("4. Atributos")# Atributos
            specifications = row[ClavesExcel.ATRIBUTOS.value]
            attributes_to_odoo(product_template_id, specifications, params, row)

            print("5. Videos")# Agregar imágenes adicionales si existen
            galeria = ClavesExcel.GALERIA.value
            video = ClavesExcel.VIDEOS.value
            is_video = False
            if video in row and pd.notna(row[video]):
                url_video = ast.literal_eval(row[video])
                if url_video:
                    url_video = ast.literal_eval(row[video])[0]
                    url_img = "https://i.ytimg.com/vi/" + url_video.split("=")[1] + "/sddefault.jpg"
                    is_video = True
                    image_id = models.execute_kw(
                        db, uid, password,
                        'product.image', 'create',
                        [{
                            'name': f"{name}-yt",
                            'product_tmpl_id': product_template_id,
                            'image_1920': Utils.image_url_to_base64(url_img),
                            'video_url': url_video,
                            'sequence': 0
                        }]
                    )

            print("5. Imágenes")
            if galeria in row and pd.notna(row[galeria]):
                try:
                    extra_urls = ast.literal_eval(row[galeria])
                    for i, url in enumerate(extra_urls, start=1):

                        #if "http" not in url: url = "http://143.47.53.74:8070/" + url
                        campo_odoo = 'image_1920' if "youtu" not in url else 'video_url'
                        valor_odoo = Utils.image_url_to_base64(url.strip()) if "youtu" not in url else url
                        ext = ".webp" if "youtu" not in url else ""

                        if valor_odoo:
                            ctx = {
                                "active_test": False,
                                "lang": "es_ES",
                                "interactive": False,
                            }
                            image_id = models.execute_kw(
                                db, uid, password,
                                'product.image', 'create',
                                [{
                                    'name': f"{name}-extra-{i}{ext}",
                                    'product_tmpl_id': product_template_id,
                                    campo_odoo: valor_odoo,
                                    'sequence': i+1 if is_video else i
                                }],
                                {"context": ctx}
                            )
                except Exception as e:
                    msg = f"⚠️ Error procesando imágenes adicionales para {product_name}: {e}"
                    emit_progress(msg)
                    print(msg)

            emit_progress(f"✅ Producto {current_product}/{total} importado: {product_name}")

        except Exception as e:
            msg = f"❌ Error importando producto {current_product}/{total}: {e}"
            emit_progress(msg)
            print(msg)
            continue

    if import_all and not NOVEDADES and not LEDXPRESS and not ADVANCE and not ODOO16:
        '''def extraer_iconos_por_sku(excel_file_path, batch_size=100):
            # 1. Conexión a Odoo
            url_b2b = 'https://optimaluz.soluntec.net'
            db_b2b = 'Real'
            username_b2b = 'Usuario programación'
            password_b2b = 'Programación'

            common_b2b = xmlrpc.client.ServerProxy(f'{url_b2b}/xmlrpc/2/common', allow_none=True)
            uid_b2b = common_b2b.authenticate(db_b2b, username_b2b, password_b2b, {})

            if not uid:
                print("❌ No se pudo autenticar.")
                return

            print(f'🔌 Conectado como {username_b2b} (uid: {uid})')
            models_b2b = xmlrpc.client.ServerProxy(f'{url_b2b}/xmlrpc/2/object', allow_none=True)

            try:
                # Leer Excel y extraer SKUs únicos
                df = pd.read_excel(excel_file_path)
                if 'SKU' not in df.columns:
                    raise ValueError("❌ El Excel debe contener una columna llamada 'SKU'.")

                skus = df['SKU'].dropna().astype(str).unique().tolist()
                resultado = {}

                # Procesar en batches
                for i in range(0, len(skus), batch_size):
                    batch_skus = skus[i:i + batch_size]
                    print(f"🔄 Procesando batch {i + 1}-{i + len(batch_skus)} / {len(skus)}")

                    # Buscar productos en Odoo por default_code
                    domain = [('default_code', 'in', batch_skus)]
                    fields = ['default_code'] + [f'x_icono{i}' for i in range(1, 9)]

                    products = models_b2b.execute_kw(
                        db_b2b, uid_b2b, password_b2b,
                        'product.template', 'search_read',
                        [domain],
                        {'fields': fields}
                    )

                    # Mapear iconos por SKU
                    for product in products:
                        sku = product['default_code']
                        iconos = [product.get(f'x_icono{i}') for i in range(1, 9) if product.get(f'x_icono{i}')]
                        resultado[sku] = iconos

                    # Añadir SKUs no encontrados (vacíos)
                    encontrados = {p['default_code'] for p in products}
                    no_encontrados = set(batch_skus) - encontrados
                    for sku in no_encontrados:
                        resultado[sku] = []
                        print(f"⚠️ SKU no encontrado en Odoo: {sku}")

                print(f"\n✅ Iconos extraídos correctamente para {len(resultado)} productos.")
                return resultado

            except Exception as e:
                print(f"❌ Error extrayendo iconos: {e}")
                return {}

        iconos_por_sku_ = extraer_iconos_por_sku(excel_path)

        def actualizar_iconos_por_sku(excel_file_path, iconos_por_sku, batch_size=100):
            # 2. Leer Excel con columna SKU
            df = pd.read_excel(excel_file_path)
            if 'SKU' not in df.columns:
                print("❌ El Excel debe contener una columna llamada 'SKU'.")
                return

            skus = df['SKU'].dropna().astype(str).unique().tolist()
            total = len(skus)
            actualizados = 0
            no_encontrados = 0

            # 3. Procesar por lotes
            for i in range(0, total, batch_size):
                batch_skus = skus[i:i + batch_size]

                # Buscar productos por default_code en lote
                productos = models.execute_kw(
                    db, uid, password,
                    'product.template', 'search_read',
                    [[('default_code', 'in', batch_skus)]],
                    {'fields': ['id', 'default_code']}
                )

                sku_to_id = {p['default_code']: p['id'] for p in productos}

                for sku in batch_skus:
                    template_id = sku_to_id.get(sku)
                    iconos = iconos_por_sku.get(sku, [])

                    if not template_id:
                        print(f"❌ Producto con SKU '{sku}' no encontrado en Odoo.")
                        no_encontrados += 1
                        continue

                    if not iconos:
                        print(f"⚠️ Sin iconos para SKU: {sku}")
                        continue

                    # Construir campos para write()
                    valores_iconos = {}
                    for j, icono_b64 in enumerate(iconos):
                        if j >= 8:
                            break  # Solo hasta x_icono8
                        campo = f'x_icono{j + 1}'
                        valores_iconos[campo] = icono_b64

                    try:
                        models.execute_kw(
                            db, uid, password,
                            'product.template', 'write',
                            [[template_id], valores_iconos]
                        )
                        actualizados += 1
                        print(f"✅ Iconos actualizados para SKU: {sku}")
                    except Exception as e:
                        print(f"❌ Error al actualizar SKU '{sku}': {e}")

            print(f"\n🔁 Proceso completado: {actualizados} productos actualizados, {no_encontrados} no encontrados.")

        actualizar_iconos_por_sku(excel_path, iconos_por_sku_)'''

        def extraer_media_por_sku(excel_path, batch_size=100):
            import pandas as pd
            import xmlrpc.client

            url = 'https://optimaluz.soluntec.net'
            db_o = 'Real'
            user = 'Usuario programación'
            pwd = 'Programación'

            common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common', allow_none=True)
            uid_o = common.authenticate(db_o, user, pwd, {})
            models_o = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object', allow_none=True)

            df = pd.read_excel(excel_path)
            skus = df['SKU'].dropna().astype(str).unique().tolist()

            resultado = {}

            for i in range(0, len(skus), batch_size):
                batch = skus[i:i + batch_size]

                templates = models_o.execute_kw(
                    db_o, uid_o, pwd,
                    'product.template', 'search_read',
                    [[('default_code', 'in', batch)]],
                    {'fields': ['id', 'default_code', 'image_1920'] + [f'x_icono{i}' for i in range(1, 9)]}
                )

                tmpl_by_id = {t['id']: t for t in templates}
                sku_to_tmpl = {t['default_code']: t['id'] for t in templates}

                images = models_o.execute_kw(
                    db_o, uid_o, pwd,
                    'product.image', 'search_read',
                    [[('product_tmpl_id', 'in', list(tmpl_by_id))]],
                    {'fields': ['product_tmpl_id', 'image_1920', 'sequence']}
                )

                images_by_tmpl = {}
                for img in images:
                    images_by_tmpl.setdefault(img['product_tmpl_id'][0], []).append(img)

                for sku, tmpl_id in sku_to_tmpl.items():
                    t = tmpl_by_id[tmpl_id]
                    resultado[sku] = {
                        "iconos": [t.get(f'x_icono{i}') for i in range(1, 9) if t.get(f'x_icono{i}')],
                        "main": t.get('image_1920'),
                        "gallery": images_by_tmpl.get(tmpl_id, [])
                    }

            return resultado

        def aplicar_media_por_sku(excel_path, media_por_sku, batch_size=100):
            import pandas as pd

            df = pd.read_excel(excel_path)
            skus = df['SKU'].dropna().astype(str).unique().tolist()

            for i in range(0, len(skus), batch_size):
                batch = skus[i:i + batch_size]

                templates = models.execute_kw(
                    db, uid, password,
                    'product.template', 'search_read',
                    [[('default_code', 'in', batch)]],
                    {'fields': ['id', 'default_code']}
                )

                sku_to_id = {t['default_code']: t['id'] for t in templates}

                for sku in batch:
                    media = media_por_sku.get(sku)
                    tmpl_id = sku_to_id.get(sku)

                    if not media or not tmpl_id:
                        continue

                    valores = {}

                    # Iconos
                    for i, icono in enumerate(media['iconos'][:8]):
                        valores[f'x_icono{i + 1}'] = icono

                    # Imagen principal
                    if media.get("main"):
                        valores['image_1920'] = media['main']

                    if valores:
                        models.execute_kw(
                            db, uid, password,
                            'product.template', 'write',
                            [[tmpl_id], valores]
                        )

                    # Galería
                    for img in media.get("gallery", []):
                        if not img.get("image_1920"):
                            continue
                        models.execute_kw(
                            db, uid, password,
                            'product.image', 'create',
                            [{
                                'product_tmpl_id': tmpl_id,
                                'image_1920': img['image_1920'],
                                'sequence': img.get('sequence', 10)
                            }]
                        )

        media_por_sku_ = extraer_media_por_sku(excel_path)
        aplicar_media_por_sku(excel_path, media_por_sku_)

    emit_progress("🎉 Importación de productos completada.")

# endregion

# region Actualizar Productos

def update_from_merge_to_odoo(excel_path, event_manager):
    """
    Actualiza todos los productos a partir de los datos en la hoja 'Update' del Excel.
    """
    from App_Connection import models, db, uid, password
    ConnectionParams = namedtuple("ConnectionParams", ["models", "db", "uid", "password"])
    params = ConnectionParams(models, db, uid, password)

    def emit_progress(message):
        if event_manager:
            event_manager.emit('progress_update', message)
            print(message)
        else:
            print(message)

    def update_odoo_product(product_sku, update_values):
        """
        Actualiza un product.template en Odoo por sku.
        """
        # Buscar producto por sku
        product_ids = models.execute_kw(
            db, uid, password,
            'product.template', 'search',
            [[['default_code', '=', product_sku]]]
        )

        if product_ids:
            # Actualizar el primer producto encontrado
            success = models.execute_kw(
                db, uid, password,
                'product.template', 'write',
                [[product_ids[0]], update_values]
            )
            return success
        else:
            print(f"Producto '{product_sku}' no encontrado.")
            return False

    df = pd.read_excel(excel_path, sheet_name="Update")
    if df.empty or df.dropna(how='all').empty:
        emit_progress("ℹ️ No hay productos para actualizar")
        return
    # Limpiar filas vacías o sin datos válidos
    df = df.dropna(subset=["SKU", "Actualizar"])
    total = len(df)
    emit_progress(f"🔄 Encontrados {total} productos para actualizar")

    for index, row in df.iterrows():
        sku = row["SKU"]
        current_product = index + 1
        campos_a_actualizar = row["Actualizar"]
        emit_progress(f"🔄 Actualizando producto {current_product}/{total}: {sku}")

        # Si es string, evalúa el diccionario (en caso de que Excel lo haya guardado como texto)
        if isinstance(campos_a_actualizar, str):
            try:
                campos_a_actualizar = eval(campos_a_actualizar)
            except Exception as e:
                print(f"⚠️ Error al interpretar los campos de '{sku}': {e}")
                continue

        # Si se incluye ATRIBUTOS, tratarlo por separado
        atributos = ClavesExcel.ATRIBUTOS.value
        if atributos in campos_a_actualizar:
            product_id = models.execute_kw(
                db, uid, password,
                'product.template', 'search',
                [[['default_code', '=', sku]]]
            )
            if product_id:
                attributes_to_odoo(product_id[0], campos_a_actualizar[atributos], params)
                campos_a_actualizar.pop(atributos)  # Evitar actualizarlo también por write()

        # Si hay más campos, hacer una única actualización
        campos_odoo = {}

        for clave_excel, valor in campos_a_actualizar.items():
            clave_excel_norm = clave_excel.strip()
            if clave_excel_norm in excel_a_odoo:
                clave_odoo = excel_a_odoo[clave_excel_norm]
                campos_odoo[clave_odoo] = valor
            else:
                print(f"⚠️ Clave no encontrada en mapeo: '{clave_excel_norm}'")

        if campos_odoo:
            actualizado = update_odoo_product(
                product_sku=sku,
                update_values=campos_odoo
            )

            if actualizado:
                print(f"✅ Producto '{sku}' actualizado correctamente.")
            else:
                print(f"❌ Fallo al actualizar '{sku}'.")

    emit_progress("🎉 Actualización de productos completada.")

def delete_from_merge_to_odoo(excel_path, event_manager=None):
    from App_Connection import models, db, uid, password

    def clean_value(val):
        if isinstance(val, (np.generic, pd._libs.tslibs.nattype.NaTType)):
            return val.item()
        elif pd.isna(val):
            return None
        return val

    def emit_progress(message):
        if event_manager:
            event_manager.emit('progress_update', message)
            print(message)
        else:
            print(message)

    try:
        df = pd.read_excel(excel_path, sheet_name='Quantity')
    except Exception as e:
        emit_progress(f"⚠️ No se encontró hoja 'Quantity' o error leyendo: {e}")
        return

    if df.empty or df.dropna(how='all').empty:
        emit_progress("ℹ️ No hay productos para eliminar")
        return

    productos_a_eliminar = df['Producto Eliminado'].dropna().unique()
    total = len(productos_a_eliminar)

    if total == 0:
        emit_progress("ℹ️ No hay productos marcados para eliminación")
        return

    emit_progress(f"🗑️ Encontrados {total} productos para eliminar")

    for index, sku in enumerate(productos_a_eliminar):
        try:
            sku = clean_value(sku)
            current_product = index + 1
            emit_progress(f"🗑️ Eliminando producto {current_product}/{total}: {sku}")

            # Buscar producto por nombre
            product_ids = models.execute_kw(
                db, uid, password,
                'product.template', 'search',
                [[('default_code', '=', sku)]]
            )

            if product_ids:
                # Desactivar en lugar de eliminar (más seguro)
                models.execute_kw(
                    db, uid, password,
                    'product.template', 'write',
                    [[product_ids[0]], {'active': False}]
                )
                emit_progress(f"✅ Producto {current_product}/{total} desactivado: {sku}")
            else:
                emit_progress(f"⚠️ Producto no encontrado {current_product}/{total}: {sku}")

        except Exception as e:
            emit_progress(f"❌ Error eliminando producto {current_product}/{total}: {e}")
            continue

    emit_progress("🎉 Eliminación de productos completada.")

#Publicar por Categoria/SKUs
def main_publish(por_categorias: bool, event_manager=None):
    from App_Connection import models, db, uid, password

    def emit_progress(message):
        if event_manager:
            event_manager.emit('progress_update', message)
            print(message)
        else:
            print(message)

    try:
        emit_progress("📁 Seleccionando archivo Excel...")
        ruta_excel = Utils.seleccionar_excel()

        if not ruta_excel:
            emit_progress("❌ Publicación cancelada - No se seleccionó archivo.")
            if event_manager:
                event_manager.emit('operation_error', "No se seleccionó archivo para publicar")
            return

        emit_progress("📊 Leyendo datos del archivo Excel...")
        df = pd.read_excel(ruta_excel)

        if por_categorias:
            if "Familia" not in df.columns:
                emit_progress("❌ No se encontró la columna 'Familia' en el Excel.")
                if event_manager:
                    event_manager.emit('operation_error', "Columna 'Familia' no encontrada")
                return

            categorias = df["Familia"].dropna().unique().tolist()
            total_categorias = len(categorias)
            emit_progress(f"🏷️ Publicando productos por categoría ({total_categorias} categorías)...")

            productos_publicados = 0
            for i, categoria in enumerate(categorias, 1):
                emit_progress(f"🏷️ Procesando categoría {i}/{total_categorias}: {categoria}")

                # Buscar ID de categoría por nombre
                categoria_ids = models.execute_kw(
                    db, uid, password,
                    'product.public.category', 'search',
                    [[('name', '=', categoria)]]
                )

                if not categoria_ids:
                    emit_progress(f"⚠️ Categoría '{categoria}' no encontrada en Odoo.")
                    continue

                product_ids = models.execute_kw(
                    db, uid, password,
                    'product.template', 'search',
                    [[('categ_id', 'in', categoria_ids)]]
                )

                if product_ids:
                    models.execute_kw(
                        db, uid, password,
                        'product.template', 'write',
                        [product_ids, {'is_published': True}]
                    )
                    productos_publicados += len(product_ids)
                    emit_progress(
                        f"✅ Publicados {len(product_ids)} productos de la categoría '{categoria}' ({i}/{total_categorias})")
                else:
                    emit_progress(
                        f"⚠️ No se encontraron productos para la categoría '{categoria}' ({i}/{total_categorias})")

            emit_progress(f"🎉 Publicación por categorías completada: {productos_publicados} productos publicados")

        else:
            if "SKU" not in df.columns:
                emit_progress("❌ No se encontró la columna 'SKU' en el Excel.")
                if event_manager:
                    event_manager.emit('operation_error', "Columna 'SKU' no encontrada")
                return

            skus = df["SKU"].dropna().astype(str).tolist()
            total_skus = len(skus)
            emit_progress(f"🏷️ Publicando productos por SKU ({total_skus} valores)...")

            productos_publicados = 0
            for i, sku in enumerate(skus, 1):
                emit_progress(f"🏷️ Procesando SKU {i}/{total_skus}: {sku}")

                product_ids = models.execute_kw(
                    db, uid, password,
                    'product.template', 'search',
                    [[('default_code', '=', sku)]]
                )

                if product_ids:
                    models.execute_kw(
                        db, uid, password,
                        'product.template', 'write',
                        [product_ids, {'is_published': True}]
                    )
                    productos_publicados += 1
                    emit_progress(f"✅ Producto SKU '{sku}' publicado ({i}/{total_skus})")
                else:
                    emit_progress(f"⚠️ Producto con SKU '{sku}' no encontrado ({i}/{total_skus})")

            emit_progress(f"🎉 Publicación por SKU completada: {productos_publicados} productos publicados")

        # Emitir completion al final
        if event_manager:
            event_manager.emit('operation_completed', "Publicación finalizada correctamente")

    except Exception as e:
        emit_progress(f"❌ Error en publicar_productos: {e}")
        if event_manager:
            event_manager.emit('operation_error', f"Error en publicación: {e}")


# endregion

# region Ventana Conexión
import os
def abrir_ajustes():
    from tkinter import Toplevel, messagebox
    from pathlib import Path
    import tkinter as tk
    import json

    config_path = os.path.expanduser(
        "~/Documents/SMI Files/Data/perfiles.json")  # Path(__file__).parent / "config.json"
    ruta_script = Path(__file__).parent / "App_Connection.py"

    ajustes_win = Toplevel()
    ajustes_win.title("Ajustes")
    ajustes_win.geometry("400x300")
    ajustes_win.resizable(False, False)

    ajustes_win.transient()
    ajustes_win.grab_set()
    ajustes_win.focus_force()

    config_backup_path = os.path.expanduser("~/Documents/SMI Files/Data/perfiles_backup.json")

    def cargar_config():
        ruta = os.path.expanduser("~/Documents/SMI Files/Data")
        os.makedirs(ruta, exist_ok=True)
        if config_path:
            with open(config_path, "r") as f:
                try:
                    data = json.load(f)
                    perfil = data.get("current_profile", "default")
                    perfiles = data.get("profiles", {})
                    if not perfiles:
                        raise ValueError
                    return perfiles.get(perfil, {}), perfiles, perfil
                except:
                    pass

        # Si no hay archivo o está mal formado, crear uno por defecto
        default_config = {
            "current_profile": "default",
            "profiles": {
                "default": {
                    "URL": "",
                    "Database": "",
                    "Usuario": "",
                    "Password": ""
                }
            }
        }
        with open(config_path, "w") as f:
            json.dump(default_config, f, indent=4)

        return default_config["profiles"]["default"], default_config["profiles"], "default"

    # Cargar configuración y perfiles
    valores_actuales, perfiles, perfil_actual = cargar_config()

    perfil_var = tk.StringVar(value=perfil_actual or "default")

    # Selector de perfil arriba a la izquierda
    tk.Label(ajustes_win, text="Perfil:").place(x=10, y=10)
    perfil_dropdown = tk.OptionMenu(ajustes_win, perfil_var, *perfiles.keys())
    perfil_dropdown.place(x=60, y=5)

    # Entry + botón para crear nuevo perfil
    nuevo_perfil_var = tk.StringVar()
    tk.Entry(ajustes_win, textvariable=nuevo_perfil_var, width=12).place(x=240, y=10)
    tk.Button(ajustes_win, text="Nuevo", command=lambda: crear_nuevo_perfil()).place(x=330, y=5)

    # Frame central para los campos
    frame_central = tk.Frame(ajustes_win)
    frame_central.place(relx=0.5, rely=0.52, anchor="center")

    # Campos
    campos = {
        "URL": tk.StringVar(value=valores_actuales.get("URL", "")),
        "Database": tk.StringVar(value=valores_actuales.get("Database", "")),
        "Usuario": tk.StringVar(value=valores_actuales.get("Usuario", "")),
        "Password": tk.StringVar(value=valores_actuales.get("Password", ""))
    }

    def actualizar_guardar_si_cambios(*args):
        colocar_botones()

    for var in campos.values():
        var.trace_add("write", actualizar_guardar_si_cambios)

    for idx, (label_text, var) in enumerate(campos.items()):
        tk.Label(frame_central, text=label_text + ":").grid(row=idx, column=0, padx=10, pady=6, sticky='e')
        entry = tk.Entry(frame_central, textvariable=var, width=30)
        if label_text.lower() == "password":
            entry.config(show="*")
        entry.grid(row=idx, column=1, padx=10, pady=6)

    # Mensaje informativo
    mensaje_label = tk.Label(frame_central, text="", fg="green")
    mensaje_label.grid(row=4, column=0, columnspan=2, pady=5)

    def colocar_botones():
        y = 250
        x0 = 70 if perfil_var.get() != "default" else 100
        if hubo_cambios():
            guardar_btn.place(x=x0, y=y)
        else:
            guardar_btn.place_forget()

        conectar_btn.place(x=x0 + 100, y=y)

        if perfil_var.get() != "default":
            eliminar_btn.place(x=x0 + 200, y=y)
        else:
            eliminar_btn.place_forget()

    def hubo_cambios():
        return any(campos[k].get() != valores_actuales.get(k, "") for k in campos)

    def actualizar_campos_desde_perfil(*args):
        nonlocal valores_actuales
        perfil = perfil_var.get()
        datos = perfiles.get(perfil, {})
        for campo, var in campos.items():
            var.set(datos.get(campo, ""))
        valores_actuales = datos
        colocar_botones()

    perfil_var.trace_add("write", actualizar_campos_desde_perfil)

    def crear_nuevo_perfil():
        nuevo = nuevo_perfil_var.get().strip()
        if not nuevo:
            messagebox.showerror("Error", "Debes ingresar un nombre para el nuevo perfil.")
            return
        if nuevo in perfiles:
            messagebox.showinfo("Perfil existente", "Ese perfil ya existe.")
            return

        perfiles[nuevo] = {"URL": "", "Database": "", "Usuario": "", "Password": ""}
        perfil_var.set(nuevo)
        nuevo_perfil_var.set("")

        guardar_config(perfiles[nuevo], False, nuevo)
        perfil_dropdown["menu"].add_command(label=nuevo, command=tk._setit(perfil_var, nuevo))
        actualizar_campos_desde_perfil()

    def eliminar_perfil():
        perfil = perfil_var.get()
        if perfil == "default":
            messagebox.showwarning("Aviso", "No se puede eliminar el perfil 'default'.")
            return

        confirm = messagebox.askyesno("Eliminar perfil", f"¿Eliminar perfil '{perfil}'?")
        if confirm:
            del perfiles[perfil]
            nuevo = "default" if "default" in perfiles else list(perfiles.keys())[0]
            perfil_var.set(nuevo)
            guardar_config(perfiles[nuevo], False, nuevo, True)
            perfil_dropdown["menu"].delete(0, "end")
            for p in perfiles:
                perfil_dropdown["menu"].add_command(label=p, command=tk._setit(perfil_var, p))
            actualizar_campos_desde_perfil()

    def guardar_config(config, set_connection, perfil="default", was_deleted = False):

        data = {"current_profile": perfil, "profiles": perfiles}
        perfiles[perfil] = config
        with open(config_path, "w") as f:
            json.dump(data, f, indent=4)
        if not was_deleted:
            with open(config_backup_path, "w") as f:
                json.dump(data, f, indent=4)
            #ruta = os.path.expanduser("~/Documents/SMI Files/Data/perfiles_backup.json")
            #shutil.copy(config_backup_path, ruta)


        if set_connection:
            with open(ruta_script, "w", encoding="utf-8") as f:
                f.write(f"""import xmlrpc.client

url = '{config["URL"]}'
db = '{config["Database"]}'
username = '{config["Usuario"]}'
password = '{config["Password"]}'

common = None
uid = None
models = None

def conectar():
    global common, uid, models
    try:
        common = xmlrpc.client.ServerProxy(f'{{url}}/xmlrpc/2/common', allow_none=True)
        uid = common.authenticate(db, username, password, {{}})
        if uid:
            print(f'Conectado como {{username}} (uid: {{uid}})')
            models = xmlrpc.client.ServerProxy(f'{{url}}/xmlrpc/2/object', allow_none=True)
        else:
            print('Error de autenticación')
        return uid
    except Exception as e:
        print(f'❌ Error en la conexión: {{e}}')
        return None
""")

    def guardar_ajustes():
        config = {campo: var.get() for campo, var in campos.items()}
        perfil = perfil_var.get()
        if not all(config.values()):
            mensaje_label.config(text="Todos los campos deben estar completos.", fg="red")
            return
        guardar_config(config, False, perfil)
        mensaje_label.config(text="Ajustes guardados correctamente.", fg="green")
        actualizar_campos_desde_perfil()

    def conectar():
        global is_connected

        config = {campo: var.get() for campo, var in campos.items()}
        perfil = perfil_var.get()
        db = config["Database"]
        cambios = hubo_cambios()
        try:
            guardar_config(config, True, perfil)

            import importlib.util
            import sys
            spec = importlib.util.spec_from_file_location("App_Connection", str(ruta_script))
            App_Connection = importlib.util.module_from_spec(spec)
            sys.modules["App_Connection"] = App_Connection
            spec.loader.exec_module(App_Connection)

            uid = App_Connection.conectar()
            is_connected = uid
            if uid:
                mensaje_label.config(text=f"Conectado a: {db}", fg="green")
                ajustes_win.update()
                if cambios:
                    mensaje_label.config(text="Reiniciando...", fg="blue")
                    ajustes_win.update()
                    import time; time.sleep(1.2)
                    os.execl(sys.executable, sys.executable, *sys.argv)
                else: ajustes_win.destroy()
            else:
                mensaje_label.config(text=f"No se pudo conectar a: {db}", fg="red")

        except Exception as e:
            mensaje_label.config(text=f"Error al conectar: {e}", fg="red")

    # Botones
    guardar_btn = tk.Button(ajustes_win, text="Guardar", command=lambda: guardar_ajustes())
    conectar_btn = tk.Button(ajustes_win, text="Conectar", command=lambda: conectar())
    eliminar_btn = tk.Button(ajustes_win, text="Eliminar", command=lambda: eliminar_perfil())
    colocar_botones()
    ajustes_win.wait_window()

# endregion


def main_import_with_event_manager(event_manager):
    global is_connected, atributos_cache, valores_cache

    def emit_progress(message):
        if event_manager:
            event_manager.emit('progress_update', message)
            print(message)
        else:
            print(message)

    emit_progress("🔧 Iniciando proceso de conexión...")
    abrir_ajustes()

    if is_connected:
        from App_Connection import models, db, uid, password

        ConnectionParams = namedtuple("ConnectionParams", ["models", "db", "uid", "password"])
        params = ConnectionParams(models, db, uid, password)

        emit_progress("✅ Conexión establecida con Odoo")

        ruta = Utils.seleccionar_excel()

        atributos_cache = Utils.cargar_atributos_existentes(params)
        valores_cache = Utils.cargar_valores_atributos_existentes(params)
        odoo16_connect = True

        url_src = 'https://odoo16.optimaluz.com/'#
        db_src = 'Real'#Test
        username_src = 'jcoronado@optimaluz.com'
        password_src = 'AlAi4ever' #AlAi4ever@optimaluz.com

        common = xmlrpc.client.ServerProxy(f"{url_src}/xmlrpc/2/common") if odoo16_connect else ""
        uid_src = common.authenticate(db_src, username_src, password_src, {}) if odoo16_connect else ""
        models_src = xmlrpc.client.ServerProxy(f"{url_src}/xmlrpc/2/object") if odoo16_connect else ""


        # region REGIÓN NORMALIZAR

        # ---------------------------------------------------------------------------
        # 1) NORMALIZAR NOMBRES PRODUCTOS B2B
        # ---------------------------------------------------------------------------
        def normalizar_nombres_b2b(TEST=False, TEST_DEFAULT_CODE=""):
            import re, ast
            from App_Connection import models, db, uid, password

            CATEGORY_NORMALIZERS = {
                #"TIRAS Y NEONES LED": "normalizar_tiras_y_neones",
                #"FUENTES DE ALIMENTACIÓN": "normalizar_fuentes",
                #"PROYECTORES LED": "normalizar_proyectores",
                #"BOMBILLAS LED": "normalizar_bombillas",
                #"PERFILERIA DE ALUMINIO PARA TIRAS Y NEONES LED": "normalizar_perfiles",
                "DESCATALOGADOS": "normalizar_descatalogados",
            }

            CATS_PERFILES = [
                "PERFILES ARQUITECTÓNICOS PARA TIRAS LED",
                "PERFILES DE SUPERFICIE PARA TIRAS Y NEONES LED",
                "PERFILES DE EMPOTRAR PARA TIRAS LED",
                "PERFILES DE ENCASTRAR PARA TIRAS LED",
                "PERFILERIA DE ALUMINIO DESCATALOGADA",
            ]

            CATS_DESCATALOGADOS = [
                "TIRAS Y NEONES LED DESCATALOGADOS",
                "FUENTES DE ALIMENTACIÓN DESCATALOGADAS",
                "PROYECTORES LED DESCATALOGADOS",
                "BOMBILLAS LED DESCATALOGADAS",
            ]

            def cargar_attrs_excel(ruta_excel):
                df = pd.read_excel(ruta_excel).fillna("")
                data = {}

                for _, row in df.iterrows():
                    sku = str(row.get("SKU", "")).strip()
                    attrs_txt = str(row.get("Atributos", "")).strip()

                    if not sku or not attrs_txt:
                        continue

                    try:
                        attrs_dict = ast.literal_eval(attrs_txt)
                        data[sku] = attrs_dict
                    except Exception as e:
                        print(f"Error leyendo atributos Excel SKU {sku}: {e}")

                return data

            ATTRS_EXCEL_POR_SKU = cargar_attrs_excel(ruta)

            def crear_atributo_producto(product_id, attr_name, value_name):
                attr_ids = models.execute_kw(
                    db, uid, password,
                    "product.attribute",
                    "search",
                    [[["name", "=", attr_name]]],
                    {"limit": 1, "context": {"lang": "es_ES"}}
                )

                if not attr_ids:
                    print(f"Atributo no existe en Odoo: {attr_name}")
                    return False

                attr_id = attr_ids[0]

                value_ids = models.execute_kw(
                    db, uid, password,
                    "product.attribute.value",
                    "search",
                    [[
                        ["name", "=", value_name],
                        ["attribute_id", "=", attr_id],
                    ]],
                    {"limit": 1, "context": {"lang": "es_ES"}}
                )

                if not value_ids:
                    value_ids = models.execute_kw(
                        db, uid, password,
                        "product.attribute.value",
                        "create",
                        [{
                            "name": value_name,
                            "attribute_id": attr_id,
                        }],
                        {"context": {"lang": "es_ES"}}
                    )
                    value_id = value_ids
                else:
                    value_id = value_ids[0]

                models.execute_kw(
                    db, uid, password,
                    "product.template.attribute.line",
                    "create",
                    [{
                        "product_tmpl_id": product_id,
                        "attribute_id": attr_id,
                        "value_ids": [(6, 0, [value_id])],
                    }],
                    {"context": {"lang": "es_ES"}}
                )

                return True

            def completar_attrs_faltantes_desde_excel(product, attrs_excel_por_sku, attrs_a_completar):
                sku = str(product.get("default_code", "")).strip()
                attrs_producto = product["attrs"]

                attrs_excel = attrs_excel_por_sku.get(sku)

                if not attrs_excel:
                    return attrs_producto, attrs_a_completar

                siguen_faltando = []

                for attr in attrs_a_completar:
                    valor_excel = attrs_excel.get(attr)

                    if not valor_excel:
                        siguen_faltando.append(attr)
                        continue

                    creado = crear_atributo_producto(
                        product["id"],
                        attr,
                        str(valor_excel).strip()
                    )

                    if creado:
                        attrs_producto[attr] = str(valor_excel).strip()
                    else:
                        siguen_faltando.append(attr)

                return attrs_producto, siguen_faltando

            def normalizar_nombres_productos(test=True, default_code_test=None):
                for category_name, normalizer_name in CATEGORY_NORMALIZERS.items():
                    productos = []

                    if normalizer_name == "normalizar_perfiles":
                        productos = buscar_productos_cats(
                            False,
                            test=test,
                            default_code_test=default_code_test
                        )

                    elif normalizer_name == "normalizar_descatalogados":
                        productos = buscar_productos_cats(
                            True,
                            test=test,
                            default_code_test=default_code_test
                        )

                    else:
                        category_ids = models.execute_kw(
                            db, uid, password,
                            "product.public.category",
                            "search",
                            [[["name", "=", category_name]]],
                            {"context": {"lang": "es_ES"}}
                        )

                        if not category_ids:
                            print(f"Categoría no encontrada: {category_name}")
                            continue

                        domain = [
                            ["public_categ_ids", "child_of", category_ids],
                            ["x_new_name", "=", False],
                        ]

                        if test and default_code_test:
                            domain.append(["default_code", "=", default_code_test])

                        product_ids = models.execute_kw(
                            db, uid, password,
                            "product.template",
                            "search",
                            [domain],
                            {"context": {"lang": "es_ES"}}
                        )

                        print(f"{category_name}: {len(product_ids)} productos encontrados")

                        productos = [(product_id, category_name) for product_id in product_ids]

                    for product_id, cat_name in productos:
                        product = leer_producto(product_id)

                        if normalizer_name == "normalizar_descatalogados":
                            product["categoria_descatalogado"] = cat_name

                        normalizer = {
                            "normalizar_tiras_y_neones": normalizar_tiras_y_neones,
                            "normalizar_fuentes": normalizar_fuentes,
                            "normalizar_proyectores": normalizar_proyectores,
                            "normalizar_bombillas": normalizar_bombillas,
                            "normalizar_perfiles": normalizar_perfiles,
                            "normalizar_descatalogados": normalizar_descatalogados,
                        }[normalizer_name]

                        nuevo_name = normalizer(product)

                        if not nuevo_name:
                            continue

                        if nuevo_name == product["name"]:
                            # print(f"Sin cambios: {product.get('default_code')}")
                            continue

                        models.execute_kw(
                            db, uid, password,
                            "product.template",
                            "write",
                            [
                                [product_id],
                                {
                                    "name": nuevo_name,
                                    "x_new_name": True,
                                }
                            ],
                            {
                                "context": {
                                    "lang": "es_ES"
                                }
                            }
                        )

                        print(f"Actualizado {product.get('default_code')}: {nuevo_name}")

            def buscar_productos_cats(descatalogados, test=True, default_code_test=None):
                productos = []
                product_ids_unicos = set()

                cats_template = CATS_DESCATALOGADOS if descatalogados else CATS_PERFILES

                for cat_name in cats_template:
                    cat_ids = models.execute_kw(
                        db, uid, password,
                        "product.public.category",
                        "search",
                        [[["name", "=", cat_name]]],
                        {"context": {"lang": "es_ES"}}
                    )

                    if not cat_ids:
                        print(f"Categoría no encontrada: {cat_name}")
                        continue

                    domain = [
                        ["public_categ_ids", "child_of", cat_ids],
                        ["x_new_name", "=", False],
                    ]

                    if test and default_code_test:
                        domain.append(["default_code", "=", default_code_test])

                    ids = models.execute_kw(
                        db, uid, password,
                        "product.template",
                        "search",
                        [domain],
                        {"context": {"lang": "es_ES"}}
                    )

                    print(f"{cat_name}: {len(ids)} productos encontrados")

                    for product_id in ids:
                        if product_id in product_ids_unicos:
                            continue

                        product_ids_unicos.add(product_id)
                        productos.append((product_id, cat_name))

                return productos

            def leer_producto(product_id):
                fields = [
                    "name",
                    "default_code",
                    "product_brand_id",
                    "attribute_line_ids",
                ]

                product = models.execute_kw(
                    db, uid, password,
                    "product.template",
                    "read",
                    [[product_id], fields],
                    {"context": {"lang": "es_ES"}}
                )[0]

                product["attrs"] = leer_atributos_producto(product["attribute_line_ids"])
                return product

            def leer_atributos_producto(attribute_line_ids):
                attrs = {}

                if not attribute_line_ids:
                    return attrs

                lines = models.execute_kw(
                    db, uid, password,
                    "product.template.attribute.line",
                    "read",
                    [attribute_line_ids, ["attribute_id", "value_ids"]],
                    {"context": {"lang": "es_ES"}}
                )

                for line in lines:
                    attr_name = line["attribute_id"][1]

                    value_ids = line.get("value_ids", [])
                    if not value_ids:
                        continue

                    values = models.execute_kw(
                        db, uid, password,
                        "product.attribute.value",
                        "read",
                        [value_ids, ["name"]],
                        {"context": {"lang": "es_ES"}}
                    )

                    attrs[attr_name] = values[0]["name"]

                return attrs

            def normalizar_tiras_y_neones(product):
                def obtener_voltaje(attrs):
                    return (
                            attrs.get("Tensión")
                            or attrs.get("Voltaje de entrada")
                    )

                name_actual = product["name"]
                attrs = product["attrs"]

                voltaje = obtener_voltaje(attrs)
                temperatura_color = obtener_temperatura_color(attrs)

                siguen_faltando = []

                if not voltaje:
                    siguen_faltando.append("Tensión/Voltaje de entrada")

                if not temperatura_color:
                    siguen_faltando.append("Temperatura de color/Tono de luz")

                if not attrs.get("Potencia/m"):
                    siguen_faltando.append("Potencia/m")

                ATTRS_RECUPERABLES_EXCEL = [
                    "Tipo LED",
                    "Cantidad de LED",
                    "Flujo luminoso/m",
                    #"Anchura",
                    #"CRI",
                    "Grado de protección IP",
                    "Longitud",
                ]
                faltantes = [
                    attr for attr in ATTRS_RECUPERABLES_EXCEL
                    if not attrs.get(attr)
                ]

                if faltantes:
                    attrs, faltantes_excel = completar_attrs_faltantes_desde_excel(
                        product,
                        ATTRS_EXCEL_POR_SKU,
                        faltantes
                    )

                    siguen_faltando.extend(faltantes_excel)

                if siguen_faltando:
                    print(
                        f"Saltado {product.get('default_code')}: "
                        f"faltan atributos {', '.join(siguen_faltando)}"
                    )
                    return False

                marca = obtener_marca(product)
                bloque_marca = ""

                if "PRO" in name_actual.upper():
                    bloque_marca = "PRO"
                elif marca.upper() in ["V-TAC", "VTAC", "V TAC"]:
                    bloque_marca = "V-TAC"

                tipo_producto = "NEÓN FLEX LED" if "NEÓN" in name_actual.upper() or "NEON" in name_actual.upper() else "TIRA LED"

                partes = [
                    tipo_producto,
                    bloque_marca,
                    normalizar_voltaje(voltaje),
                    normalizar_simple(attrs["Potencia/m"]),
                ]

                if "CREE" in name_actual.upper():
                    partes.append("CHIP CREE")

                partes += [
                    normalizar_simple(attrs["Tipo LED"]),
                    normalizar_cantidad_led(attrs["Cantidad de LED"]),
                    normalizar_flujo(attrs["Flujo luminoso/m"]),
                    normalizar_anchura(attrs.get("Anchura", "")),
                    normalizar_cri(attrs.get("CRI", "")),
                    normalizar_simple(attrs["Grado de protección IP"]),
                    normalizar_simple(temperatura_color),
                    normalizar_longitud(attrs["Longitud"]),#(attrs.get("Longitud", "5 metros")),
                ]

                cuerpo = " ".join([p for p in partes if p]).strip()
                cuerpo = limpiar_espacios(cuerpo)

                return aplicar_prefijo_protegido(name_actual, cuerpo)

            def normalizar_fuentes(product):
                name_actual = product["name"]
                attrs = product["attrs"]

                ATTRS_RECUPERABLES_EXCEL_FUENTES = [
                    "Salida",
                    "Potencia",
                    "Corriente",
                    "Grado de protección IP",
                    #"Dimensiones",
                ]

                faltantes = [
                    attr for attr in ATTRS_RECUPERABLES_EXCEL_FUENTES
                    if not attrs.get(attr)
                ]

                if faltantes:
                    attrs, siguen_faltando = completar_attrs_faltantes_desde_excel(
                        product,
                        ATTRS_EXCEL_POR_SKU,
                        faltantes
                    )

                    if siguen_faltando:
                        print(
                            f"Saltado {product.get('default_code')}: "
                            f"faltan atributos {', '.join(siguen_faltando)}"
                        )
                        return False

                marca = obtener_marca(product)
                bloque_marca = ""

                if marca.upper() in ["V-TAC", "VTAC", "V TAC"]:
                    bloque_marca = "V-TAC"

                name_upper = name_actual.upper()

                compacta = "COMPACTA" if "COMPACTA" in name_upper else ""

                regulable = ""
                if "REGULABLE DALI" in name_upper:
                    regulable = "REGULABLE DALI"
                elif "DALI" in name_upper:
                    regulable = "REGULABLE DALI"
                elif "REGULABLE" in name_upper:
                    regulable = "REGULABLE"

                partes = [
                    "FUENTE DE ALIMENTACIÓN",
                    compacta,
                    bloque_marca,
                    normalizar_voltaje(attrs["Salida"]),
                    normalizar_simple(attrs["Potencia"]),
                    normalizar_simple(attrs["Corriente"]),
                    normalizar_simple(attrs["Grado de protección IP"]),
                    regulable,
                    normalizar_dimensiones(attrs.get("Dimensiones", "")),#(attrs["Dimensiones"]),
                ]

                cuerpo = " ".join([p for p in partes if p]).strip()
                cuerpo = limpiar_espacios(cuerpo)

                return aplicar_prefijo_protegido(name_actual, cuerpo)

            def normalizar_proyectores(product):
                def obtener_color_cuerpo_desde_name(name_actual):
                    match = re.search(r"\bCUERPO\s+([A-ZÁÉÍÓÚÜÑ]+)", name_actual.upper())

                    if not match:
                        return ""

                    return match.group(1).capitalize()

                name_actual = product["name"]
                attrs = product["attrs"]

                temperatura_color = obtener_temperatura_color(attrs)

                siguen_faltando = []

                if not temperatura_color:
                    siguen_faltando.append("Temperatura de color/Tono de luz")

                ATTRS_RECUPERABLES_EXCEL_PROYECTORES = [
                    "Potencia",
                    "Tipo LED",
                    "Flujo luminoso (lm)",
                    #"Ángulo de Apertura",
                    "Color del cuerpo",
                    "Grado de protección IP",
                ]

                if not attrs.get("Color del cuerpo"):
                    color_desde_name = obtener_color_cuerpo_desde_name(name_actual)

                    if color_desde_name:
                        attrs["Color del cuerpo"] = color_desde_name

                faltantes = [
                    attr for attr in ATTRS_RECUPERABLES_EXCEL_PROYECTORES
                    if not attrs.get(attr)
                ]

                if faltantes:
                    attrs, faltantes_excel = completar_attrs_faltantes_desde_excel(
                        product,
                        ATTRS_EXCEL_POR_SKU,
                        faltantes
                    )

                    siguen_faltando.extend(faltantes_excel)

                if siguen_faltando:
                    print(
                        f"Saltado {product.get('default_code')}: "
                        f"faltan atributos {', '.join(siguen_faltando)}"
                    )
                    return False

                marca = obtener_marca(product)
                bloque_marca = ""

                if marca.upper() in ["V-TAC", "VTAC", "V TAC"]:
                    bloque_marca = "V-TAC"

                regulable = "REGULABLE" if "REGULABLE" in name_actual.upper() else ""

                partes = [
                    "PROYECTOR LED",
                    bloque_marca,
                    normalizar_simple(attrs["Potencia"]),
                    normalizar_tipo_led(attrs["Tipo LED"]),
                    normalizar_lumen(attrs["Flujo luminoso (lm)"]),
                    normalizar_simple(attrs.get("Ángulo de Apertura", "")),#(attrs["Ángulo de Apertura"]),
                    normalizar_color_cuerpo(attrs["Color del cuerpo"]),
                    normalizar_simple(attrs["Grado de protección IP"]),
                    regulable,
                    normalizar_simple(temperatura_color),
                ]

                cuerpo = " ".join([p for p in partes if p]).strip()
                cuerpo = limpiar_espacios(cuerpo)

                return aplicar_prefijo_protegido(name_actual, cuerpo)

            def normalizar_bombillas(product):
                name_actual = product["name"]
                attrs = product["attrs"]

                temperatura_color = obtener_temperatura_color(attrs)

                siguen_faltando = []

                if not temperatura_color:
                    siguen_faltando.append("Temperatura de color/Tono de luz")

                ATTRS_RECUPERABLES_EXCEL_BOMBILLAS = [
                    "Casquillo",
                    "Ampolla exterior",
                    "Potencia",
                    "Tipo LED",
                    "Flujo luminoso (lm)",
                    #"Ángulo de Apertura",
                    #"Dimensiones",
                ]

                faltantes = [
                    attr for attr in ATTRS_RECUPERABLES_EXCEL_BOMBILLAS
                    if not attrs.get(attr)
                ]

                if faltantes:
                    attrs, faltantes_excel = completar_attrs_faltantes_desde_excel(
                        product,
                        ATTRS_EXCEL_POR_SKU,
                        faltantes
                    )

                    siguen_faltando.extend(faltantes_excel)

                if siguen_faltando:
                    print(
                        f"Saltado {product.get('default_code')}: "
                        f"faltan atributos {', '.join(siguen_faltando)}"
                    )
                    return False

                marca = obtener_marca(product)
                bloque_marca = ""

                if marca.upper() in ["V-TAC", "VTAC", "V TAC"]:
                    bloque_marca = "V-TAC"

                partes = [
                    "BOMBILLA LED",
                    bloque_marca,
                    normalizar_simple(attrs["Casquillo"]),
                    normalizar_simple(attrs["Ampolla exterior"]),
                    normalizar_simple(attrs["Potencia"]),
                    normalizar_tipo_led(attrs["Tipo LED"]),
                    normalizar_lumen(attrs["Flujo luminoso (lm)"]),
                    normalizar_simple(attrs.get("Ángulo de Apertura", "")),#(attrs["Ángulo de Apertura"]),
                    normalizar_dimensiones(attrs.get("Dimensiones", "")),#(attrs["Dimensiones"]),
                    normalizar_simple(temperatura_color),
                ]

                cuerpo = " ".join([p for p in partes if p]).strip()
                cuerpo = limpiar_espacios(cuerpo)

                return aplicar_prefijo_protegido(name_actual, cuerpo)

            def normalizar_perfiles(product):
                def obtener_tipo_perfil(name_upper):
                    if "SUPERFICIE" in name_upper:
                        if "FLEXIBLE" in name_upper:
                            return "DE SUPERFICIE FLEXIBLE"
                        return "DE SUPERFICIE"

                    if "EMPOTRAR" in name_upper or "EMPOTRABLE" in name_upper:
                        return "DE EMPOTRAR"

                    if "ENCASTRAR" in name_upper:
                        return "DE ENCASTRAR"

                    return ""

                def obtener_color_perfil(name_upper):
                    if "NEGRO" in name_upper or "C. NEGRO" in name_upper or "LACADO NEGRO" in name_upper:
                        return "COLOR NEGRO"

                    if "BLANCO" in name_upper or "C. BLANCO" in name_upper or "LACADO BLANCO" in name_upper:
                        return "COLOR BLANCO"

                    return "COLOR ALUMINIO"

                def obtener_difusor_perfil(name_upper):
                    if "DIFUSOR NEGRO" in name_upper:
                        return "DIFUSOR NEGRO"

                    if "DIFUSOR OPAL" in name_upper or "OPAL" in name_upper:
                        return "DIFUSOR OPAL"

                    return ""

                def obtener_dimensiones_perfil(name_actual):
                    import re

                    texto = str(name_actual).replace("×", "x").replace("*", "")
                    texto = texto.replace("mm.", "mm").replace("MM.", "MM")

                    patrones = [
                        r"\d+(?:[.,]\d+)?\s*[xX]\s*\d+(?:[.,]\d+)?\s*[xX]\s*\d+(?:[.,]\d+)?\s*mm",
                        r"\d+(?:[.,]\d+)?\s*[xX]\s*\d+(?:[.,]\d+)?H[xX]\s*\d+(?:[.,]\d+)?\s*mm",
                        r"\d+\s*m\b",
                        r"\d+\s*mm\b",
                    ]

                    for patron in patrones:
                        match = re.search(patron, texto, re.IGNORECASE)
                        if match:
                            dim = match.group(0)
                            dim = dim.replace(" ", "")
                            dim = dim.replace("mm", "MM")
                            dim = dim.replace("Mm", "MM")
                            dim = dim.replace("m", "M")
                            return dim

                    return ""

                def obtener_texto_parentesis(name_actual):
                    import re

                    matches = re.findall(r'\([^)]*\)', name_actual)

                    if not matches:
                        return ""

                    return " ".join(m.upper() for m in matches)

                name_actual = product["name"]
                name_upper = name_actual.upper()

                if "POLICARBONATO" in name_upper or "ARQUITECTÓNICO" in name_upper or "ARQUITECTONICO" in name_upper:
                    print(f"Saltado {product.get('default_code')}: policarbonato/arquitectónico")
                    return False

                tipo = obtener_tipo_perfil(name_upper)

                if not tipo:
                    print(f"Saltado {product.get('default_code')}: no se detecta tipo de perfil")
                    return False

                marca = obtener_marca(product)

                if marca.upper() in ["V-TAC", "VTAC", "V TAC"]:
                    bloque_marca = "V-TAC"
                else:
                    bloque_marca = "PRO"

                color = obtener_color_perfil(name_upper)
                difusor = obtener_difusor_perfil(name_upper)
                dimensiones = obtener_dimensiones_perfil(name_actual)
                texto_parentesis = obtener_texto_parentesis(name_actual)

                if not dimensiones:
                    print(f"Saltado {product.get('default_code')}: no se detectan dimensiones")
                    return False

                partes = [
                    "PERFIL DE ALUMINIO",
                    tipo,
                    bloque_marca,
                    color,
                    difusor,
                    dimensiones,
                    texto_parentesis,
                ]

                cuerpo = " ".join([p for p in partes if p]).strip()
                cuerpo = limpiar_espacios(cuerpo)

                return aplicar_prefijo_protegido(name_actual, cuerpo)

            def normalizar_descatalogados(product):
                cat_normalizers = {
                    "TIRAS Y NEONES LED DESCATALOGADOS": normalizar_tiras_y_neones,
                    "FUENTES DE ALIMENTACIÓN DESCATALOGADAS": normalizar_fuentes,
                    "PROYECTORES LED DESCATALOGADOS": normalizar_proyectores,
                    "BOMBILLAS LED DESCATALOGADAS": normalizar_bombillas,
                }

                cat_name = product.get("categoria_descatalogado")

                normalizer = cat_normalizers.get(cat_name)

                if not normalizer:
                    print(
                        f"Saltado {product.get('default_code')}: "
                        f"categoría descatalogada no reconocida ({cat_name})"
                    )
                    return False

                return normalizer(product)

            def obtener_marca(product):
                marca = product.get("product_brand_id")

                if isinstance(marca, list) and len(marca) >= 2:
                    return marca[1]

                return ""

            def aplicar_prefijo_protegido(name_actual, nuevo_cuerpo):
                if "]" in name_actual:
                    prefijo = name_actual[:name_actual.index("]") + 1]
                    return f"{prefijo} {nuevo_cuerpo}"

                if name_actual.startswith("#"):
                    return f"# {nuevo_cuerpo}"

                return nuevo_cuerpo

            def normalizar_voltaje(valor):
                valor = valor.upper().replace(" ", "")

                match = re.search(r"DC:?(\d+V)", valor)
                if match:
                    return f"{match.group(1)} DC"

                return valor.replace(":", " ")

            def normalizar_cantidad_led(valor):
                valor = str(valor).upper()

                match = re.search(r'(\d+)', valor)

                if not match:
                    return ""

                return f"{match.group(1)}LED/M"

            def normalizar_flujo(valor):
                valor = str(valor).upper()

                match = re.search(r'(\d+)', valor)

                if not match:
                    return ""

                return f"{match.group(1)}LM/M"

            def normalizar_anchura(valor):
                if not valor:
                    return ""
                valor = valor.upper()
                valor = valor.replace(" ", "")
                return f"PCB {valor}"

            def normalizar_cri(valor):
                if not valor:
                    return ""
                valor = valor.upper()
                valor = valor.replace("≥", "")
                valor = valor.replace(">=", "")
                valor = valor.replace("CRI", "")
                valor = valor.strip()
                return f"CRI{valor}"

            def normalizar_longitud(valor):
                valor = valor.strip()
                return f"(Rollo de {valor})"

            def normalizar_dimensiones(valor):
                if not valor:
                    return ""
                valor = str(valor).upper()
                valor = valor.replace(" ", "")
                valor = valor.replace("×", "x")
                valor = valor.replace("*", "x")
                valor = valor.replace("mm.", "MM")
                valor = valor.replace("milímetros", "MM")
                valor = valor.replace("milimetros", "MM")

                if not valor.endswith("mm"):
                    valor += "MM"

                return valor

            def normalizar_simple(valor):
                if not valor:
                    return ""
                return limpiar_espacios(str(valor).upper().strip())

            def limpiar_espacios(texto):
                return re.sub(r"\s+", " ", texto).strip()

            def normalizar_tipo_led(valor):
                valor = normalizar_simple(valor)

                contiene_samsung = "SAMSUNG" in valor
                contiene_cree = "CREE" in valor
                contiene_smd = "SMD" in valor

                partes = []

                if contiene_samsung:
                    partes.append("CHIP SAMSUNG")
                elif contiene_cree:
                    partes.append("CHIP CREE")
                elif "CHIP" in valor:
                    partes.append("CHIP")

                if contiene_smd:
                    partes.append("SMD")

                if partes:
                    return " ".join(partes)

                return valor

            def normalizar_lumen(valor):
                valor = str(valor).upper()
                valor = valor.replace(" ", "")
                valor = valor.replace("LM.", "LM")

                if not valor.endswith("LM"):
                    valor += "LM"

                return valor

            def normalizar_color_cuerpo(valor):
                valor = normalizar_simple(valor)
                return f"COLOR {valor}" if valor else ""

            def obtener_temperatura_color(attrs):
                return (
                        attrs.get("Temperatura de color")
                        or attrs.get("Tono de luz")
                )

            normalizar_nombres_productos(
                test=TEST,
                default_code_test=TEST_DEFAULT_CODE
            )

        #normalizar_nombres_b2b()#(True, "GT9PAGLAXAL20M1")

        def normalizar_lite_nombres_b2b():
            ctx = {"context": {"lang": "es_ES"}}

            def cat_public(n):
                return models.execute_kw(db, uid, password, "product.public.category", "search", [[["name", "=", n]]],
                                         ctx)

            def sep(n):
                n = n.upper().strip()
                if n.startswith("#[") and "]" in n:
                    i = n.index("]") + 1
                    return n[:i] + " ", n[i:].strip()
                if n.startswith("[") and "]" in n:
                    i = n.index("]") + 1
                    return n[:i] + " ", n[i:].strip()
                if n.startswith("#"):
                    return "#", n[1:].strip()
                return "", n

            def limpio(x):
                return " ".join(x.split()).strip()

            def extraer_primero(tokens, patron):
                for x in tokens:
                    if re.fullmatch(patron, x):
                        return x
                return ""

            def extraer_todos(tokens, patron):
                return [x for x in tokens if re.fullmatch(patron, x)]

            def quitar_tokens(tokens, quitar):
                quitar = set(quitar)
                usados = []
                res = []
                for x in tokens:
                    if x in quitar and x not in usados:
                        usados.append(x)
                        continue
                    res.append(x)
                return res

            def tiras_neones(n):
                p, s = sep(n)
                if "(" in s or ")" in s:
                    return limpio(p + s)
                t = s.split()
                cri = extraer_todos(t, r"CRI[>=]?\d+")
                ip = extraer_todos(t, r"IP\d+")
                color = extraer_todos(t, r"\d{3,5}K|RGB|RGBW|RGB\+W|RGB\+WW|RGB\+CW|RGB\+WW\+CW")
                rollo = []
                for i, x in enumerate(t):
                    if "ROLLO" in x:
                        rollo.append(x)
                    elif re.fullmatch(r"\d+(?:[.,]\d+)?M", x) and (
                            (i > 0 and "ROLLO" in t[i - 1]) or (i + 1 < len(t) and "ROLLO" in t[i + 1])):
                        rollo.append(x)
                quitar = cri + ip + color + rollo
                base = quitar_tokens(t, quitar)
                return limpio(p + " ".join(base + cri + ip + color + rollo))

            def fuentes(n):
                p, s = sep(n)
                s = s.replace("VDC", "V DC").replace("VAC", "V AC")
                if "(" in s or ")" in s:
                    return limpio(p + s)
                t = s.split()
                v = extraer_primero(t, r"\d+(?:[.,]\d+)?V")
                dc = "DC" if "DC" in t else ""
                ac = "AC" if "AC" in t else ""
                w = extraer_primero(t, r"\d+(?:[.,]\d+)?W")
                a = extraer_primero(t, r"\d+(?:[.,]\d+)?A")
                quitar = [x for x in [v, dc, ac, w, a] if x]
                base = quitar_tokens(t, quitar)
                return limpio(p + " ".join(base + [x for x in [v, dc or ac, w, a] if x]))

            def leer_ip(product_id):
                lines = models.execute_kw(db, uid, password, "product.template.attribute.line", "search_read",
                                          [[["product_tmpl_id", "=", product_id],
                                            ["attribute_id.name", "=", "Grado de protección IP"]]],
                                          {"fields": ["value_ids"], **ctx})
                if not lines or not lines[0].get("value_ids"):
                    return ""
                values = models.execute_kw(db, uid, password, "product.attribute.value", "read",
                                           [lines[0]["value_ids"]], {"fields": ["name"], **ctx})
                return values[0]["name"].upper() if values else ""

            def proyectores(product):
                n = product["name"].upper().strip()
                ip = leer_ip(product["id"])
                return limpio(n if not ip or ip in n else n + " " + ip)

            grupos = [
                ("TIRAS Y NEONES LED", tiras_neones),
                ("FUENTES DE ALIMENTACIÓN", fuentes),
                ("PROYECTORES LED", proyectores),
            ]

            total = 0

            for cat_name, funcion in grupos:
                cat_ids = cat_public(cat_name)
                if not cat_ids:
                    print(f"Categoría no encontrada: {cat_name}")
                    continue

                domain = [["public_categ_ids", "child_of", cat_ids], ["x_new_name", "=", False]]

                products = models.execute_kw(db, uid, password, "product.template", "search_read", [domain],
                                             {"fields": ["id", "name", "default_code"], **ctx})

                print(f"{cat_name}: {len(products)} productos encontrados")

                for product in products:
                    nuevo_name = funcion(product) if funcion == proyectores else funcion(product["name"])
                    if nuevo_name and nuevo_name != product["name"]:
                        models.execute_kw(db, uid, password, "product.template", "write",
                                          [[product["id"]], {"name": nuevo_name}], ctx)
                        total += 1
                        print(f"Actualizado {product.get('default_code')}: {nuevo_name}")

            print(f"Total actualizados: {total}")
            return total

        #normalizar_lite_nombres_b2b()

        def norm_atr(excel):
            from services import utils_excel
            utils_excel.normalizar_atributos_excel(excel, "it", excel)

        #norm_atr(ruta)

        def actualizar_proveedores_y_costes_excel(
                excel_path,
                TEST=False,
                sku_test=None,
                proveedor_id=3277,
        ):
            """
            Lee Sheet1 del Excel y actualiza/crea supplierinfo en Odoo.
            También actualiza standard_price del producto.

            Columnas esperadas en Sheet1:
            SKU, NOMBRE, COMPRA, COSTE
            """

            df = pd.read_excel(excel_path, sheet_name="Sheet1", dtype={"SKU": str})

            columnas_obligatorias = {"SKU", "NOMBRE", "COMPRA", "COSTE"}
            faltantes = columnas_obligatorias - set(df.columns)

            if faltantes:
                raise ValueError(f"Faltan columnas obligatorias en el Excel: {faltantes}")

            df["SKU"] = df["SKU"].astype(str).str.strip()

            if TEST:
                if not sku_test:
                    raise ValueError("Si TEST=True debes indicar sku_test")

                df = df[df["SKU"] == str(sku_test).strip()]

                if df.empty:
                    print(f"No se encontró el SKU {sku_test} en el Excel")
                    return

            for _, row in df.iterrows():
                sku = str(row["SKU"]).strip()
                nombre = str(row["NOMBRE"]).strip()

                compra = row["COMPRA"]
                coste = row["COSTE"]

                if not sku or sku.lower() == "nan":
                    continue

                product_ids = models.execute_kw(
                    db,
                    uid,
                    password,
                    "product.product",
                    "search",
                    [[["default_code", "=", sku]]],
                    {"limit": 1},
                )

                if not product_ids:
                    print(f"{sku}")
                    continue

                product_id = product_ids[0]

                template_data = models.execute_kw(
                    db,
                    uid,
                    password,
                    "product.product",
                    "read",
                    [product_id],
                    {"fields": ["product_tmpl_id"]},
                )

                product_tmpl_id = template_data[0]["product_tmpl_id"][0]

                # Actualizar coste del producto
                models.execute_kw(
                    db,
                    uid,
                    password,
                    "product.product",
                    "write",
                    [[product_id], {"standard_price": float(coste)}],
                )

                supplierinfo_ids = models.execute_kw(
                    db,
                    uid,
                    password,
                    "product.supplierinfo",
                    "search",
                    [[
                        ["partner_id", "=", proveedor_id],
                        ["product_tmpl_id", "=", product_tmpl_id],
                    ]],
                    {"limit": 1},
                )

                vals_supplier = {
                    "partner_id": proveedor_id,
                    "product_tmpl_id": product_tmpl_id,
                    "product_name": nombre,
                    "product_code": sku,
                    "min_qty": 1,
                    "price": float(compra),
                }

                if supplierinfo_ids:
                    models.execute_kw(
                        db,
                        uid,
                        password,
                        "product.supplierinfo",
                        "write",
                        [supplierinfo_ids, vals_supplier],
                    )
                    #print(f"Proveedor actualizado: {sku}")
                else:
                    models.execute_kw(
                        db,
                        uid,
                        password,
                        "product.supplierinfo",
                        "create",
                        [vals_supplier],
                    )
                    #print(f"Proveedor creado: {sku}")

                #print(f"Coste actualizado: {sku} -> {coste}")

            print("Proceso Terminado")

        #actualizar_proveedores_y_costes_excel(ruta)#, True, 21200236) #V-TAC MADRID

        def poner_names_mayusculas_b2b():
            CTX = {"context": {"lang": "es_ES"}}

            CATEGORIAS_ORIGEN = [
                "TIRAS Y NEONES LED",
                "FUENTES DE ALIMENTACIÓN",
                "PROYECTORES LED",
                "BOMBILLAS LED",
                "PERFILERIA DE ALUMINIO PARA TIRAS Y NEONES LED",
            ]

            total = 0

            cat_ids = models.execute_kw(
                db, uid, password,
                "product.public.category",
                "search",
                [[["name", "in", CATEGORIAS_ORIGEN]]],
                CTX
            )

            if not cat_ids:
                return 0

            products = models.execute_kw(
                db, uid, password,
                "product.template",
                "search_read",
                [[
                    ["public_categ_ids", "child_of", cat_ids],
                    ["name", "!=", False],
                ]],
                {
                    "fields": ["id", "name", "default_code"],
                    **CTX
                }
            )

            for p in products:
                nuevo = p["name"].upper()
                if nuevo != p["name"]:
                    models.execute_kw(
                        db, uid, password,
                        "product.template",
                        "write",
                        [[p["id"]], {"name": nuevo}],
                        CTX
                    )
                    total += 1
                    print(f"Actualizado {p.get('default_code')}: {nuevo}")

            print(f"Total actualizados: {total}")
            return total

        #poner_names_mayusculas_b2b()

        def actualizar_coste_y_precio_proveedor(excel_path):
            from App_Connection import models, db, uid, password
            import pandas as pd

            ctx = {"lang": "es_ES"}
            partner_id = 3277

            df = pd.read_excel(excel_path, sheet_name=0, dtype={"default_code": str})

            df.columns = [str(c).strip() for c in df.columns]

            required_cols = {"default_code", "standard_price"}
            missing = required_cols - set(df.columns)

            if missing:
                raise ValueError(f"Faltan columnas obligatorias en el Excel: {missing}")

            df["default_code"] = (
                df["default_code"]
                .astype(str)
                .str.strip()
                .str.removesuffix(".0")
            )

            updated_products = 0
            updated_supplierinfos = 0
            not_found = []

            for _, row in df.iterrows():
                sku = str(row["default_code"]).strip().removesuffix(".0")

                if not sku or sku.lower() == "nan":
                    continue

                try:
                    price = float(row["standard_price"])
                except Exception:
                    print(f"⚠ Precio inválido para SKU {sku}: {row['standard_price']}")
                    continue

                product_ids = models.execute_kw(
                    db,
                    uid,
                    password,
                    "product.product",
                    "search",
                    [[["default_code", "=", sku]]],
                    {"context": ctx, "limit": 1},
                )

                if not product_ids:
                    print(f"❌ Producto no encontrado: {sku}")
                    not_found.append(sku)
                    continue

                product_id = product_ids[0]

                product_data = models.execute_kw(
                    db,
                    uid,
                    password,
                    "product.product",
                    "read",
                    [product_id],
                    {"context": ctx, "fields": ["product_tmpl_id"]},
                )

                product_tmpl_id = product_data[0]["product_tmpl_id"][0]

                # 1) Actualizar standard_price en product.template
                models.execute_kw(
                    db,
                    uid,
                    password,
                    "product.template",
                    "write",
                    [[product_tmpl_id], {"standard_price": price}],
                    {"context": ctx},
                )

                updated_products += 1

                # 2) Buscar línea de compra existente del proveedor 3277
                supplierinfo_ids = models.execute_kw(
                    db,
                    uid,
                    password,
                    "product.supplierinfo",
                    "search",
                    [[
                        ["product_tmpl_id", "=", product_tmpl_id],
                        ["partner_id", "=", partner_id],
                    ]],
                    {"context": ctx},
                )

                if supplierinfo_ids:
                    models.execute_kw(
                        db,
                        uid,
                        password,
                        "product.supplierinfo",
                        "write",
                        [supplierinfo_ids, {"price": price}],
                        {"context": ctx},
                    )

                    updated_supplierinfos += len(supplierinfo_ids)
                    print(f"✅ {sku} | standard_price={price} | supplierinfo actualizado")
                else:
                    print(f"⚠ {sku} | standard_price={price} | sin supplierinfo partner_id={partner_id}")

            print("\n===== RESUMEN =====")
            print(f"Productos actualizados: {updated_products}")
            print(f"Líneas proveedor actualizadas: {updated_supplierinfos}")
            print(f"No encontrados: {len(not_found)}")

            return {
                "updated_products": updated_products,
                "updated_supplierinfos": updated_supplierinfos,
                "not_found": not_found,
            }

        #actualizar_coste_y_precio_proveedor(ruta)

        # ---------------------------------------------------------------------------
        # 2) NORMALIZAR ATRIBUTOS PRODUCTOS B2B
        # ---------------------------------------------------------------------------

        def migrar_valores_atributos(
                attr_origen_name,
                attr_destino_name,
                dry_run=True,
                lang="es_ES",
        ):
            """
            Migra valores que contienen kWh desde un atributo origen a un atributo destino.

            - Busca valores con kWh en attr_origen_name.
            - Añade esos valores al atributo destino.
            - Crea la línea del atributo destino en el producto si no existe.
            - Elimina esos valores de la línea origen.
            - Si la línea origen queda vacía, la elimina.
            - Al final borra del atributo origen los valores kWh que ya no estén usados.
            """

            ctx = {"lang": lang}

            def odoo_search(model, domain, limit=None):
                kwargs = {"context": ctx}
                if limit:
                    kwargs["limit"] = limit

                return models.execute_kw(
                    db, uid, password,
                    model, "search",
                    [domain],
                    kwargs,
                )

            def odoo_read(model, ids, fields):
                if not ids:
                    return []

                return models.execute_kw(
                    db, uid, password,
                    model, "read",
                    [ids],
                    {"fields": fields, "context": ctx},
                )

            def odoo_create(model, vals):
                return models.execute_kw(
                    db, uid, password,
                    model, "create",
                    [vals],
                    {"context": ctx},
                )

            def odoo_write(model, ids, vals):
                return models.execute_kw(
                    db, uid, password,
                    model, "write",
                    [ids, vals],
                    {"context": ctx},
                )

            def odoo_unlink(model, ids):
                return models.execute_kw(
                    db, uid, password,
                    model, "unlink",
                    [ids],
                    {"context": ctx},
                )

            def contiene_kwh(value_name):
                return "kwh" in (value_name or "").lower()

            print("Buscando atributos...")

            origen_attr_ids = odoo_search(
                "product.attribute",
                [("name", "=", attr_origen_name)],
                limit=1,
            )

            if not origen_attr_ids:
                raise Exception(f'No existe el atributo origen: "{attr_origen_name}"')

            origen_attr_id = origen_attr_ids[0]

            destino_attr_ids = odoo_search(
                "product.attribute",
                [("name", "=", attr_destino_name)],
                limit=1,
            )

            if destino_attr_ids:
                destino_attr_id = destino_attr_ids[0]
            else:
                print(f'No existe el atributo destino "{attr_destino_name}". Se creará.')

                if dry_run:
                    destino_attr_id = None
                    print(f'[DRY RUN] Crearía atributo destino: "{attr_destino_name}"')
                else:
                    destino_attr_id = odoo_create(
                        "product.attribute",
                        {
                            "name": attr_destino_name,
                            "display_type": "radio",
                            "create_variant": "no_variant",
                        },
                    )

            print(f'ID atributo origen "{attr_origen_name}": {origen_attr_id}')
            print(f'ID atributo destino "{attr_destino_name}": {destino_attr_id}')

            origen_value_ids = odoo_search(
                "product.attribute.value",
                [
                    ("attribute_id", "=", origen_attr_id),
                    ("name", "ilike", "kWh"),
                ],
            )

            origen_values = odoo_read(
                "product.attribute.value",
                origen_value_ids,
                ["name", "attribute_id"],
            )

            origen_kwh_value_ids = [
                value["id"]
                for value in origen_values
                if contiene_kwh(value["name"])
            ]

            if not origen_kwh_value_ids:
                print(f'No se han encontrado valores con kWh dentro del atributo "{attr_origen_name}".')
                return

            print(f'Valores kWh encontrados en "{attr_origen_name}": {len(origen_kwh_value_ids)}')

            origen_line_ids = odoo_search(
                "product.template.attribute.line",
                [
                    ("attribute_id", "=", origen_attr_id),
                    ("value_ids", "in", origen_kwh_value_ids),
                ],
            )

            lines = odoo_read(
                "product.template.attribute.line",
                origen_line_ids,
                ["product_tmpl_id", "attribute_id", "value_ids"],
            )

            print(f"Productos/líneas afectadas: {len(lines)}")

            created_or_found_destino_values = {}
            productos_modificados = 0
            valores_movidos = 0

            for line in lines:
                line_id = line["id"]
                product_tmpl_id = line["product_tmpl_id"][0]
                product_name = line["product_tmpl_id"][1]
                current_value_ids = line["value_ids"]

                kwh_ids_in_line = [
                    value_id
                    for value_id in current_value_ids
                    if value_id in origen_kwh_value_ids
                ]

                if not kwh_ids_in_line:
                    continue

                kwh_values = odoo_read(
                    "product.attribute.value",
                    kwh_ids_in_line,
                    ["name"],
                )

                print("")
                print(f'Producto: "{product_name}"')
                print(f'Línea origen "{attr_origen_name}" ID: {line_id}')

                destino_value_ids_to_add = []

                for value in kwh_values:
                    value_name = value["name"].strip()

                    print(f'  Mover valor: "{value_name}"')

                    if value_name in created_or_found_destino_values:
                        destino_value_id = created_or_found_destino_values[value_name]
                    else:
                        existing_destino_value_ids = odoo_search(
                            "product.attribute.value",
                            [
                                ("attribute_id", "=", destino_attr_id),
                                ("name", "=", value_name),
                            ],
                            limit=1,
                        ) if destino_attr_id else []

                        if existing_destino_value_ids:
                            destino_value_id = existing_destino_value_ids[0]
                        else:
                            if dry_run:
                                destino_value_id = None
                                print(
                                    f'  [DRY RUN] Crearía valor "{value_name}" '
                                    f'en atributo destino "{attr_destino_name}"'
                                )
                            else:
                                destino_value_id = odoo_create(
                                    "product.attribute.value",
                                    {
                                        "name": value_name,
                                        "attribute_id": destino_attr_id,
                                    },
                                )

                        created_or_found_destino_values[value_name] = destino_value_id

                    if destino_value_id:
                        destino_value_ids_to_add.append(destino_value_id)

                destino_line_ids = odoo_search(
                    "product.template.attribute.line",
                    [
                        ("product_tmpl_id", "=", product_tmpl_id),
                        ("attribute_id", "=", destino_attr_id),
                    ],
                    limit=1,
                ) if destino_attr_id else []

                if destino_line_ids:
                    destino_line_id = destino_line_ids[0]

                    print(f'Línea destino "{attr_destino_name}" existente ID: {destino_line_id}')

                    if dry_run:
                        print(
                            f"  [DRY RUN] Añadiría valores a destino: "
                            f"{destino_value_ids_to_add}"
                        )
                    else:
                        odoo_write(
                            "product.template.attribute.line",
                            [destino_line_id],
                            {
                                "value_ids": [
                                    (4, value_id)
                                    for value_id in destino_value_ids_to_add
                                ]
                            },
                        )
                else:
                    print(f'No existe línea destino "{attr_destino_name}". Se creará.')

                    if dry_run:
                        print(
                            f"  [DRY RUN] Crearía línea destino con valores: "
                            f"{destino_value_ids_to_add}"
                        )
                    else:
                        odoo_create(
                            "product.template.attribute.line",
                            {
                                "product_tmpl_id": product_tmpl_id,
                                "attribute_id": destino_attr_id,
                                "value_ids": [(6, 0, destino_value_ids_to_add)],
                            },
                        )

                remaining_origen_value_ids = [
                    value_id
                    for value_id in current_value_ids
                    if value_id not in kwh_ids_in_line
                ]

                if remaining_origen_value_ids:
                    print(
                        f'  "{attr_origen_name}" conserva valores restantes: '
                        f"{remaining_origen_value_ids}"
                    )

                    if dry_run:
                        print(
                            f'  [DRY RUN] Eliminaría de "{attr_origen_name}" '
                            f"los valores: {kwh_ids_in_line}"
                        )
                    else:
                        odoo_write(
                            "product.template.attribute.line",
                            [line_id],
                            {
                                "value_ids": [
                                    (3, value_id)
                                    for value_id in kwh_ids_in_line
                                ]
                            },
                        )
                else:
                    print(f'  La línea origen "{attr_origen_name}" quedaría vacía. Se eliminará.')

                    if dry_run:
                        print(f"  [DRY RUN] Eliminaría línea origen ID: {line_id}")
                    else:
                        odoo_unlink(
                            "product.template.attribute.line",
                            [line_id],
                        )

                productos_modificados += 1
                valores_movidos += len(kwh_ids_in_line)

            print("")
            print(f'Limpieza final de valores kWh en atributo origen "{attr_origen_name}"...')

            valores_eliminados = 0

            for value_id in origen_kwh_value_ids:
                value_data = odoo_read(
                    "product.attribute.value",
                    [value_id],
                    ["name"],
                )

                if not value_data:
                    continue

                value_name = value_data[0]["name"]

                usage_line_ids = odoo_search(
                    "product.template.attribute.line",
                    [
                        ("attribute_id", "=", origen_attr_id),
                        ("value_ids", "in", [value_id]),
                    ],
                    limit=1,
                )

                if usage_line_ids:
                    print(
                        f'  No se elimina "{value_name}" porque todavía está usado '
                        f'en "{attr_origen_name}".'
                    )
                    continue

                print(f'  Eliminar valor de "{attr_origen_name}": "{value_name}"')

                if dry_run:
                    print(f"  [DRY RUN] Eliminaría product.attribute.value ID {value_id}")
                else:
                    odoo_unlink(
                        "product.attribute.value",
                        [value_id],
                    )

                valores_eliminados += 1

            print("")
            print("Resumen:")
            print(f'  Atributo origen: "{attr_origen_name}"')
            print(f'  Atributo destino: "{attr_destino_name}"')
            print(f"  Productos modificados: {productos_modificados}")
            print(f"  Valores movidos: {valores_movidos}")
            print(f"  Valores eliminados del origen: {valores_eliminados}")
            print(f"  Dry run: {dry_run}")

        #migrar_valores_atributos("Potencia (W)", "Capacidad de la batería (kWh)", False)

        def actualizar_valores_atributos(ruta_excel, nombre_atributo, dry_run=True):
            df = pd.read_excel(ruta_excel)

            atributo_ids = models.execute_kw(
                db, uid, password,
                "product.attribute", "search",
                [[("name", "=", nombre_atributo)]],
                {"limit": 1}
            )

            if not atributo_ids:
                print(f"Error: No se encontró el atributo '{nombre_atributo}'")
                return

            atributo_id = atributo_ids[0]

            for index, fila in df.iterrows():
                valor_actual = str(fila["VALOR"]).strip()
                valor_nuevo = str(fila["CAMBIO"]).strip()

                if not valor_actual or not valor_nuevo or valor_actual == "nan" or valor_nuevo == "nan":
                    print(f"Fila {index + 2}: saltada por VALOR/CAMBIO vacío")
                    continue

                valor_ids = models.execute_kw(
                    db, uid, password,
                    "product.attribute.value", "search",
                    [[
                        ("name", "=", valor_actual),
                        ("attribute_id", "=", atributo_id),
                    ]],
                    {"limit": 1}
                )

                if valor_ids:
                    print(f"Actualizando: '{valor_actual}' -> '{valor_nuevo}'")

                    if not dry_run:
                        models.execute_kw(
                            db, uid, password,
                            "product.attribute.value", "write",
                            [valor_ids, {"name": valor_nuevo}]
                        )
                else:
                    print(
                        f"Advertencia: No se encontró el valor '{valor_actual}' "
                        f"para el atributo '{nombre_atributo}'"
                    )

        #actualizar_valores_atributos(ruta, "Potencia", False)

        def borrar_valores_no_usados_atributo(
                atr_name="Potencia (W)",
                dry_run=True,
                lang="es_ES",
        ):
            ctx = {"lang": lang}

            attr_ids = models.execute_kw(
                db, uid, password,
                "product.attribute", "search",
                [[("name", "=", atr_name)]],
                {"limit": 1, "context": ctx},
            )

            if not attr_ids:
                print(f'No existe el atributo "{atr_name}"')
                return

            attr_id = attr_ids[0]

            value_ids = models.execute_kw(
                db, uid, password,
                "product.attribute.value", "search",
                [[("attribute_id", "=", attr_id)]],
                {"context": ctx},
            )

            if not value_ids:
                print(f'No hay valores en "{atr_name}"')
                return

            values = models.execute_kw(
                db, uid, password,
                "product.attribute.value", "read",
                [value_ids],
                {"fields": ["name"], "context": ctx},
            )

            usados_line_ids = models.execute_kw(
                db, uid, password,
                "product.template.attribute.line", "search",
                [[
                    ("attribute_id", "=", attr_id),
                    ("value_ids", "in", value_ids),
                ]],
                {"context": ctx},
            )

            usados_value_ids = set()

            if usados_line_ids:
                lines = models.execute_kw(
                    db, uid, password,
                    "product.template.attribute.line", "read",
                    [usados_line_ids],
                    {"fields": ["value_ids"], "context": ctx},
                )

                for line in lines:
                    usados_value_ids.update(line.get("value_ids", []))

            no_usados = [
                value
                for value in values
                if value["id"] not in usados_value_ids
            ]

            print(f'Valores totales en "{atr_name}": {len(values)}')
            print(f"Valores usados: {len(usados_value_ids)}")
            print(f"Valores no usados a borrar: {len(no_usados)}")

            for value in no_usados:
                print(f'Borrar: ID {value["id"]} - "{value["name"]}"')

            if dry_run:
                print("[DRY RUN] No se ha borrado nada.")
                return

            if no_usados:
                models.execute_kw(
                    db, uid, password,
                    "product.attribute.value", "unlink",
                    [[value["id"] for value in no_usados]],
                    {"context": ctx},
                )

            print("Borrado terminado.")

        #endregion


        ruta = Utils.seleccionar_excel()

        if not ruta:
            msg = "❌ Import cancelado - No se seleccionó archivo."
            emit_progress(msg)
            print(msg)
            if event_manager:
                event_manager.emit('operation_error', "No se seleccionó archivo para importar")
            return

        def run_import():
            try:
                emit_progress("📋 Preparando archivo Excel para importación...")
                Utils.preparar_excel_import(ruta)
                df_import = pd.read_excel(ruta, sheet_name='Import'); df_update = pd.read_excel(ruta, sheet_name='Update'); df_quantity = pd.read_excel(ruta, sheet_name='Quantity')
                import_all = True if df_import.empty and df_update.empty and df_quantity.empty else False

                emit_progress("🚀 Iniciando importación de productos...")
                import_to_odoo(ruta, import_all=import_all, event_manager=event_manager)

            except Exception as e:
                msg = f"❌ Error en import_to_odoo: {e}"
                emit_progress(msg)
                print(msg)
                if event_manager:
                    event_manager.emit('operation_error', f"Error en importación: {e}")
                return

            # Solo continuar si no hubo errores
            run_update()

        def run_update():
            try:
                emit_progress("🔄 Iniciando actualización de productos...")
                update_from_merge_to_odoo(ruta, event_manager)
            except Exception as e:
                msg = f"❌ Error en update_from_merge_to_odoo: {e}"
                emit_progress(msg)
                print(msg)
                if event_manager:
                    event_manager.emit('operation_error', f"Error en actualización: {e}")
                return

            # Solo continuar si no hubo errores
            run_delete()

        def run_delete():
            try:
                emit_progress("🗑️ Iniciando eliminación de productos obsoletos...")
                delete_from_merge_to_odoo(ruta, event_manager)
                emit_progress("✅ Proceso de importación completado exitosamente")
                if event_manager:
                    event_manager.emit('operation_completed', "Importación finalizada correctamente")
            except Exception as e:
                msg = f"❌ Error en delete_from_merge_to_odoo: {e}"
                emit_progress(msg)
                print(msg)
                if event_manager:
                    event_manager.emit('operation_error', f"Error en eliminación: {e}")

        # Ejecutar en el hilo principal para que sea síncrono
        run_import()
    else:
        msg = "❌ Error en la conexión, no es posible hacer el Import."
        emit_progress(msg)
        print(msg)
        if event_manager:
            event_manager.emit('operation_error', "Error en la conexión con Odoo")
