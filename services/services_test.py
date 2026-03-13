import pandas as pd
from openpyxl import load_workbook
from utils import Utils

ruta = Utils.seleccionar_excel()

def detectar_cambios_excel(ruta_excel):
    """
    Compara las hojas Odoo16 y Odoo18 y crea la hoja 'Cambios' con:
      - default_code
      - Solo las columnas que difieren (valores tomados de Odoo16)
        Si el valor fue eliminado (vacío en 16 pero no en 18) se marca con '*'.
    Las demás columnas quedan vacías.
    """

    # Leer las hojas
    df16 = pd.read_excel(ruta_excel, sheet_name='Odoo16')
    df18 = pd.read_excel(ruta_excel, sheet_name='Odoo18')

    if 'default_code' not in df16.columns or 'default_code' not in df18.columns:
        raise ValueError("Falta la columna 'default_code' en una o ambas hojas.")

    # Limpiar duplicados y vacíos
    df16 = df16.dropna(subset=['default_code']).drop_duplicates(subset=['default_code'])
    df18 = df18.dropna(subset=['default_code']).drop_duplicates(subset=['default_code'])

    # Indexar por default_code
    df16 = df16.set_index('default_code')
    df18 = df18.set_index('default_code')

    comunes = df16.index.intersection(df18.index)
    columnas_comunes = [col for col in df16.columns if col in df18.columns]

    cambios = []

    for code in comunes:
        fila16 = df16.loc[code]
        fila18 = df18.loc[code]

        dif_cols = [
            col for col in columnas_comunes
            if str(fila16[col]).strip() != str(fila18[col]).strip()
        ]

        if dif_cols:
            fila_resultado = {col: "" for col in columnas_comunes}  # vacío por defecto
            fila_resultado['default_code'] = code

            for col in dif_cols:
                val16 = str(fila16[col]).strip()
                val18 = str(fila18[col]).strip()

                # Si el valor fue eliminado (antes existía y ahora está vacío)
                if val16 in ["", "nan", "None"] and val18 not in ["", "nan", "None"]:
                    fila_resultado[col] = "*"  # marca eliminación
                else:
                    fila_resultado[col] = fila16[col]  # valor normal de Odoo16

            cambios.append(fila_resultado)

    if not cambios:
        print("✅ No se detectaron diferencias entre Odoo16 y Odoo18.")
        return

    # Crear DataFrame con las mismas columnas
    df_cambios = pd.DataFrame(cambios)
    columnas_finales = ['default_code'] + [c for c in columnas_comunes if c != 'default_code']
    df_cambios = df_cambios[columnas_finales]

    # Guardar resultados
    with pd.ExcelWriter(ruta_excel, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df_cambios.to_excel(writer, sheet_name='Cambios', index=False)

    print(f"💾 {len(df_cambios)} productos con diferencias guardados en 'Cambios'.")

def rellenar_stock_multialmacen(ruta_excel):
    import pandas as pd
    from openpyxl import load_workbook

    try:
        # Leer todas las hojas del Excel
        xls = pd.ExcelFile(ruta_excel)
        df_main = pd.read_excel(xls, sheet_name="Sheet1")
        df_madrid = pd.read_excel(xls, sheet_name="Madrid")
        df_bulgaria = pd.read_excel(xls, sheet_name="Bulgaria")

        # Asegurar columnas de destino en Sheet1
        for col in ["M_STOCK", "B_STOCK", "UNDELIVERED ORDER"]:
            if col not in df_main.columns:
                df_main[col] = ""

        print("🔍 Rellenando datos desde las hojas 'Madrid' y 'Bulgaria'...")

        # --- 1️⃣ Copiar STOCK de Madrid a M_STOCK ---
        dict_madrid = dict(zip(df_madrid["SKU"], df_madrid["STOCK"]))
        df_main["M_STOCK"] = df_main["SKU"].map(dict_madrid).fillna(df_main["M_STOCK"])

        # --- 2️⃣ Copiar STOCK de Bulgaria a B_STOCK ---
        dict_bulgaria_stock = dict(zip(df_bulgaria["SKU"], df_bulgaria["STOCK"]))
        df_main["B_STOCK"] = df_main["SKU"].map(dict_bulgaria_stock).fillna(df_main["B_STOCK"])

        # --- 3️⃣ Si STOCK vacío o < 0 y UNDELIVERED ORDER tiene valor → copiarlo ---
        if "UNDELIVERED ORDER" in df_bulgaria.columns:
            dict_bulgaria_order = dict(zip(df_bulgaria["SKU"], df_bulgaria["UNDELIVERED ORDER"]))
            for i, row in df_main.iterrows():
                sku = row["SKU"]
                if pd.isna(sku):
                    continue

                stock_bulg = dict_bulgaria_stock.get(sku, None)
                undelivered_val = dict_bulgaria_order.get(sku, None)

                if (pd.isna(stock_bulg) or stock_bulg < 0) and pd.notna(undelivered_val):
                    df_main.at[i, "UNDELIVERED ORDER"] = undelivered_val

        # --- Guardar cambios ---
        with pd.ExcelWriter(ruta_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_main.to_excel(writer, sheet_name="Sheet1", index=False)

        print("✅ Datos actualizados correctamente en 'Sheet1'.")

    except Exception as e:
        print(f"❌ Error en el proceso: {e}")


import xmlrpc.client

def listar_informes(url, db, username, password, modelo="sale.order"):
    """
    Lista todos los informes PDF (ir.actions.report) registrados en Odoo
    para un modelo dado (por defecto 'sale.order').
    """
    print(f"🔌 Conectando a {url}...")
    common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(db, username, password, {})
    if not uid:
        print("❌ Autenticación fallida.")
        return

    models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True)
    print(f"✅ Autenticado con UID {uid}")

    # Buscar informes asociados al modelo indicado
    report_ids = models.execute_kw(
        db, uid, password,
        "ir.actions.report", "search_read",
        [[("model", "=", modelo), ("report_type", "=", "qweb-pdf")]],
        {"fields": ["id", "name", "report_name", "binding_model_id", "report_file"]}
    )

    if not report_ids:
        print(f"⚠️ No se encontraron informes PDF para el modelo '{modelo}'.")
        return

    print(f"📄 Informes encontrados ({len(report_ids)}):\n")
    for r in report_ids:
        print(f"  ID: {r['id']}")
        print(f"  Nombre visible: {r['name']}")
        print(f"  report_name: {r['report_name']}")
        print(f"  report_file: {r['report_file']}")
        print(f"  Modelo: {r.get('binding_model_id', [''])[1] if r.get('binding_model_id') else '—'}")
        print("-" * 60)


import xmlrpc.client
import math
import time


def migrar_tarea(origen, destino, origen_task_id, destino_task_id, lote=30, delay=0.2):
    """
    Migra una tarea de project.task desde Odoo origen → destino.

    origen: dict con url, db, user, password
    destino: dict con url, db, user, password
    origen_task_id : ID de tarea en origen
    destino_task_id : ID de tarea en destino
    lote : cantidad de adjuntos por lote (default 30)
    delay : segundos entre lotes para evitar saturación (default 0.2)
    """

    print(f"\n=== MIGRANDO TAREA {origen_task_id} → {destino_task_id} ===")

    # --- LOGIN ORIGEN ---
    common = xmlrpc.client.ServerProxy(f'{origen['url']}/xmlrpc/2/common', allow_none=True)
    uid = common.authenticate(origen['db'], origen['user'], origen['password'], {})

    if not uid:
        print("❌ No se pudo autenticar.")
        return

    print(f'🔌 Conectado como {origen['user']} (uid: {uid})')
    models = xmlrpc.client.ServerProxy(f'{origen['url']}/xmlrpc/2/object', allow_none=True)

    # --- LOGIN DESTINO ---
    dest_common = xmlrpc.client.ServerProxy(f'{destino['url']}/xmlrpc/2/common', allow_none=True)
    dest_uid = dest_common.authenticate(destino['db'], destino['user'], destino['password'], {})

    if not dest_uid:
        print("❌ No se pudo autenticar.")
        return

    print(f'🔌 Conectado como {destino['user']} (uid: {dest_uid})')
    dest_models = xmlrpc.client.ServerProxy(f'{destino['url']}/xmlrpc/2/object', allow_none=True)

    # -----------------------------
    # 3) MIGRAR ADJUNTOS
    # -----------------------------
    print("\nLeyendo adjuntos del origen...")

    attach_ids = models.execute_kw(
        origen['db'], uid, origen['password'],
        'ir.attachment', 'search',
        [[['res_model', '=', 'project.task'], ['res_id', '=', origen_task_id]]]
    )

    total_adjuntos = len(attach_ids)
    print(f"Total adjuntos encontrados: {total_adjuntos}")

    if total_adjuntos == 0:
        print("No hay adjuntos que migrar.")
        return

    # Calculamos lotes
    num_lotes = math.ceil(total_adjuntos / lote)

    for i in range(num_lotes):
        inicio = i * lote
        fin = inicio + lote
        lote_ids = attach_ids[inicio:fin]

        print(f"\nProcesando lote {i + 1}/{num_lotes} ({len(lote_ids)} adjuntos)...")

        '''adjuntos = models.execute_kw(
            origen['db'], uid, origen['password'],
            'ir.attachment', 'read',
            [lote_ids, ['name', 'file_size', 'mimetype']]
        )

        for att in adjuntos:
            print(f"Adjunto: {att['name']} pesa: {att['file_size']}")'''

        adjuntos = models.execute_kw(
            origen['db'], uid, origen['password'],
            'ir.attachment', 'read',
            [lote_ids, ['name', 'datas', 'mimetype']]
        )

        # Crear adjuntos en destino
        for att in adjuntos:
            try:
                dest_models.execute_kw(
                    destino['db'], dest_uid, destino['password'],
                    'ir.attachment', 'create',
                    [{
                        'name': att['name'],
                        'datas': att['datas'],
                        'mimetype': att['mimetype'],
                        'res_model': 'project.task',
                        'res_id': destino_task_id,
                    }]
                )
            except Exception as e:
                print(f"⚠️ Error importando {att['name']}: {e}")

        print(f"✓ Lote {i + 1}/{num_lotes} completado.")

        time.sleep(delay)

    print("\n=== MIGRACIÓN COMPLETADA CORRECTAMENTE ===")

origen = {
    'url': "https://optimaluz.soluntec.net",
    'db': "Test",#"Real",
    'user': "jcoronado@optimaluz.com",
    'password': "AlAi4ever"
}

destino = {
    'url': "http://79.72.61.76:8070/",
    'db': "odoo1",
    'user': "admin",
    'password': "admin"
}

'''migrar_tarea(
    origen,
    destino,
    origen_task_id=1426,
    destino_task_id=1523,
    lote=1,      # tamaño del lote
    delay=0.1     # descanso entre lotes
)'''


def rellenar_ref3_excel(
    ruta_excel,
    hoja="Sheet1",
    col_ref1="Ref1",
    col_ref2="Ref2",
    col_ref3="Ref3"
):
    # Leer solo la hoja necesaria
    df = pd.read_excel(ruta_excel, sheet_name=hoja)

    # Normalizar valores (string limpio)
    ref2_set = set(
        df[col_ref2]
        .dropna()
        .astype(str)
        .str.strip()
    )

    def calcular_ref3(valor):
        if pd.isna(valor):
            return None
        valor = str(valor).strip()
        return valor if valor not in ref2_set else None

    # Aplicar lógica
    df[col_ref3] = df[col_ref1].apply(calcular_ref3)

    # Guardar SOLO la hoja indicada sin borrar el resto del Excel
    with pd.ExcelWriter(
        ruta_excel,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        df.to_excel(writer, sheet_name=hoja, index=False)

    print("✅ Ref3 rellenado correctamente sin borrar el resto del Excel")

def detectar_cambios_excel(ruta_excel):
    """
    Compara las hojas Odoo16 y Odoo18 y crea la hoja 'Cambios' con:
      - default_code
      - Solo las columnas que difieren (valores tomados de Odoo16)
        Si el valor fue eliminado (vacío en 16 pero no en 18) se marca con '*'.
    Las demás columnas quedan vacías.
    """

    # Leer las hojas
    df16 = pd.read_excel(ruta_excel, sheet_name='ODOO16')
    df18 = pd.read_excel(ruta_excel, sheet_name='ODOO18')

    if 'default_code' not in df16.columns or 'default_code' not in df18.columns:
        raise ValueError("Falta la columna 'default_code' en una o ambas hojas.")

    # Limpiar duplicados y vacíos
    df16 = df16.dropna(subset=['default_code']).drop_duplicates(subset=['default_code'])
    df18 = df18.dropna(subset=['default_code']).drop_duplicates(subset=['default_code'])

    # Indexar por default_code
    df16 = df16.set_index('default_code')
    df18 = df18.set_index('default_code')

    comunes = df16.index.intersection(df18.index)
    columnas_comunes = [col for col in df16.columns if col in df18.columns]

    cambios = []

    for code in comunes:
        fila16 = df16.loc[code]
        fila18 = df18.loc[code]

        dif_cols = [
            col for col in columnas_comunes
            if str(fila16[col]).strip() != str(fila18[col]).strip()
        ]

        if dif_cols:
            fila_resultado = {col: "" for col in columnas_comunes}  # vacío por defecto
            fila_resultado['default_code'] = code

            for col in dif_cols:
                val16 = str(fila16[col]).strip()
                val18 = str(fila18[col]).strip()

                # Si el valor fue eliminado (antes existía y ahora está vacío)
                if val16 in ["", "nan", "None"] and val18 not in ["", "nan", "None"]:
                    fila_resultado[col] = "*"  # marca eliminación
                else:
                    fila_resultado[col] = fila16[col]  # valor normal de Odoo16

            cambios.append(fila_resultado)

    if not cambios:
        print("✅ No se detectaron diferencias entre Odoo16 y Odoo18.")
        return

    # Crear DataFrame con las mismas columnas
    df_cambios = pd.DataFrame(cambios)
    columnas_finales = ['default_code'] + [c for c in columnas_comunes if c != 'default_code']
    df_cambios = df_cambios[columnas_finales]

    # Guardar resultados
    with pd.ExcelWriter(ruta_excel, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df_cambios.to_excel(writer, sheet_name='CAMBIOS', index=False)

    print(f"💾 {len(df_cambios)} productos con diferencias guardados en 'CAMBIOS'.")

    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill

    def resaltar_cambios_no_default_code(
            ruta_excel,
            hoja="CAMBIOS",
            columna_default_code="default_code"
    ):
        wb = load_workbook(ruta_excel)
        ws = wb[hoja]

        # Estilo amarillo
        amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # ---------------------------------------
        # 1️⃣ Detectar índice de columna default_code
        # ---------------------------------------
        headers = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}

        if columna_default_code not in headers:
            raise ValueError(f"No existe la columna '{columna_default_code}'")

        col_default = headers[columna_default_code]

        # ---------------------------------------
        # 2️⃣ Obtener todos los valores default_code
        # ---------------------------------------
        default_codes = set()
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=col_default).value
            if val not in (None, ""):
                default_codes.add(str(val).strip())

        # ---------------------------------------
        # 3️⃣ Recorrer resto de columnas
        # ---------------------------------------
        for col_name, col_idx in headers.items():
            if col_name == columna_default_code:
                continue

            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                val = cell.value

                if val in (None, ""):
                    continue

                if str(val).strip() not in default_codes:
                    cell.fill = amarillo

        # ---------------------------------------
        # 4️⃣ Guardar cambios
        # ---------------------------------------
        wb.save(ruta_excel)

        print("✅ Valores resaltados correctamente en amarillo")

    resaltar_cambios_no_default_code(ruta_excel)

from openpyxl import load_workbook

def copiar_datos_por_sku_excel(
    excel_path,
    sheet_dest="Sheet1",
    sheet_src="Sheet2",
    col_sku="SKU"
):
    """
    Copia las columnas de Sheet2 a Sheet1 para SKUs coincidentes.
    Abre el Excel en modo apéndice (sin borrar nada).
    """

    wb = load_workbook(excel_path)
    ws_dest = wb[sheet_dest]
    ws_src = wb[sheet_src]

    # --- Leer cabeceras ---
    headers_dest = {cell.value: idx + 1 for idx, cell in enumerate(ws_dest[1])}
    headers_src = {cell.value: idx + 1 for idx, cell in enumerate(ws_src[1])}

    if col_sku not in headers_dest or col_sku not in headers_src:
        raise ValueError("❌ La columna SKU no existe en alguna de las hojas")

    # --- Crear columnas faltantes en destino ---
    for col_name in headers_src:
        if col_name == col_sku:
            continue
        if col_name not in headers_dest:
            ws_dest.cell(row=1, column=ws_dest.max_column + 1, value=col_name)
            headers_dest[col_name] = ws_dest.max_column

    # --- Mapa SKU → fila destino ---
    sku_dest_map = {}
    for row in range(2, ws_dest.max_row + 1):
        sku = ws_dest.cell(row=row, column=headers_dest[col_sku]).value
        if sku:
            sku_dest_map[str(sku).strip()] = row

    # --- Copiar datos ---
    copiados = 0

    for row in range(2, ws_src.max_row + 1):
        sku = ws_src.cell(row=row, column=headers_src[col_sku]).value
        if not sku:
            continue

        sku = str(sku).strip()
        if sku not in sku_dest_map:
            continue

        dest_row = sku_dest_map[sku]

        for col_name, src_col_idx in headers_src.items():
            if col_name == col_sku:
                continue

            dest_col_idx = headers_dest[col_name]
            ws_dest.cell(
                row=dest_row,
                column=dest_col_idx,
                value=ws_src.cell(row=row, column=src_col_idx).value
            )

        copiados += 1

    wb.save(excel_path)

    print(f"✅ SKUs actualizados en Sheet1: {copiados}")

def copiar_payable_entre_hojas(
    excel_entrada,
    excel_salida=None,
    sheet_origen="Sheet1",
    sheet_destino="Sheet2"
):
    """
    Copia property_account_payable_id de Sheet1 a Sheet2
    haciendo match por la columna 'name'.
    """

    # 📥 Leer Excel
    xls = pd.ExcelFile(excel_entrada)

    df_origen = pd.read_excel(xls, sheet_name=sheet_origen)
    df_destino = pd.read_excel(xls, sheet_name=sheet_destino)

    # 🧪 Validaciones
    for col in ["name", "property_account_payable_id"]:
        if col not in df_origen.columns:
            raise ValueError(f"❌ Falta columna '{col}' en {sheet_origen}")

    if "name" not in df_destino.columns:
        raise ValueError(f"❌ Falta columna 'name' en {sheet_destino}")

    # ➕ Crear columna en destino si no existe
    if "property_account_payable_id" not in df_destino.columns:
        df_destino["property_account_payable_id"] = None

    # 🧠 Diccionario name → payable (origen)
    mapa_payable = (
        df_origen
        .dropna(subset=["name"])
        .drop_duplicates(subset=["name"])
        .set_index("name")["property_account_payable_id"]
        .to_dict()
    )

    # 🔁 Copiar valores en destino
    coincidencias = 0

    for idx, row in df_destino.iterrows():
        name = row["name"]
        if name in mapa_payable:
            df_destino.at[idx, "property_account_payable_id"] = mapa_payable[name]
            coincidencias += 1

    print(f"✔ Coincidencias actualizadas: {coincidencias}")

    # 💾 Guardar resultado
    salida = excel_salida or excel_entrada

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df_origen.to_excel(writer, sheet_name=sheet_origen, index=False)
        df_destino.to_excel(writer, sheet_name=sheet_destino, index=False)

    print(f"📄 Archivo generado: {salida}")


def copiar_filas_por_sku(
    excel_path,
    hoja_origen="ES",
    hoja_destino="NEW",
    columna_sku="SKU",
    sobrescribir_destino=True
):
    """
    Copia filas completas de la hoja ORIGEN a DESTINO
    cuando el SKU coincide en ambas hojas.

    excel_path : ruta del Excel
    hoja_origen : nombre hoja origen (ES)
    hoja_destino : nombre hoja destino (NEW)
    columna_sku : nombre columna SKU
    sobrescribir_destino : True = reemplaza filas en NEW
    """

    # --------------------------------------------------
    # 📥 LEER EXCEL
    # --------------------------------------------------
    df_es = pd.read_excel(excel_path, sheet_name=hoja_origen)
    df_new = pd.read_excel(excel_path, sheet_name=hoja_destino)

    if columna_sku not in df_es.columns or columna_sku not in df_new.columns:
        raise ValueError(f"❌ La columna '{columna_sku}' debe existir en ambas hojas")

    # Normalizar SKU
    df_es[columna_sku] = df_es[columna_sku].astype(str).str.strip()
    df_new[columna_sku] = df_new[columna_sku].astype(str).str.strip()

    # --------------------------------------------------
    # 🔍 DETECTAR COINCIDENCIAS
    # --------------------------------------------------
    skus_comunes = set(df_es[columna_sku]) & set(df_new[columna_sku])
    print(f"🔎 SKUs coincidentes encontrados: {len(skus_comunes)}")

    if not skus_comunes:
        print("⚠️ No hay SKUs coincidentes.")
        return

    # Filtrar filas ES coincidentes
    filas_es = df_es[df_es[columna_sku].isin(skus_comunes)]

    # --------------------------------------------------
    # 🧠 ACTUALIZAR DESTINO
    # --------------------------------------------------
    if sobrescribir_destino:
        # Eliminar filas antiguas en NEW
        df_new = df_new[~df_new[columna_sku].isin(skus_comunes)]

    # Añadir filas copiadas
    df_new_final = pd.concat([df_new, filas_es], ignore_index=True)

    # --------------------------------------------------
    # 💾 GUARDAR EXCEL
    # --------------------------------------------------
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_es.to_excel(writer, sheet_name=hoja_origen, index=False)
        df_new_final.to_excel(writer, sheet_name=hoja_destino, index=False)

    print("✅ Filas copiadas correctamente de ES → NEW")


def transferencia_productos_por_sku():
    def migrate_products_by_sku(models_src, db_src, uid_src, pwd_src,
                                models_dst, db_dst, uid_dst, pwd_dst,
                                excel_path):
        """
        Migra productos filtrados por SKU (default_code).
        SOLUCIÓN 1:
        - NO se leen imágenes en el read masivo
        - image_1920 se lee producto a producto
        """

        con_atributos = True

        # -------------------------------------------------
        # 1️⃣ Leer SKUs desde Excel
        # -------------------------------------------------
        df = pd.read_excel(excel_path, sheet_name="Sheet1")

        if "SKU" not in df.columns:
            raise Exception("❌ El Excel no contiene la columna 'SKU'")

        skus = (
            df["SKU"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
            .tolist()
        )

        if not skus:
            print("ℹ️ No hay SKUs en el Excel")
            return

        print(f"🧾 SKUs detectados: {len(skus)}")

        # -------------------------------------------------
        # 2️⃣ Campos (SIN image_1920)
        # -------------------------------------------------
        FIELDS = [
            'id',
            'active',
            'name',
            'default_code',
            'invoice_policy',
            'standard_price',
            'categ_id',
            'product_brand_id',
            'x_url',
            'description',
            'out_of_stock_message',
            'public_categ_ids'
        ]

        # -------------------------------------------------
        # 3️⃣ Buscar productos por SKU
        # -------------------------------------------------
        domain = [('default_code', 'in', skus)]

        product_ids = models_src.execute_kw(
            db_src, uid_src, pwd_src,
            'product.template', 'search',
            [domain],
            {'context': {'active_test': False}}
        )

        if not product_ids:
            print("ℹ️ No se encontraron productos para esos SKUs")
            return

        products = models_src.execute_kw(
            db_src, uid_src, pwd_src,
            'product.template', 'read',
            [product_ids],
            {'fields': FIELDS}
        )

        print(f"📦 Productos a migrar: {len(products)}")

        # -------------------------------------------------
        # 4️⃣ Migración producto a producto
        # -------------------------------------------------
        for product in products:
            try:
                # -------------------------
                # 🔎 Comprobar si ya existe en destino
                # -------------------------
                existing_ids = models_dst.execute_kw(
                    db_dst, uid_dst, pwd_dst,
                    'product.template', 'search',
                    [[('default_code', '=', product.get('default_code'))]],
                    {'limit': 1, 'context': {'active_test': False}}
                )

                if existing_ids:
                    print(f"⏭ SKU {product.get('default_code')} ya existe en destino (ID {existing_ids[0]}) — se omite")
                    continue

                # ---------------------------------------------
                # 🔹 Leer imagen SOLO de este producto
                # ---------------------------------------------
                image_data = models_src.execute_kw(
                    db_src, uid_src, pwd_src,
                    'product.template', 'read',
                    [[product['id']]],
                    {'fields': ['image_1920']}
                )[0]['image_1920']

                vals = {
                    'active': product.get('active', True),
                    'image_1920': image_data,
                    'name': product.get('name'),
                    'default_code': product.get('default_code'),
                    'invoice_policy': product.get('invoice_policy'),
                    'standard_price': product.get('standard_price') or 0.0,
                    'x_url': product.get('x_url'),
                    'description': product.get('description'),
                    'out_of_stock_message': product.get('out_of_stock_message'),
                    'list_price': 0.0
                }

                # -------------------------
                # Categoría interna
                # -------------------------
                if product.get('categ_id'):
                    categ_name = product['categ_id'][1]
                    categ_ids = models_dst.execute_kw(
                        db_dst, uid_dst, pwd_dst,
                        'product.category', 'search',
                        [[('name', '=', categ_name)]],
                        {'limit': 1}
                    )
                    if categ_ids:
                        vals['categ_id'] = categ_ids[0]

                # -------------------------
                # Marca
                # -------------------------
                if product.get('product_brand_id'):
                    brand_name = product['product_brand_id'][1]
                    brand_ids = models_dst.execute_kw(
                        db_dst, uid_dst, pwd_dst,
                        'product.brand', 'search',
                        [[('name', '=', brand_name)]],
                        {'limit': 1}
                    )
                    if brand_ids:
                        vals['product_brand_id'] = brand_ids[0]
                    else:
                        vals['product_brand_id'] = models_dst.execute_kw(
                            db_dst, uid_dst, pwd_dst,
                            'product.brand', 'create',
                            [{'name': brand_name}]
                        )

                # -------------------------
                # Categorías web
                # -------------------------
                if product.get('public_categ_ids'):
                    web_categ_ids = []
                    for wc in product['public_categ_ids']:
                        wc_name = models_src.execute_kw(
                            db_src, uid_src, pwd_src,
                            'product.public.category', 'read',
                            [[wc]],
                            {'fields': ['name']}
                        )[0]['name']

                        dst_wc = models_dst.execute_kw(
                            db_dst, uid_dst, pwd_dst,
                            'product.public.category', 'search',
                            [[('name', '=', wc_name)]],
                            {'limit': 1}
                        )

                        if dst_wc:
                            web_categ_ids.append(dst_wc[0])

                    if web_categ_ids:
                        vals['public_categ_ids'] = [(6, 0, web_categ_ids)]

                # -------------------------
                # Crear producto
                # -------------------------
                new_id = models_dst.execute_kw(
                    db_dst, uid_dst, pwd_dst,
                    'product.template', 'create',
                    [vals]
                )

                print(f"➕ Creado {vals.get('default_code')} (ID {new_id})")

                # -------------------------------------------------
                # 5️⃣ Migrar atributos
                # -------------------------------------------------
                if con_atributos:
                    attribute_lines = models_src.execute_kw(
                        db_src, uid_src, pwd_src,
                        'product.template.attribute.line', 'search_read',
                        [[('product_tmpl_id', '=', product['id'])]],
                        {'fields': ['attribute_id', 'value_ids']}
                    )

                    for line in attribute_lines:
                        attribute_name = line['attribute_id'][1]

                        attr_ids = models_dst.execute_kw(
                            db_dst, uid_dst, pwd_dst,
                            'product.attribute', 'search',
                            [[('name', '=', attribute_name)]],
                            {'limit': 1}
                        )

                        attr_id = attr_ids[0] if attr_ids else models_dst.execute_kw(
                            db_dst, uid_dst, pwd_dst,
                            'product.attribute', 'create',
                            [{'name': attribute_name}]
                        )

                        value_ids_dst = []

                        for value_id in line['value_ids']:
                            value_name = models_src.execute_kw(
                                db_src, uid_src, pwd_src,
                                'product.attribute.value', 'read',
                                [[value_id]],
                                {'fields': ['name']}
                            )[0]['name']

                            val_ids = models_dst.execute_kw(
                                db_dst, uid_dst, pwd_dst,
                                'product.attribute.value', 'search',
                                [[
                                    ('name', '=', value_name),
                                    ('attribute_id', '=', attr_id)
                                ]],
                                {'limit': 1}
                            )
                            val_id = val_ids[0] if val_ids else models_dst.execute_kw(
                                db_dst, uid_dst, pwd_dst,
                                'product.attribute.value', 'create',
                                [{
                                    'name': value_name,
                                    'attribute_id': attr_id
                                }]
                            )

                            value_ids_dst.append(val_id)

                            models_dst.execute_kw(
                                db_dst, uid_dst, pwd_dst,
                                'product.template.attribute.line', 'create',
                                [{
                                    'product_tmpl_id': new_id,
                                    'attribute_id': attr_id,
                                    'value_ids': [(6, 0, value_ids_dst)]
                                }]
                            )

                # -------------------------------------------------
                # 6️⃣ Galería (ojo: sigue siendo pesada)
                # -------------------------------------------------
                images = models_src.execute_kw(
                    db_src, uid_src, pwd_src,
                    'product.image', 'search_read',
                    [[('product_tmpl_id', '=', product['id'])]],
                    {'fields': ['image_1920', 'name', 'video_url', 'sequence']}
                )

                for img in images:
                    models_dst.execute_kw(
                        db_dst, uid_dst, pwd_dst,
                        'product.image', 'create',
                        [{
                            'product_tmpl_id': new_id,
                            'image_1920': img.get('image_1920'),
                            'name': img.get('name'),
                            'video_url': img.get('video_url'),
                            'sequence': img.get('sequence', 0)
                        }]
                    )

            except Exception as e:
                print(f"❌ Error en {product.get('default_code')}: {e}")

            print("✅ Migración por SKU finalizada")

    import xmlrpc.client
    # region LLAMADA
    # -----------------------
    # Configuración conexiones
    # -----------------------

    origen = {
        'url': "http://79.72.61.76:8069/",#"https://optimaluz.soluntec.net",
        'db': "odoo0",#"Real",
        'user': "admin",#"jcoronado@optimaluz.com",
        'password': "admin",#"AlAi4ever"
    }

    destino = {
        'url': "http://143.47.33.148:8070/",
        'db': "odoo1",  # "odoo0",
        'user': "admin",
        'password': "admin"
    }

    # -----------------------
    # Conexión ORIGEN
    # -----------------------

    common_src = xmlrpc.client.ServerProxy(
        f"{origen['url']}/xmlrpc/2/common",
        allow_none=True
    )

    uid_src = common_src.authenticate(
        origen['db'],
        origen['user'],
        origen['password'],
        {}
    )

    if not uid_src:
        raise Exception("❌ No se pudo autenticar en ORIGEN")

    models_src = xmlrpc.client.ServerProxy(
        f"{origen['url']}/xmlrpc/2/object",
        allow_none=True
    )

    print(f"🔌 Conectado a ORIGEN (uid={uid_src})")

    # -----------------------
    # Conexión DESTINO
    # -----------------------

    common_dst = xmlrpc.client.ServerProxy(
        f"{destino['url']}/xmlrpc/2/common",
        allow_none=True
    )

    uid_dst = common_dst.authenticate(
        destino['db'],
        destino['user'],
        destino['password'],
        {}
    )

    if not uid_dst:
        raise Exception("❌ No se pudo autenticar en DESTINO")

    models_dst = xmlrpc.client.ServerProxy(
        f"{destino['url']}/xmlrpc/2/object",
        allow_none=True
    )

    print(f"🔌 Conectado a DESTINO (uid={uid_dst})")

    # -----------------------
    # LLAMADA A LA MIGRACIÓN
    # -----------------------

    migrate_products_by_sku(
        models_src=models_src,
        db_src=origen['db'],
        uid_src=uid_src,
        pwd_src=origen['password'],

        models_dst=models_dst,
        db_dst=destino['db'],
        uid_dst=uid_dst,
        pwd_dst=destino['password'],

        excel_path=ruta
    )
    # endregion


def actualizar_categorias_por_sku(models_src, db_src, uid_src, pwd_src,
                                  models, db, uid, pwd,
                                  excel_path):

    try:
        df = pd.read_excel(excel_path)
        only_cats=False
        skus = (
            df["SKU"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
            .tolist()
        )

        if not skus:
            print("ℹ️ No hay SKUs")
            return

        print(f"🧾 SKUs detectados: {len(skus)}")

        # -------------------------------------------------
        # 🔹 Productos ORIGEN
        # -------------------------------------------------

        src_products = models_src.execute_kw(
            db_src, uid_src, pwd_src,
            'product.template', 'search_read',
            [[('default_code', 'in', skus)]],
            {'fields': ['default_code', 'name', 'public_categ_ids', 'is_favorite']}
        )

        src_map = {p['default_code']: p for p in src_products}

        # -------------------------------------------------
        # 🔹 Productos DESTINO
        # -------------------------------------------------

        dst_products = models.execute_kw(
            db, uid, pwd,
            'product.template', 'search_read',
            [[('default_code', 'in', skus)]],
            {'fields': ['id', 'default_code', 'public_categ_ids']}
        )

        dst_map = {p['default_code']: p for p in dst_products}

        # -------------------------------------------------
        # 🔹 Categorías públicas DESTINO
        # -------------------------------------------------

        public_dst = models.execute_kw(
            db, uid, pwd,
            'product.public.category', 'search_read',
            [[]],
            {'fields': ['id', 'parent_id', 'x_id_interno']}
        )

        public_map_by_xid = {
            c['x_id_interno']: c
            for c in public_dst
            if c.get('x_id_interno')
        }

        public_map_by_id = {c['id']: c for c in public_dst}

        # -------------------------------------------------
        # 🔹 Categorías internas DESTINO
        # -------------------------------------------------

        internal_cats = models.execute_kw(
            db, uid, pwd,
            'product.category', 'search_read',
            [[]],
            {'fields': ['id', 'name']}
        )

        internal_map = {c['name']: c['id'] for c in internal_cats}

        # -------------------------------------------------
        # 🔹 Función subir a raíz pública
        # -------------------------------------------------

        def obtener_raiz(cat_id):
            while True:
                cat = public_map_by_id.get(cat_id)
                if not cat:
                    return cat_id
                parent = cat['parent_id'][0] if cat['parent_id'] else False
                if not parent:
                    return cat_id
                cat_id = parent

        # -------------------------------------------------
        # 🔹 LOOP PRINCIPAL
        # -------------------------------------------------

        for sku in skus:

            if sku not in src_map:
                print(f"⚠ No existe en origen: {sku}")
                continue

            if sku not in dst_map:
                print(f"⚠ No existe en destino: {sku}")
                continue

            product_src = src_map[sku]
            product_dst = dst_map[sku]

            # 🚀 Si ya tiene categorías públicas → saltar
            if product_dst.get('public_categ_ids') and only_cats:
                print(f"⏭ Ya tiene categorías asignadas: {sku}")
                continue

            new_public_ids = []
            raiz_public_id = False

            for pub_id in product_src.get('public_categ_ids', []):

                if pub_id not in public_map_by_xid:
                    print(f"⚠ Categoría pública no encontrada por x_id_interno: {pub_id}")
                    continue

                cat_dst = public_map_by_xid[pub_id]
                new_public_ids.append(cat_dst['id'])

                if not raiz_public_id:
                    raiz_public_id = obtener_raiz(cat_dst['id'])

            if not new_public_ids and only_cats:
                print(f"⚠ No se pudieron resolver categorías públicas para {sku}")
                continue

            vals = {}

            if new_public_ids:
                vals = {'public_categ_ids': [(6, 0, new_public_ids)]}

                # 🔹 Resolver categ_id (modelo interno) usando nombre de la raíz pública
                if raiz_public_id:

                    raiz_public = public_map_by_id.get(raiz_public_id)

                    if raiz_public:
                        # necesitamos el name real → lo leemos una sola vez
                        raiz_data = models.execute_kw(
                            db, uid, pwd,
                            'product.public.category', 'read',
                            [[raiz_public_id]],
                            {'fields': ['name']}
                        )[0]

                        raiz_name = raiz_data['name']

                        if raiz_name in internal_map:
                            vals['categ_id'] = internal_map[raiz_name]
                        else:
                            print(f"⚠ No existe categoría interna para raíz: {raiz_name}")

            vals['name'] = product_src.get('name')
            #vals['is_favorite'] = product_src.get('is_favorite')

            models.execute_kw(
                db, uid, pwd,
                'product.template', 'write',
                [[product_dst['id']], vals]
            )

            print(f"✅ Actualizado: {sku}")

        print("🎯 Proceso finalizado correctamente")

    except Exception as e:
        print(f"❌ Error: {e}")

def migrar_supplierinfo_por_sku(models_src, db_src, uid_src, pwd_src,
                                models, db, uid, pwd,
                                excel_path):

    try:
        # -------------------------------------------------
        # 1️⃣ Leer SKUs
        # -------------------------------------------------

        df = pd.read_excel(excel_path)

        skus = (
            df["SKU"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
            .tolist()
        )

        if not skus:
            print("ℹ️ No hay SKUs")
            return

        print(f"🧾 SKUs detectados: {len(skus)}")

        # -------------------------------------------------
        # 2️⃣ Leer productos ORIGEN
        # -------------------------------------------------

        src_products = models_src.execute_kw(
            db_src, uid_src, pwd_src,
            'product.template', 'search_read',
            [[('default_code', 'in', skus), ('active', 'in', [True, False])]],
            {'fields': ['id', 'default_code']}
        )

        src_map = {p['id']: p['default_code'] for p in src_products}

        if not src_map:
            print("ℹ️ No se encontraron productos en origen")
            return

        # -------------------------------------------------
        # 3️⃣ Leer productos DESTINO
        # -------------------------------------------------

        dst_products = models.execute_kw(
            db, uid, pwd,
            'product.template', 'search_read',
            [[('default_code', 'in', list(src_map.values())), ('active', 'in', [True, False])]],
            {'fields': ['id', 'default_code']}
        )

        dst_map = {p['default_code']: p['id'] for p in dst_products}

        # -------------------------------------------------
        # 4️⃣ Leer supplierinfo ORIGEN
        # -------------------------------------------------

        supplierinfos = models_src.execute_kw(
            db_src, uid_src, pwd_src,
            'product.supplierinfo', 'search_read',
            [[('product_tmpl_id', 'in', list(src_map.keys()))]],
            {'fields': ['product_tmpl_id',
                        'partner_id',
                        'product_name',
                        'product_code',
                        'price', 'min_qty']}
        )

        if not supplierinfos:
            print("ℹ️ No hay supplierinfo en origen")
            return

        print(f"📦 Supplierinfos encontrados: {len(supplierinfos)}")

        # -------------------------------------------------
        # 5️⃣ Procesar uno a uno
        # -------------------------------------------------

        for sup in supplierinfos:

            src_product_id = sup['product_tmpl_id'][0]
            sku = src_map.get(src_product_id)

            if not sku or sku not in dst_map:
                print(f"⚠ Producto destino no encontrado para SKU {sku}")
                continue

            dst_product_id = dst_map[sku]

            # 🔹 Resolver partner por x_id_interno
            if not sup.get('partner_id'):
                print(f"⚠ Supplier sin partner en SKU {sku}")
                continue

            partner_src_id = sup['partner_id'][0]

            partner_dst = models.execute_kw(
                db, uid, pwd,
                'res.partner', 'search',
                [[('x_id_interno', '=', partner_src_id)]],
                {'limit': 1}
            )

            if not partner_dst:
                print(f"⚠ Partner no encontrado en destino (x_id_interno={partner_src_id})")
                continue

            partner_dst_id = partner_dst[0]

            '''# 🔹 Evitar duplicado
            existing = models.execute_kw(
                db, uid, pwd,
                'product.supplierinfo', 'search',
                [[
                    ('product_tmpl_id', '=', dst_product_id),
                    ('partner_id', '=', partner_dst_id),
                    ('product_code', '=', sup.get('product_code'))
                ]],
                {'limit': 1}
            )

            if existing:
                print(f"⏭ Ya existe supplierinfo para SKU {sku}")
                continue'''

            # 🔹 Crear supplierinfo
            vals = {
                'product_tmpl_id': dst_product_id,
                'partner_id': partner_dst_id,
                'product_name': sup.get('product_name'),
                'product_code': sup.get('product_code'),
                'price': sup.get('price') or 0.0,
                'min_qty': sup.get('min_qty') or 1.0,
            }

            models.execute_kw(
                db, uid, pwd,
                'product.supplierinfo', 'create',
                [vals]
            )

            print(f"✅ Supplierinfo migrado para {sku}")

        print("🎯 Migración supplierinfo finalizada")

    except Exception as e:
        print(f"❌ Error migrando supplierinfo: {e}")

def migrar_media_y_web_por_sku(models_src, db_src, uid_src, pwd_src,
                               models, db, uid, pwd,
                               excel_path,
                               batch_size=20):  # 🔥 batch pequeño

    try:
        df = pd.read_excel(excel_path)

        skus = (
            df["SKU"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
            .tolist()
        )

        if not skus:
            print("ℹ️ No hay SKUs")
            return

        print(f"🧾 SKUs detectados: {len(skus)}")

        for i in range(0, len(skus), batch_size):

            batch = skus[i:i + batch_size]

            # -------------------------------------------------
            # 🔹 Leer solo IDs ORIGEN (sin binarios)
            # -------------------------------------------------

            templates_src = models_src.execute_kw(
                db_src, uid_src, pwd_src,
                'product.template', 'search_read',
                [[('default_code', 'in', batch)]],
                {'fields': ['id', 'default_code']}
            )

            src_map = {t['default_code']: t['id'] for t in templates_src}

            # -------------------------------------------------
            # 🔹 Leer solo IDs DESTINO
            # -------------------------------------------------

            templates_dst = models.execute_kw(
                db, uid, pwd,
                'product.template', 'search_read',
                [[('default_code', 'in', batch)]],
                {'fields': ['id', 'default_code', 'website_description']}
            )

            dst_map = {t['default_code']: t for t in templates_dst}

            for sku in batch:

                if sku not in src_map or sku not in dst_map:
                    continue

                tmpl_id_src = src_map[sku]
                product_dst = dst_map[sku]

                if product_dst.get('website_description'):
                    print(f"⏭ Ya tiene descripción web: {sku}")
                    continue

                tmpl_id_dst = product_dst['id']

                # -------------------------------------------------
                # 🔹 Ahora sí leer BINARIOS SOLO de este producto
                # -------------------------------------------------

                product_src = models_src.execute_kw(
                    db_src, uid_src, pwd_src,
                    'product.template', 'read',
                    [[tmpl_id_src]],
                    {'fields': [
                        'image_1920',
                        'website_description'
                    ] + [f'x_icono{i}' for i in range(1, 9)]}
                )[0]

                vals = {}

                if product_src.get('image_1920'):
                    vals['image_1920'] = product_src['image_1920']

                if product_src.get('website_description'):
                    vals['website_description'] = product_src['website_description']

                for idx in range(1, 9):
                    campo = f'x_icono{idx}'
                    if product_src.get(campo):
                        vals[campo] = product_src[campo]

                if vals:
                    models.execute_kw(
                        db, uid, pwd,
                        'product.template', 'write',
                        [[tmpl_id_dst], vals]
                    )

                # -------------------------------------------------
                # 🔹 GALERÍA (leer solo del producto actual)
                # -------------------------------------------------

                images_src = models_src.execute_kw(
                    db_src, uid_src, pwd_src,
                    'product.image', 'search_read',
                    [[('product_tmpl_id', '=', tmpl_id_src)]],
                    {'fields': ['image_1920', 'name', 'video_url', 'sequence']}
                )

                # Borrar galería destino
                existing_imgs = models.execute_kw(
                    db, uid, pwd,
                    'product.image', 'search',
                    [[('product_tmpl_id', '=', tmpl_id_dst)]]
                )

                if existing_imgs:
                    models.execute_kw(
                        db, uid, pwd,
                        'product.image', 'unlink',
                        [existing_imgs]
                    )

                # Insertar nuevas
                for img in images_src:

                    if not img.get('image_1920') and not img.get('video_url'):
                        continue

                    models.execute_kw(
                        db, uid, pwd,
                        'product.image', 'create',
                        [{
                            'product_tmpl_id': tmpl_id_dst,
                            'image_1920': img.get('image_1920'),
                            'video_url': img.get('video_url'),
                            'sequence': img.get('sequence', 10),
                            'name': img.get('name'),
                        }]
                    )

                print(f"✅ Media migrada: {sku}")

        print("🎯 Migración multimedia finalizada sin problemas de memoria")

    except Exception as e:
        print(f"❌ Error migrando multimedia: {e}")

def migrar_ausencias(models_src, db_src, uid_src, pwd_src,
                     models, db, uid, pwd):

    try:

        # -------------------------------------------------
        # 🔹 1️⃣ Leer ausencias ORIGEN
        # -------------------------------------------------

        leaves_src = models_src.execute_kw(
            db_src, uid_src, pwd_src,
            'hr.leave', 'search_read',
            [[("employee_id.active", "=", True)]],
            {'fields': [
                'employee_id',
                'department_id',
                'request_date_from',
                'request_date_to',
                'date_from',
                'date_to',
                'name',
                'holiday_status_id',
                'state'
            ]}
        )

        if not leaves_src:
            print("ℹ️ No hay ausencias en origen")
            return

        print(f"🗓 Ausencias detectadas: {len(leaves_src)}")

        # -------------------------------------------------
        # 🔹 2️⃣ Precargar empleados DESTINO
        # -------------------------------------------------

        employees_dst = models.execute_kw(
            db, uid, pwd,
            'hr.employee', 'search_read',
            [[]],
            {'fields': ['id', 'x_id_interno']}
        )

        employee_map = {
            e['x_id_interno']: e['id']
            for e in employees_dst
            if e.get('x_id_interno')
        }

        # -------------------------------------------------
        # 🔹 3️⃣ Precargar departamentos DESTINO
        # -------------------------------------------------

        departments_dst = models.execute_kw(
            db, uid, pwd,
            'hr.department', 'search_read',
            [[]],
            {'fields': ['id', 'name']}
        )

        department_map = {d['name']: d['id'] for d in departments_dst}

        # -------------------------------------------------
        # 🔹 4️⃣ Precargar tipos de ausencia DESTINO
        # -------------------------------------------------

        types_dst = models.execute_kw(
            db, uid, pwd,
            'hr.leave.type', 'search_read',
            [[]],
            {'fields': ['id', 'x_id_interno']}
        )

        type_map = {
            e['x_id_interno']: e['id']
            for e in types_dst
            if e.get('x_id_interno')
        }

        # -------------------------------------------------
        # 🔹 5️⃣ Migración
        # -------------------------------------------------

        for leave in leaves_src:

            # 🔹 Resolver empleado
            if not leave.get('employee_id'):
                print("⚠ Ausencia sin empleado")
                continue

            employee_src_id = leave['employee_id'][0]

            if employee_src_id not in employee_map:
                print(f"⚠ Empleado no encontrado (x_id_interno={employee_src_id})")
                continue

            employee_dst_id = employee_map[employee_src_id]

            # 🔹 Resolver departamento
            department_dst_id = False
            if leave.get('department_id'):
                dept_name = leave['department_id'][1]
                department_dst_id = department_map.get(dept_name)

                if not department_dst_id:
                    print(f"⚠ Departamento no encontrado: {dept_name}")

            # 🔹 Resolver tipo ausencia
            type_src_id = leave['holiday_status_id'][0]

            if type_src_id not in type_map:
                print(f"⚠ Tipo de ausencia no encontrado: (x_id_interno={type_src_id})")
                continue

            type_dst_id = type_map[type_src_id]



            # 🔹 Preparar valores
            vals = {
                'employee_id': employee_dst_id,
                'holiday_status_id': type_dst_id,
                'request_date_from': leave.get('request_date_from'),
                'request_date_to': leave.get('request_date_to'),
                'date_from': leave.get('date_from'),
                'date_to': leave.get('date_to'),
                'name': leave.get('name'),
            }

            if department_dst_id:
                vals['department_id'] = department_dst_id

            # -------------------------------------------------
            # 🔹 Crear en estado borrador
            # -------------------------------------------------

            leave_id = models.execute_kw(
                db, uid, pwd,
                'hr.leave', 'create',
                [vals]
            )

            # -------------------------------------------------
            # 🔹 Ajustar estado
            # -------------------------------------------------

            state = leave.get('state')

            if state == 'confirm':
                models.execute_kw(db, uid, pwd,
                                  'hr.leave', 'action_confirm',
                                  [[leave_id]])

            elif state == 'validate':
                models.execute_kw(db, uid, pwd,
                                  'hr.leave', 'action_validate',
                                  [[leave_id]])

            elif state == 'refuse':
                models.execute_kw(db, uid, pwd,
                                  'hr.leave', 'action_refuse',
                                  [[leave_id]])

            print(f"✅ Ausencia migrada: {leave.get('name')}")

        print("🎯 Migración de ausencias finalizada")

    except Exception as e:
        print(f"❌ Error migrando ausencias: {e}")

def ejecutar_funciones_transferencia():
    # region CONEXION
    # -----------------------
    # Configuración conexiones
    # -----------------------

    origen = {
        'url': "http://158.179.220.107:8069/",#"https://optimaluz.soluntec.net",#
        'db': "odoo0",#"Real",#
        'user': "admin",#"jcoronado@optimaluz.com",#
        'password': "admin",#"AlAi4ever"#
    }

    destino = {
        'url': "http://141.253.197.145:8070/",
        'db': "odoo1",  # "odoo0",
        'user': "admin",
        'password': "admin"
    }

    # -----------------------
    # Conexión ORIGEN
    # -----------------------

    common_src = xmlrpc.client.ServerProxy(
        f"{origen['url']}/xmlrpc/2/common",
        allow_none=True
    )

    uid_src = common_src.authenticate(
        origen['db'],
        origen['user'],
        origen['password'],
        {}
    )

    if not uid_src:
        raise Exception("❌ No se pudo autenticar en ORIGEN")

    models_src = xmlrpc.client.ServerProxy(
        f"{origen['url']}/xmlrpc/2/object",
        allow_none=True
    )

    print(f"🔌 Conectado a ORIGEN (uid={uid_src})")

    # -----------------------
    # Conexión DESTINO
    # -----------------------

    common_dst = xmlrpc.client.ServerProxy(
        f"{destino['url']}/xmlrpc/2/common",
        allow_none=True
    )

    uid_dst = common_dst.authenticate(
        destino['db'],
        destino['user'],
        destino['password'],
        {}
    )

    if not uid_dst:
        raise Exception("❌ No se pudo autenticar en DESTINO")

    models_dst = xmlrpc.client.ServerProxy(
        f"{destino['url']}/xmlrpc/2/object",
        allow_none=True
    )

    print(f"🔌 Conectado a DESTINO (uid={uid_dst})")
    # endregion

    # -----------------------
    # LLAMADA A LA MIGRACIÓN
    # -----------------------

    actualizar_categorias_por_sku(
        models_src=models_src,
        db_src=origen['db'],
        uid_src=uid_src,
        pwd_src=origen['password'],

        models=models_dst,
        db=destino['db'],
        uid=uid_dst,
        pwd=destino['password'],

        excel_path=ruta
    )


#ejecutar_funciones_transferencia()

def rellenar_nuevos(path_excel=ruta, hoja=0):
    '''
    Detecta “nuevos” comparando las columnas SKU_2 contra SKU_1.
    Concretamente, considera nuevo los SKUs que aparecen en SKU_2 pero no aparece en SKU_1
    '''
    # Leer Excel
    df = pd.read_excel(path_excel, sheet_name=hoja)

    # Normalizar a string y eliminar NaN
    sku1 = df["SKU_1"].dropna().astype(str).str.strip()
    sku2 = df["SKU_2"].dropna().astype(str).str.strip()

    # Convertir SKU_1 a set para comparación rápida
    sku1_set = set(sku1)

    # Detectar nuevos
    nuevos = [sku for sku in sku2 if sku not in sku1_set]

    # Vaciar columna NUEVOS
    df["NUEVOS"] = None

    # Escribir resultados
    df.loc[:len(nuevos)-1, "NUEVOS"] = nuevos

    # Guardar Excel
    df.to_excel(path_excel, index=False)

    print(f"SKUs nuevos encontrados: {len(nuevos)}")

def extraer_skus_pdf_catalogo_a_excel(pdf_path, excel_salida="skus_extraidos.xlsx"):
    import pdfplumber
    import re

    skus = {}

    # patrón SKU válido
    patron_sku = re.compile(r'^\d{4,}(?:-\d+)?$')

    with pdfplumber.open(pdf_path) as pdf:
        for num_pagina, page in enumerate(pdf.pages, start=1):

            text = page.extract_text()

            if not text:
                continue

            for line in text.split("\n"):
                line = line.strip()

                if "|" in line:
                    posible_sku = line.split("|")[0].strip()
                else:
                    match = re.match(r'^\S+', line)
                    if not match:
                        continue
                    posible_sku = match.group()

                # validar SKU
                if patron_sku.match(posible_sku):

                    # guardar solo la primera vez que aparece
                    if posible_sku not in skus:
                        skus[posible_sku] = num_pagina

    # crear dataframe
    df = pd.DataFrame(
        [(sku, pagina) for sku, pagina in skus.items()],
        columns=["SKU", "Pagina"]
    )

    df.sort_values("SKU", inplace=True)

    df.to_excel(excel_salida, index=False)

    print(f"✔ {len(skus)} SKUs extraídos")
    print(f"📁 Excel generado: {excel_salida}")
