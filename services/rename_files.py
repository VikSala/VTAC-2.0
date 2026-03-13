import os
import pandas as pd
import ast
import base64
import json

def actualizar_galeria_con_imagenes(ruta_excel, ruta_imagenes):
    df = pd.read_excel(ruta_excel)

    for index, fila in df.iterrows():
        sku = str(fila['default_code']).strip()
        galeria_raw = fila['Galeria']

        # Convertimos la columna Galeria a lista
        try:
            galeria = ast.literal_eval(galeria_raw) if isinstance(galeria_raw, str) else []
        except:
            galeria = []

        if not isinstance(galeria, list):
            galeria = []

        # Ruta al subdirectorio con nombre SKU
        ruta_sku = os.path.join(ruta_imagenes, sku)
        if not os.path.isdir(ruta_sku):
            continue

        # Buscar archivos tipo N.jpg (1.jpg, 2.jpg, etc.)
        archivos = [f for f in os.listdir(ruta_sku) if f.endswith(".jpg") and f[:-4].isdigit()]
        archivos_ordenados = sorted(archivos, key=lambda x: int(x[:-4]))  # ordenar por número

        for nombre_archivo in archivos_ordenados:
            posicion = int(nombre_archivo[:-4]) - 1  # 1.jpg → posición 0
            valor = f"{sku}/{nombre_archivo}"

            # Insertar en la posición, desplazando si ya hay algo
            if posicion >= len(galeria):
                galeria.extend([None] * (posicion - len(galeria) + 1))
            galeria.insert(posicion, valor)

        # Limpiar valores None
        galeria = [x for x in galeria if x is not None]

        # Guardar la nueva galería en el DataFrame
        df.at[index, 'Galeria'] = str(galeria)

    # Guardar nuevo Excel
    df.to_excel(ruta_excel, index=False)
    print(f"Archivo actualizado guardado en: {ruta_excel}")

def renombrar_archivos(ruta_base):
    for carpeta_raiz, _, archivos in os.walk(ruta_base):
        for archivo in archivos:
            ruta_original = os.path.join(carpeta_raiz, archivo)

            nombre_base, _ = os.path.splitext(archivo)
            nuevo_nombre = nombre_base.replace("galeria_", "") + ".jpg"
            nueva_ruta = os.path.join(carpeta_raiz, nuevo_nombre)

            # Evitar sobreescribir si ya existe un archivo con ese nombre
            if ruta_original != nueva_ruta:
                try:
                    os.rename(ruta_original, nueva_ruta)
                    print(f"Renombrado: {archivo} -> {nuevo_nombre}")
                except Exception as e:
                    print(f"Error renombrando {archivo}: {e}")


def guardar_imagenes_desde_excel(ruta_excel, ruta_imgs):
    df = pd.read_excel(ruta_excel)

    for index, fila in df.iterrows():
        sku = str(fila['default_code']).strip()
        imagen_base64 = fila['image_1920']

        # Validar que la celda tiene contenido
        if not isinstance(imagen_base64, str) or not imagen_base64.strip():
            print(f"[{sku}] Sin imagen válida. Saltando.")
            continue

        try:
            # Crear carpeta destino si no existe
            ruta_destino = os.path.join(ruta_imgs, sku)
            os.makedirs(ruta_destino, exist_ok=True)

            # Decodificar y guardar como 0.jpg
            ruta_archivo = os.path.join(ruta_destino, "0.jpg")

            with open(ruta_archivo, "wb") as f:
                f.write(base64.b64decode(imagen_base64))


            print(f"[{sku}] Imagen guardada en: {ruta_archivo}")

        except Exception as e:
            print(f"[{sku}] Error procesando imagen: {e}")

def guardar_imagenes_desde_json(ruta_json, ruta_imgs):
    with open(ruta_json, "r", encoding="utf-8") as f:
        productos = json.load(f)

    for producto in productos:
        sku = str(producto.get('default_code', '')).strip()
        imagen_base64 = producto.get('image_1920', '')

        # Validar SKU e imagen
        if not sku or not isinstance(imagen_base64, str) or not imagen_base64.strip():
            print(f"[{sku or '???'}] Sin imagen válida. Saltando.")
            continue

        try:
            # Limpiar base64
            imagen_base64 = imagen_base64.replace("\\/", "/").strip()
            faltan = len(imagen_base64) % 4
            if faltan:
                imagen_base64 += "=" * (4 - faltan)

            # Crear carpeta si no existe
            ruta_destino = os.path.join(ruta_imgs, sku)
            os.makedirs(ruta_destino, exist_ok=True)

            ruta_archivo = os.path.join(ruta_destino, "0.jpg")
            with open(ruta_archivo, "wb") as f:
                f.write(base64.b64decode(imagen_base64))

            print(f"[{sku}] Imagen guardada en: {ruta_archivo}")

        except Exception as e:
            print(f"[{sku}] Error procesando imagen: {e}")


def actualizar_columna_image_con_sku(ruta_excel):
    df = pd.read_excel(ruta_excel)

    if 'SKU' not in df.columns and 'default_code' in df.columns:
        df['SKU'] = df['default_code']

    if 'SKU' not in df.columns:
        raise ValueError("No se encontró la columna 'SKU' en el Excel.")

    # Reescribir la columna 'Image' con formato SKU/0.jpg
    df['Image'] = df['SKU'].astype(str).str.strip() + '/0.jpg'

    df.to_excel(ruta_excel, index=False)
    print(f"✅ Excel actualizado guardado en: {ruta_excel}")

if __name__ == "__main__":
    ruta_objetivo = r"D:\BANCOS\OPTIMA\main-img"
    actualizar_columna_image_con_sku("v_tac_products.xlsx")#guardar_imagenes_desde_json("productos.json", ruta_objetivo)#guardar_imagenes_desde_excel("v_tac_products.xlsx", ruta_objetivo)#actualizar_galeria_con_imagenes("v_tac_products.xlsx", ruta_objetivo)
