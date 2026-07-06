from services.campos_odoo import ClavesExcel
import pandas as pd
import ast

#region Atributos Excel

# España
MAP_ES = {
    "Adicional": "Comentarios",
    "Cantidad del LED": "Cantidad de LED",
    "Código de orden": "SKU",
    #"Código de producto": "Código de familia",
    "Dimensiones de montaje": "Dimensiones",
    "Jacket Dia": "Jacket Diameter",
    "Modo de carga corriente de entrada": "",
    "Modo de emergencia corriente de salida": "",
    "Product Features": "",
    "Protección IP": "Grado de protección IP",
    "Tipo casquillo": "Casquillo",
    "Ángulo de haz": "Ángulo de Apertura",
    "CRI 90+": "CRI",
    "Montar": "Montaje"
}

# Italia
MAP_IT = {
    "Forma del foco": "Ampolla exterior",
    "CRI": "CRI",
    "Número de LED": "Cantidad de LED",
    "Número de LED por metro": "Cantidad de LED",
    "Capacidad de la batería": "Capacidad de la batería",
    "Casquillo": "Casquillo",
    "Clase de eficiencia": "Clase de eficiencia energética",
    "Classe efficienza": "Clase de eficiencia energética",
    "Color": "Color del cuerpo",
    "Colore": "Color del cuerpo",
    "Corriente máxima": "Corriente",
    "SKU": "SKU",
    "Código SKU": "SKU",
    "Código de familia": "Código de producto",#Código de familia
    "Codice famiglia": "Código de producto",#Código de familia
    "Tamaño del producto": "Dimensiones",
    "Dimensione Prodotto": "Dimensiones",
    "EAN": "EAN Código",
    "Eficiencia luminosa": "Lúmenes útiles (lm)",
    "Factor de potencia": "Factor de potencia",
    "Fattore di potenza": "Factor de potencia",
    "Sistema eléctrico inversor": "Fase",
    "Flujo luminoso": "Flujo luminoso (lm)",
    "Flujo luminoso por metro": "Flujo luminoso/m",
    "Flusso Luminoso al Metro": "Flujo luminoso/m",
    "Garantía": "Garantía",
    "Garanzia": "Garantía",
    "Clasificación de protección contra impactos": "Grado de protección IK",
    "Grado de protección": "Grado de protección IP",
    "Grado di Protezione": "Grado de protección IP",
    "Longitud de la bobina": "Longitud",
    "Lunghezza della Bobina": "Longitud",
    "Longitud del cable": "Longitud del cable",
    "Marca": "Marca",
    "Marchio": "Marca",
    "Material": "Material",
    "Materiale": "Material",
    "Tipo de instalación": "Montaje",
    "Installation Type": "Montaje",
    "Peso": "Peso del artículo",
    "Fuerza": "Potencia",
    "Energía de paneles solares": "Potencia",
    "Potencia equivalente": "Potencia equivalente",
    "Potencia por metro": "Potencia/m",
    "Potenza al Metro": "Potencia/m",
    "Regulable": "Regulable",
    "Dimmerabile": "Regulable",
    "Temperatura de color (CTI)": "Temperatura de color",
    "Temperatura di Colore (CTI)": "Temperatura de color",
    "Tipo de fuente de alimentación": "Tensión",
    "Tiempo de carga": "Tiempo de carga",
    "Autonomía": "Tiempo de trabajo",
    "Tipo di Installazione": "Montaje",
    "tipo LED": "Tipo LED",
    "Color de la luz": "Tono de luz",
    "Collore della Luce": "Tono de luz",
    "Piezas en la caja": "Unidades por caja",
    "Pezzi nella scatola": "Unidades por caja",
    "Duración": "Vida útil",
    "Durata": "Vida útil",
    "Voltaje de suministro": "Voltaje de entrada",
    "Tensione di Alimentazione": "Voltaje de entrada",
    "Volumen": "Volumen del artículo",
    "Volume": "Volumen del artículo",
    "Rayo de luz": "Ángulo de Apertura",
    "Fascio Luminoso": "Ángulo de Apertura",
}

# UK
MAP_UK = {
    "CRI": "CRI",
    "Número de LED": "Cantidad de LED",
    "Tipo de batería": "Capacidad de la batería",
    "Carga nominal": "Carga nominal",
    "Base": "Casquillo",
    "Etiqueta de clasificación energética": "Clase de eficiencia energética",
    "Energy Rating Label": "Clase de eficiencia energética",
    "SKU": "SKU",
    "Modelo": "Código de producto", # Código de familia
    "Model": "Código de producto", # Código de familia
    "Velocidad de detección de movimiento": "Detección por movimiento",
    "Detection Motion Speed": "Detección por movimiento",
    "Dimensión": "Dimensiones",
    "Dimension": "Dimensiones",
    "Código Ean": "EAN Código",
    "Ean Code": "EAN Código",
    "PF": "Factor de potencia",
    "Lúmenes": "Flujo luminoso (lm)",
    "Lumens": "Flujo luminoso (lm)",
    "Lúmenes por metro": "Flujo luminoso/m",
    "En: Potencia de entrada": "Voltaje de entrada",
    "Clasificación IK": "Grado de protección IK",
    "IK Rating": "Grado de protección IK",
    "Inicio instantáneo": "Hora de inicio al 100% encendido",
    "Instant Start": "Hora de inicio al 100% encendido",
    "Cantidad de rollos": "Longitud",
    "Longitud del cable": "Longitud del cable",
    "Wire Length": "Longitud del cable",
    "Luz ambiental": "Luz ambiental",
    "Ambient Light": "Luz ambiental",
    "lúmenes por vatio": "Lúmenes útiles (lm)",
    "lumens per watt": "Lúmenes útiles (lm)",
    "Corriente de carga de la batería": "Modo de carga corriente de carga",
    "Corriente de descarga de la batería": "Modo de emergencia corriente de descarga",
    "Tipo de instalación": "Montaje",
    "Peso bruto (kg)": "Peso del artículo",
    "Gross Weight (Kgs)": "Peso del artículo",
    "Vatios": "Potencia",
    "Watts": "Potencia",
    "Potencia equivalente": "Potencia equivalente",
    "Vatios por metro": "Potencia/m",
    "Clasificación IP": "Grado de protección IP",
    "IP Rating": "Grado de protección IP",
    "Ángulo de detección": "Rango de detección",
    "Detection Angle": "Rango de detección",
    "Regulable": "Regulable",
    "Dimmable": "Regulable",
    "Voltaje de salida": "Salida",
    "Output Voltage": "Salida",
    "Temperatura de color": "Temperatura de color",
    "Color Temperature": "Temperatura de color",
    "Potencia de entrada": "Tensión",
    "Input Power": "Tensión",
    "Tiempo de carga": "Tiempo de carga",
    "Charging Time": "Tiempo de carga",
    "Retraso_Retardo de tiempo": "Tiempo de retardo",
    "Time Delay": "Tiempo de retardo",
    "Tiempo de trabajo": "Tiempo de trabajo",
    "Tipo de chip LED": "Tipo LED",
    "LED Chip Type": "Tipo LED",
    "Tipo de sensor": "Tipo de sensor",
    "Sensor Type": "Tipo de sensor",
    "Color para la Web": "Tono de luz",
    "Color for Web": "Tono de luz",
    "Cantidad por caja": "Unidades por caja",
    "Box Qty": "Unidades por caja",
    "Larga vida": "Vida útil",
    "Long Life": "Vida útil",
    "Unidad CBM": "Volumen del artículo",
    "Unit CBM": "Volumen del artículo",
    "Ángulo de haz": "Ángulo de Apertura",
    "Beam Angle": "Ángulo de Apertura",
}

# Mapeos agrupados por país
ATTRIBUTE_MAPS = {
    "es": MAP_ES,
    "it": MAP_IT,
    "uk": MAP_UK,
}

#endregion

def normalizar_atributos_excel(excel_path: str, region: str, output_path: str = None):
    """
    Normaliza la columna 'Atributos' de la hoja 'Products' según la región.
    - Si region != 'es': reemplaza las keys usando MAP_REGION y elimina las que no estén en el mapping.
    - Si region == 'es': mantiene todas las claves (sólo renombra si hay equivalencia).

    :param excel_path: ruta del Excel original
    :param region: 'es', 'it', 'uk', etc.
    :param output_path: ruta opcional para guardar un nuevo Excel
    :return: DataFrame modificado
    """
    df = pd.read_excel(excel_path, sheet_name=region.upper())

    if "Atributos" not in df.columns:
        raise ValueError("❌ La hoja 'Products' no contiene la columna 'Atributos'.")

    region = region.lower()
    mapping = ATTRIBUTE_MAPS.get(region)
    if mapping is None:
        raise ValueError(f"❌ No existe un mapping para la región: {region}")

    def normalize_dict(attr_str):
        try:
            attrs = ast.literal_eval(attr_str) if isinstance(attr_str, str) else {}
            normalized = {}
            for k, v in attrs.items():
                std_key = mapping.get(k)
                if std_key:  # hay equivalencia → usar la clave española
                    normalized[std_key] = v
                else:
                    if region == "es":
                        # España mantiene aunque no esté en el mapping
                        normalized[k] = v
                    # si region != "es" y no hay equivalencia → se descarta
            return normalized
        except Exception:
            return {}

    # Normalizar columna
    df["Atributos"] = df["Atributos"].apply(normalize_dict)

    # Guardar si se pide
    if output_path:
        with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, index=False, sheet_name=region.upper())
        print(f"✅ Archivo normalizado guardado en: {output_path}")

def preparar_excel(ruta_archivo):
    # Leer el archivo Excel
    with pd.ExcelFile(ruta_archivo) as xls:
        sheet_names = xls.sheet_names  # Obtener los nombres de las hojas

    # Si la primera hoja no se llama 'Products', la renombramos
    if sheet_names[0] != 'Products':
        print(f"Cambiando el nombre de la primera hoja a 'Products'.")
        df_products = pd.read_excel(ruta_archivo, sheet_name=sheet_names[0])
        with pd.ExcelWriter(ruta_archivo, engine='openpyxl') as writer:
            df_products.to_excel(writer, sheet_name='Products', index=False)
            # Escribimos las otras hojas, si existen
            for sheet in sheet_names[1:]:
                df = pd.read_excel(ruta_archivo, sheet_name=sheet)
                df.to_excel(writer, sheet_name=sheet, index=False)

        # Ahora, recargamos el archivo después de renombrar la primera hoja
        with pd.ExcelFile(ruta_archivo) as xls:
            sheet_names = xls.sheet_names

    # Verificar si hay al menos dos hojas adicionales y que se llamen 'Update' y 'Quantity'
    if len(sheet_names) < 4:
        with pd.ExcelWriter(ruta_archivo, engine='openpyxl', mode='a') as writer:
            # Si no existe la hoja 'Update', la creamos con un DataFrame vacío
            if 'Update' not in sheet_names:
                print("Creando la hoja 'Update'.")
                df_update = pd.DataFrame()
                df_update.to_excel(writer, sheet_name='Update', index=False)

            # Si no existe la hoja 'Quantity', la creamos con un DataFrame vacío
            if 'Quantity' not in sheet_names:
                print("Creando la hoja 'Quantity'.")
                df_quantity = pd.DataFrame()
                df_quantity.to_excel(writer, sheet_name='Quantity', index=False)

            # Si no existe la hoja 'Import', la creamos con un DataFrame vacío
            if 'Import' not in sheet_names:
                print("Creando la hoja 'Import'.")
                df_import = pd.DataFrame()
                df_import.to_excel(writer, sheet_name='Import', index=False)

    elif 'Update' not in sheet_names:
        # Si no existe la hoja 'Update', la creamos
        with pd.ExcelWriter(ruta_archivo, engine='openpyxl', mode='a') as writer:
            print("Renombrando o creando la hoja 'Update'.")
            df_update = pd.DataFrame()
            df_update.to_excel(writer, sheet_name='Update', index=False)

    elif 'Quantity' not in sheet_names:
        # Si no existe la hoja 'Quantity', la creamos
        with pd.ExcelWriter(ruta_archivo, engine='openpyxl', mode='a') as writer:
            print("Renombrando o creando la hoja 'Quantity'.")
            df_quantity = pd.DataFrame()
            df_quantity.to_excel(writer, sheet_name='Quantity', index=False)

    elif 'Import' not in sheet_names:
        # Si no existe la hoja 'Import', la creamos
        with pd.ExcelWriter(ruta_archivo, engine='openpyxl', mode='a') as writer:
            print("Renombrando o creando la hoja 'Import'.")
            df_import = pd.DataFrame()
            df_import.to_excel(writer, sheet_name='Import', index=False)

def seleccionar_excel_():
    from tkinter import filedialog
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel original",
        filetypes=[("Archivos Excel", "*.xlsx *.xls *.pdf")]
    )

    if archivo:
        return archivo
    return None

def save_to_excel_(items, filename, region):
    import os
    from datetime import datetime
    filename += "-" + datetime.now().strftime('%d-%m-%Y') + ".xlsx"

    if not items:
        print("No hay datos para guardar en Excel.")
        return

    df_main = pd.DataFrame(items)

    ruta = os.path.expanduser("~/Documents/SMI Files")
    os.makedirs(ruta, exist_ok=True)
    ruta += "/" + filename
    df_main[ClavesExcel.REFERENCIA.value] = df_main[ClavesExcel.REFERENCIA.value].apply(lambda x: f"'{x}")

    # Convertir listas de categorías en string plano
    if ClavesExcel.CATEGORIA.value in df_main.columns:
        df_main[ClavesExcel.CATEGORIA.value] = df_main[ClavesExcel.CATEGORIA.value].apply(lambda x: ", ".join(x) if isinstance(x, list) else x)

    with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
        df_main.to_excel(writer, sheet_name=region.upper(), index=False)

def excel_read_and_parse_(filename, region):
    import string
    import os
    from datetime import datetime
    def clean_float(value):
        import re
        try:
            if value is None or str(value).strip().lower() in ['nan', '', 'none']:
                return 0.0
            value = re.sub(r'[^\d.,-]', '', str(value).strip())
            return float(value.replace(',', '.')) if value else 0.0
        except (ValueError, TypeError):
            return 0.0
    filename += "-" + datetime.now().strftime('%d-%m-%Y') + ".xlsx"
    ruta = os.path.expanduser("~/Documents/SMI Files")
    df = pd.read_excel(ruta + "/" + filename, sheet_name=region.upper())
    df[ClavesExcel.PESO.value] = df[ClavesExcel.PESO.value].apply(clean_float)

    barcode_column = ClavesExcel.REFERENCIA.value
    df[barcode_column] = df[barcode_column].astype(str)
    barcode_count = {}

    for index, row in df.iterrows():
        original_barcode = row[barcode_column]
        if original_barcode in barcode_count and original_barcode != "'":
            count = barcode_count[original_barcode]
            suffix = string.ascii_uppercase[count - 1]
            new_barcode = f"{original_barcode}{suffix}"
            while new_barcode in barcode_count:
                count += 1
                suffix = string.ascii_uppercase[count]
                new_barcode = f"{original_barcode}{suffix}"
            df.at[index, barcode_column] = new_barcode.lstrip("'")
            barcode_count[original_barcode] += 1
            barcode_count[new_barcode] = 1
        else:
            barcode_count[original_barcode] = 1
            df.at[index, barcode_column] = original_barcode.lstrip("'")

    # Guardamos resultados
    final_path = ruta + "/" + filename
    df.to_excel(final_path, sheet_name=region.upper(), index=False)
    merge_path = os.path.expanduser("~/Documents/SMI Files/Merge")
    os.makedirs(merge_path, exist_ok=True)
    excel_path = merge_path + "/scraped.xlsx"

    if not os.path.exists(excel_path):
        df.to_excel(excel_path, index=False, sheet_name=region.upper())
    else:
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, index=False, sheet_name=region.upper())

    normalizar_atributos_excel(excel_path, region, excel_path)
    print(f"✅ El archivo Excel está listo: {filename}")

