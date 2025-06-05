import pandas as pd

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
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )

    if archivo:
        return archivo
    return None

def save_to_excel_(items, filename):
    import os
    from datetime import datetime
    filename += "-" + datetime.now().strftime('%d-%m-%Y') + ".xlsx"

    if not items:
        print("No hay datos para guardar en Excel.")
        return

    df_main = pd.DataFrame(items)
    df_second = df_main[["Name", "Image_Urls"]].copy()

    ruta = os.path.expanduser("~/Documents/SMI Files")
    os.makedirs(ruta, exist_ok=True)
    ruta += "/" + filename
    df_main["barcode"] = df_main["barcode"].apply(lambda x: f"'{x}")

    # Convertir listas de categorías en string plano
    if "x_categoria" in df_main.columns:
        df_main["x_categoria"] = df_main["x_categoria"].apply(lambda x: ", ".join(x) if isinstance(x, list) else x)

    with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
        df_main.to_excel(writer, sheet_name='Products', index=False)
        df_second.to_excel(writer, sheet_name='image bank', index=False)

def excel_read_and_parse_(filename):
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
    df = pd.read_excel(ruta + "/" + filename, sheet_name="Products")
    df['weight'] = df['weight'].apply(clean_float)

    barcode_column = 'barcode'
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

    final_path = ruta + "/" + filename
    df.to_excel(final_path, sheet_name="Products", index=False)
    merge_path = os.path.expanduser("~/Documents/SMI Files/Merge")
    os.makedirs(merge_path, exist_ok=True)
    df.to_excel(merge_path + "/scraped.xlsx", sheet_name="Products", index=False)
    print(f"✅ El archivo Excel está listo: {filename}")