from collections import namedtuple

from services.utils import Utils
import pandas as pd
import ast

is_connected = False

#region Importar Productos

def attributes_to_odoo(product_template_id, specifications, params, import_all = False):
    if isinstance(specifications, str):
        try:
            specifications = ast.literal_eval(specifications)
        except Exception as e:
            print(f"❌ Error convirtiendo especificaciones: {e}")
            specifications = {}

    # Agregar atributos
    if import_all:
        if isinstance(specifications, dict):
            for key, value in specifications.items():
                #print(f"🛠️ Key: {key} | Valor: {value}")
                attr_id = Utils.get_or_create_attribute(key, params)
                value_id = Utils.get_or_create_attribute_value(attr_id, value, params)
                Utils.create_attribute_line(product_template_id, attr_id, value_id, params)
    else:
        if isinstance(specifications, dict):
            for key, value in specifications.items():
                attr_id = Utils.get_attribute(key, params)
                value_id = Utils.get_or_create_attribute_value(attr_id, value, params)
                Utils.create_attribute_line(product_template_id, attr_id, value_id, params)

# === IMPORT: Leer Excel y crear Productos en Odoo ===
def import_to_odoo(excel_path, import_all = False):
    from App_Connection import models, db, uid, password
    from deep_translator import GoogleTranslator

    ConnectionParams = namedtuple("ConnectionParams", ["models", "db", "uid", "password"])
    params = ConnectionParams(models, db, uid, password)

    if import_all: df = pd.read_excel(excel_path, sheet_name="Products")
    else: df = pd.read_excel(excel_path, sheet_name="Import")
    total = len(df)

    if import_all:
        #print("PROBANDO IMPORT ATRIBUTOS DESDE CER0!!!!")

        #Utils.delete_all_attributes(params)

        atributos = set()

        Utils.create_attribute("Marca", params)

        for spec in df["Specifications"].dropna():
            if isinstance(spec, str):
                try:
                    spec_dict = ast.literal_eval(spec)
                    if isinstance(spec_dict, dict):
                        for key in spec_dict:
                            if key != "Marca":# QUITAR LUEGO
                                atributos.add(key)
                except Exception as e:
                    print(f"Error evaluando Specifications: {spec} → {e}")

        for atr in atributos:
            Utils.create_attribute(atr, params)

    # === Iterar filas y crear productos ===
    for index, row in df.iterrows():
        # Actualizar progreso en el label
        if total>0:
            print(f"Importando: {index + 1}/{total}")

        name = row["Name"]

        product_data = Utils.build_product_data_from_row(row, Utils, df)

        product_template_id = models.execute_kw(
            db, uid, password,
            'product.template', 'create',
            [product_data]
        )
        #print(f"✅ Producto '{name}' creado con ID {product_template_id}")

        #Atributos
        specifications = row["Specifications"]
        attributes_to_odoo(product_template_id, specifications, params, import_all)

        #Categorías
        def eliminar_todas_las_categorias(params):
            """
            Elimina todas las categorías de producto en Odoo, excepto las protegidas por defecto como 'All', 'Expenses', etc.
            """
            # Categorías que no deben eliminarse
            categorias_protegidas = ['All', 'Deliveries', 'Expenses', 'Saleable', 'Courses', 'Events', 'PoS']

            # Buscar todas las categorías que NO estén protegidas
            ids = params.models.execute_kw(params.db, params.uid, params.password,
                'product.category', 'search', [[('name', 'not in', categorias_protegidas)]])

            if not ids:
                print("✅ No hay categorías para eliminar (las protegidas se conservan).")
                return

            # Eliminar solo las no protegidas
            params.models.execute_kw(params.db, params.uid, params.password,
                'product.category', 'unlink', [ids])
        #eliminar_todas_las_categorias(params)

        if isinstance(row["x_categoria"], str) and row["x_categoria"].strip():
            row["x_categoria"] = GoogleTranslator(source='auto', target='es').translate(row["x_categoria"])
        Utils.get_or_create_categories(row["x_categoria"], params, product_template_id)

        # Agregar imágenes adicionales si existen
        extra_urls = ast.literal_eval(row["Image_Urls"])
        for i, url in enumerate(extra_urls, start=1):
            image_b64 = Utils.image_url_to_base64(url.strip())
            if image_b64:
                image_id = models.execute_kw(
                    db, uid, password,
                    'product.image', 'create',
                    [{
                        'name': f"{name}-extra-{i}.webp",
                        'product_tmpl_id': product_template_id,
                        'image_1920': image_b64,
                        'sequence': i
                    }]
                )
    #label.config(text="Productos Importados correctamente")
    print("🎉 Todos los productos han sido procesados.")

# endregion

# region Actualizar Productos
def update_from_merge_to_odoo(excel_path):
    """
    Actualiza todos los productos a partir de los datos en la hoja 'Update' del Excel.
    """
    from App_Connection import models, db, uid, password

    ConnectionParams = namedtuple("ConnectionParams", ["models", "db", "uid", "password"])
    params = ConnectionParams(models, db, uid, password)

    df = pd.read_excel(excel_path, sheet_name="Update")
    if df.empty or df.dropna(how='all').empty:
        return
    # Limpiar filas vacías o sin datos válidos
    df = df.dropna(subset=["Name", "Actualizar"])
    total = len(df)

    for index, row in df.iterrows():
        if total > 0:
            print(f"Actualizando: {index + 1}/{total}")

        nombre = row["Name"]
        campos_a_actualizar = row["Actualizar"]

        # Si es string, evalúa el diccionario (en caso de que Excel lo haya guardado como texto)
        if isinstance(campos_a_actualizar, str):
            try:
                campos_a_actualizar = eval(campos_a_actualizar)
            except Exception as e:
                print(f"⚠️ Error al interpretar los campos de '{nombre}': {e}")
                continue

        print(f"🔧 Actualizando producto: {nombre}")

        # Si se incluye 'Specifications', tratarlo por separado
        if 'Specifications' in campos_a_actualizar:
            product_id = models.execute_kw(
                db, uid, password,
                'product.template', 'search',
                [[['name', '=', nombre]]]
            )
            if product_id:
                attributes_to_odoo(product_id[0], campos_a_actualizar['Specifications'], params)
                campos_a_actualizar.pop('Specifications')  # Evitar actualizarlo también por write()

        # Si hay más campos, hacer una única actualización
        if campos_a_actualizar:
            actualizado = Utils.update_odoo_product(
                product_name=nombre,
                update_values=campos_a_actualizar
            )

            if actualizado:
                print(f"✅ Producto '{nombre}' actualizado correctamente.")
            else:
                print(f"❌ Fallo al actualizar '{nombre}'.")

    print("🎉 Todos los productos han sido actualizados.")

def delete_from_merge_to_odoo(excel_path):
    from App_Connection import models, db, uid, password
    # Leer la hoja 'Quantity'
    df = pd.read_excel(excel_path, sheet_name='Quantity')
    total = len(df)
    index = 0

    if df.empty or df.dropna(how='all').empty:
        return

    # Obtener nombres de productos a eliminar
    productos_a_eliminar = df['Producto Eliminado'].dropna().unique()

    for name in productos_a_eliminar:
        if total>0:
            print(f"Eliminando: {index + 1}/{total}")
            index += 1

        # Buscar producto por nombre
        product_ids = models.execute_kw(
            db, uid, password,
            'product.template', 'search',
            [[['name', '=', name]]]
        )

        if product_ids:
            # Eliminar el producto
            models.execute_kw(
                db, uid, password,
                'product.template', 'unlink',
                [product_ids]
            )
            print(f"🗑️ Producto '{name}' eliminado (IDs: {product_ids})")
        else:
            print(f"⚠️ Producto '{name}' no encontrado, no se eliminó")
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
        import shutil

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

def main_import():
    global is_connected
    import threading

    abrir_ajustes()

    if is_connected:

        ruta = Utils.seleccionar_excel()

        def run_import():
            try:
                if not ruta:
                    print("Import cancelado.")
                    return
                Utils.preparar_excel_import(ruta)
                df_import = pd.read_excel(ruta, sheet_name='Import')
                import_all = True if df_import.empty else False

                import_to_odoo(ruta, import_all)

            except Exception as e:
                print(f"❌ Error en import_to_odoo: {e}")
            else:
                threading.Thread(target=run_update).start()

        def run_update():
            try:
                update_from_merge_to_odoo(ruta)
            except Exception as e:
                print(f"❌ Error en update_from_merge_to_odoo: {e}")
            else:
                threading.Thread(target=run_delete).start()

        def run_delete():
            try:
                delete_from_merge_to_odoo(ruta)
            except Exception as e:
                print(f"❌ Error en delete_from_merge_to_odoo: {e}")

        # Inicia la cadena de ejecución con el primer hilo
        threading.Thread(target=run_import).start()
    else:
        print("Error en la conexión, no es posible hacer el Import.")