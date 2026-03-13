# App_Merge.py

import os
import threading
import time
import webbrowser
import json
import pandas as pd

# -------------------------------------------------------------
# Variables globales para controlar el estado de las operaciones
# -------------------------------------------------------------
stop_current_operation = threading.Event()
operation_active = threading.Event()
comparison_server_stop_event = threading.Event()
comparison_server_thread = None


# ------------------------------------------
# Funciones de ayuda para gestión de hilos
# ------------------------------------------

def is_critical_operation_active() -> bool:
    """
    Indica si ya hay una operación larga en curso.
    Equivalente a: self.is_critical_operation_active() en el Controller.
    """
    return operation_active.is_set()


def start_long_operation(target, *args):
    """
    Arranca la función target en un hilo aparte, marcando operation_active.
    Una vez que target termine (o se cancele), limpia operation_active.
    """
    def wrapper():
        try:
            target(*args)
        finally:
            # Si hay un servidor Flask activo, forzamos su parada
            if comparison_server_stop_event.is_set() is False:
                # Si stop_current_operation se ha marcado, detenemos el servidor
                if stop_current_operation.is_set():
                    stop_comparison_server(args[0])  # args[0] = event_manager
            operation_active.clear()

    if is_critical_operation_active():
        return

    stop_current_operation.clear()
    operation_active.set()

    thread = threading.Thread(target=wrapper, daemon=True)
    thread.start()


# ----------------------------------------
# Lógica de “Merge → comparación + HTML”
# ----------------------------------------

def execute_comparison_and_launch_html(
    event_manager,
    excel_path: str,
    user_selections_path: str,
    skip_skus_path: str = None,
):
    """
    Inicia la comparación de Excel y arranca el servidor Flask para el HTML interactivo.
    Parámetros:
      - event_manager: objeto responsable de emitir eventos (progress_update, operation_error, etc.)
      - excel_path: ruta al archivo Excel base
      - user_selections_path: ruta donde guardaremos el JSON de selecciones
      - skip_skus_path: (opcional) ruta al JSON con SKUs a omitir
    """
    print(f"App_Merge: execute_comparison_and_launch_html con excel='{excel_path}', selections='{user_selections_path}'")

    if is_critical_operation_active():
        event_manager.emit('operation_error', "Ya hay otra operación en curso.")
        return

    # Arrancamos la tarea en segundo plano
    start_long_operation(run_comparison_task, event_manager, excel_path, user_selections_path, skip_skus_path)


def run_comparison_task(
    event_manager,
    excel_path: str,
    user_selections_path: str,
    skip_skus_path: str = None,
):
    """
    Función que contiene toda la lógica de comparación y generación del HTML
    """
    print(
        f"App_Merge: run_comparison_task con excel='{excel_path}', selections='{user_selections_path}', skip_skus='{skip_skus_path}'"
    )
    try:
        event_manager.emit('progress_update', "Iniciando proceso de comparación y generación HTML...")

        # Nos aseguramos de que exista la carpeta donde se guardarán las selecciones
        os.makedirs(os.path.dirname(user_selections_path), exist_ok=True)

        # Directorio de salida para el HTML
        html_output_dir = os.path.expanduser("~/Documents/SMI Files/Data/data/common/html_comparison")
        os.makedirs(html_output_dir, exist_ok=True)
        html_file_name = "comparison_interactive.html"
        html_file_path = os.path.join(html_output_dir, html_file_name)

        # Validar existencia de skip_skus_path
        current_skip_skus_path = skip_skus_path
        if current_skip_skus_path and not os.path.exists(current_skip_skus_path):
            event_manager.emit(
                'progress_update',
                f"Advertencia: Archivo SKUs a omitir '{current_skip_skus_path}' no encontrado. Se ignorará."
            )
            current_skip_skus_path = None

        # Importar y ejecutar ComparisonHandler
        try:
            from services.utils_comparison import ComparisonHandler

            handler = ComparisonHandler(
                excel_path=excel_path,
                skip_skus_path=current_skip_skus_path,
                selections_path=user_selections_path,
                html_output_path=html_file_path,
                event_manager=event_manager
            )
            if stop_current_operation.is_set():
                return
            handler.process_and_generate_html()
        except ImportError:
            event_manager.emit('operation_error', "Error crítico: No se pudo importar ComparisonHandler.")
            raise
        except Exception as e:
            event_manager.emit('operation_error', f"Error durante el procesamiento con ComparisonHandler: {str(e)}")
            raise

        if stop_current_operation.is_set():
            return

        # Una vez generado el HTML, arrancamos el servidor Flask
        event_manager.emit('progress_update', "HTML generado. Iniciando servidor local (puerto 5050)...")
        start_comparison_server(event_manager, excel_path, user_selections_path, port=5050)

        time.sleep(0.5)
        if stop_current_operation.is_set():
            stop_comparison_server(event_manager)
            return

        event_manager.emit('progress_update', f"Abriendo {html_file_path} en el navegador...")
        webbrowser.open(f"file://{os.path.abspath(html_file_path)}")
        event_manager.emit(
            'progress_update',
            "Proceso de comparación iniciado. Revisa tu navegador. Cierra la pestaña del HTML después de guardar para detener el servidor."
        )

        # Esperamos hasta que se cierre el servidor o se cancele la operación
        while not comparison_server_stop_event.is_set() and not stop_current_operation.is_set():
            if not (comparison_server_thread and comparison_server_thread.is_alive()):
                event_manager.emit(
                    'operation_error',
                    "El servidor Flask (puerto 5050) se detuvo inesperadamente."
                )
                break
            time.sleep(0.2)

        # Conclusión de la tarea
        if not stop_current_operation.is_set() and comparison_server_stop_event.is_set():
            event_manager.emit('operation_completed', "Proceso de comparación y selección finalizado.")
        elif stop_current_operation.is_set():
            event_manager.emit('progress_update', "Tarea de comparación cancelada.")

    except FileNotFoundError as e_fnf:
        if not stop_current_operation.is_set():
            event_manager.emit('operation_error', f"Archivo no encontrado: {str(e_fnf)}")
    except Exception as e:
        import traceback
        traceback.print_exc()
        if not stop_current_operation.is_set():
            event_manager.emit('operation_error', f"Error en tarea de comparación: {str(e)}")


# -------------------------------------------------
# Arranque y parada del servidor Flask de comparación
# -------------------------------------------------

def start_comparison_server(event_manager, excel, selections_file_path: str, port: int = 5050):
    """
    Crea un servidor Flask que escucha POST en /save_selections.
    Cuando llegan las selecciones, las guarda y detiene el servidor.
    Equivale a _start_comparison_server en Controller, pero sin `self`.
    """
    global comparison_server_thread, comparison_server_stop_event

    if comparison_server_thread and comparison_server_thread.is_alive():
        event_manager.emit(
            'progress_update',
            f"Advertencia: Servidor de comparación en puerto {port} ya podría estar activo."
        )
        print(f"App_Merge: Servidor de comparación en puerto {port} ya activo o conflicto.")

    try:
        from flask import Flask, request, jsonify
        from flask_cors import CORS
        import logging

        # Reducir logging de werkzeug
        log = logging.getLogger('werkzeug')
        log.setLevel(logging.ERROR)

        app = Flask(__name__)
        CORS(app)

        @app.route('/save_selections', methods=['POST'])
        def save_selections_route():
            try:
                data = request.json

                with open(selections_file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)

                event_manager.emit('progress_update', f"Selecciones guardadas en {selections_file_path}")
                print(f"App_Merge: /save_selections (puerto {port}) - Activando comparison_server_stop_event")
                comparison_server_stop_event.set()

                # 🧠 OPCIONAL: aplicar directamente las selecciones al Excel final
                try:
                    from services.utils_comparison import ComparisonHandler
                    handler = ComparisonHandler(
                        excel_path=excel,
                        skip_skus_path=None,
                        selections_path=selections_file_path,
                        html_output_path=None,
                        event_manager=event_manager
                    )
                    handler.load_user_selections()#Cambiado
                    handler.apply_user_selections_to_excel()
                except Exception as e_apply:
                    print(f"⚠ Error aplicando selecciones al Excel: {e_apply}")

                return jsonify({"status": "success", "message": "Selecciones guardadas y aplicadas."})
            except Exception as e_save:
                event_manager.emit('progress_update', f"Error guardando selecciones: {e_save}")
                return jsonify({"status": "error", "message": str(e_save)}), 500

        @app.route('/save_new_edits', methods=['POST'])
        def save_new_edits_route():
            try:
                data = request.json  # El JSON con los cambios

                # Guardamos el JSON en disco si lo deseas
                #with open("nuevos_editados.json", 'w', encoding='utf-8') as f: json.dump(data, f, indent=2, ensure_ascii=False)

                event_manager.emit('progress_update', f"Selecciones guardadas en {selections_file_path}")
                print(f"App_Merge: /save_new_edits (puerto {port}) - Activando comparison_server_stop_event")
                comparison_server_stop_event.set()

                # Reutilizamos ComparisonHandler
                from services.utils_comparison import ComparisonHandler
                handler = ComparisonHandler(
                    excel_path=excel,
                    skip_skus_path=None,
                    selections_path=selections_file_path,
                    html_output_path=None,
                    event_manager=event_manager
                )

                # Reutilizamos el métode ya hecho, pero inyectando directamente el dict
                handler.user_selections = data
                handler.apply_direct_edits_to_excel()

                return jsonify({"status": "success", "message": "Cambios guardados y aplicados al Excel."})
            except Exception as e:
                return jsonify({"status": "error", "message": str(e)}), 500

        @app.route('/process_presence', methods=['POST'])
        def process_presence():
            try:
                data = request.get_json()  # [{sku: '123', tipo: 'scrap'/'local'}, ...]
                excel_path = excel

                if not os.path.exists(excel_path):
                    return jsonify({"status": "error", "message": f"No se encontró el archivo: {excel_path}"}), 400

                xls = pd.ExcelFile(excel_path)
                sheets = {sheet: pd.read_excel(xls, sheet).fillna('') for sheet in xls.sheet_names}
                sku_col = 'SKU'
                grupo_nuevo = ['ES', 'UK', 'ITA']
                grupo_local = 'OPT'

                for item in data:
                    sku = str(item['sku'])
                    tipo = item['tipo']

                    # Verificar existencia del SKU
                    en_nuevo = None
                    for sheet_name in grupo_nuevo:
                        df = sheets.get(sheet_name)
                        if df is not None and sku in df[sku_col].astype(str).values:
                            en_nuevo = sheet_name
                            break

                    en_local = sku in sheets.get(grupo_local, pd.DataFrame())[sku_col].astype(str).values

                    if tipo == 'scrap':
                        if en_nuevo and not en_local:
                            # Caso 1: producto en NUEVO y elijo scrap -> copiar a LOCAL (OPT)
                            row = sheets[en_nuevo][sheets[en_nuevo][sku_col].astype(str) == sku]
                            sheets[grupo_local] = pd.concat([sheets.get(grupo_local, pd.DataFrame()), row],
                                                            ignore_index=True)
                        elif en_local and not en_nuevo:
                            # Caso 4: producto en LOCAL y elijo scrap -> eliminar de LOCAL
                            sheets[grupo_local] = sheets[grupo_local][sheets[grupo_local][sku_col].astype(str) != sku]

                    elif tipo == 'local':
                        if en_nuevo and not en_local:
                            # Caso 2: producto en NUEVO y elijo local -> eliminar de ES/UK/ITA
                            for sheet_name in grupo_nuevo:
                                df = sheets.get(sheet_name)
                                if df is not None:
                                    sheets[sheet_name] = df[df[sku_col].astype(str) != sku].reset_index(drop=True)
                        elif en_local and not en_nuevo:
                            # Caso 3: producto en LOCAL y elijo local -> copiar a ES, UK, ITA
                            row = sheets[grupo_local][sheets[grupo_local][sku_col].astype(str) == sku]
                            for sheet_name in grupo_nuevo:
                                sheets[sheet_name] = pd.concat([sheets.get(sheet_name, pd.DataFrame()), row],
                                                               ignore_index=True)

                # Guardar el archivo sobrescribiendo
                with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
                    for name, df in sheets.items():
                        df.to_excel(writer, sheet_name=name, index=False)

                comparison_server_stop_event.set()
                return jsonify({"status": "success"})

            except Exception as e:
                return jsonify({"status": "error", "message": str(e)}), 500

        @app.route('/process_images', methods=['POST'])
        def process_images():
            from PIL import Image
            import imagehash
            import requests
            from io import BytesIO

            try:
                data = request.json  # Espera: { sku, images_by_country }
                sku = data.get("sku")
                images = data.get("images_by_country", {})
                hash_cache = {}

                # Determinar país base
                ref_country = max(images.items(), key=lambda x: len(x[1]))[0]
                ref_imgs = images[ref_country]
                ref_hashes = {}

                def download_hash(url):
                    if url in hash_cache:
                        return hash_cache[url]
                    if "http" not in url: url = "http://143.47.53.74:8070/" + url
                    headers = {"User-Agent": "Mozilla/5.0"}
                    response = requests.get(url, headers=headers, timeout=10)
                    response.raise_for_status()
                    img = Image.open(BytesIO(response.content))
                    h = imagehash.phash(img)
                    hash_cache[url] = h
                    return h

                for url in ref_imgs:
                    ref_hashes[url] = download_hash(url)

                result = {c: [] for c in images}
                result[ref_country] = ref_imgs  # Mantenerlas todas

                for country, urls in images.items():
                    if country == ref_country:
                        continue
                    for u in urls:
                        h = download_hash(u)
                        is_dup = False
                        for ref_u, ref_h in ref_hashes.items():
                            if h - ref_h < 5:#== 0:
                                # Eliminar del otro si el ref es OPT, o si el otro no lo es
                                if ref_country == "OPT" or country != "OPT":
                                    is_dup = True
                                    break
                        if not is_dup:
                            result[country].append(u)

                return jsonify({"status": "ok", "filtered": result})
            except Exception as e:
                return jsonify({"status": "error", "message": str(e)}), 500

        def run_server():
            from werkzeug.serving import make_server
            server = make_server('localhost', port, app, threaded=True)
            event_manager.emit('progress_update', f"Servidor Flask configurado en http://localhost:{port}.")

            def shutdown_monitor():
                comparison_server_stop_event.wait()
                print(
                    f"App_Merge: Evento de parada del servidor (puerto {port}) detectado, llamando a server.shutdown()..."
                )
                try:
                    server.shutdown()
                except Exception as e_shutdown:
                    print(f"App_Merge: Error durante server.shutdown() (puerto {port}): {e_shutdown}")

            monitor_thread = threading.Thread(target=shutdown_monitor, daemon=True)
            monitor_thread.start()

            try:
                server.serve_forever()
            except Exception as e_serve:
                print(f"App_Merge: Excepción durante server.serve_forever() (puerto {port}): {e_serve}")
            finally:
                print(f"App_Merge: server.serve_forever() (puerto {port}) ha terminado.")

        comparison_server_stop_event.clear()
        comparison_server_thread = threading.Thread(
            target=run_server,
            daemon=True,
            name=f"FlaskServerThreadPort{port}"
        )
        comparison_server_thread.start()

    except ImportError:
        event_manager.emit('operation_error', "Flask no instalado.")
    except Exception as e_serv:
        event_manager.emit('operation_error', f"No se pudo iniciar servidor en puerto {port}: {e_serv}")
        stop_current_operation.set()


def stop_comparison_server(event_manager):
    """
    Forza la parada del servidor Flask (si está corriendo) y emite progreso.
    Equivale a _stop_comparison_server en Controller, pero sin `self`.
    """
    global comparison_server_thread, comparison_server_stop_event

    print("App_Merge: stop_comparison_server llamado.")

    if not comparison_server_stop_event.is_set():
        comparison_server_stop_event.set()

    if comparison_server_thread and comparison_server_thread.is_alive():
        port_name_part = comparison_server_thread.name.replace("FlaskServerThreadPort", "")
        port_name_part = port_name_part if port_name_part.isdigit() else "desconocido"
        print(
            f"App_Merge: Servidor Flask (puerto {port_name_part}, hilo {comparison_server_thread.name}) activo, esperando que termine..."
        )
        comparison_server_thread.join(timeout=5)
        if comparison_server_thread.is_alive():
            event_manager.emit(
                'progress_update',
                f"Advertencia: Servidor (puerto {port_name_part}) no se detuvo limpiamente."
            )
        else:
            event_manager.emit('progress_update', f"Servidor (puerto {port_name_part}) detenido.")
    else:
        print(
            "App_Merge: Servidor Flask no estaba activo o ya se había detenido al llamar a stop_comparison_server."
        )

    comparison_server_thread = None


# ---------------------------------------------------
# Lógica de “Merge → generar archivo final (Excel/JSON)”
# ---------------------------------------------------

def preparar_excel_opt(event_manager, ruta_archivo):
    from datetime import datetime
    from collections import defaultdict
    import numpy as np
    import json

    def clean_value(val):
        """Convierte valores numpy a tipos nativos de Python."""
        if isinstance(val, (np.generic, pd._libs.tslibs.nattype.NaTType)):
            return val.item()
        elif pd.isna(val):
            return None
        return val

    def comparar_dicts_json(json_old, json_new):
        """Compara dos strings JSON y devuelve solo los pares clave-valor que han cambiado."""
        try:
            d1 = json.loads(json_old) if isinstance(json_old, str) else {}
            d2 = json.loads(json_new) if isinstance(json_new, str) else {}
        except Exception:
            return {}, {}
        cambios_actualizar = {}
        cambios_no_elegido = {}
        for k in d1.keys() & d2.keys():
            if d1[k] != d2[k]:
                cambios_actualizar[k] = d2[k]
                cambios_no_elegido[k] = d1[k]
        return cambios_actualizar, cambios_no_elegido

    event_manager.emit('progress_update', "Iniciando generación de archivo fusionado final...")
    hoy = datetime.now().strftime("%d/%m/%Y")

    df_old = pd.read_excel(ruta_archivo, sheet_name="OLD")
    df_new = pd.read_excel(ruta_archivo, sheet_name="NEW")

    df_old.columns = df_old.columns.str.strip()
    df_new.columns = df_new.columns.str.strip()

    clave = 'SKU' if 'SKU' in df_new.columns else df_new.columns[0]

    diferencias = []
    for _, old_row in df_old.iterrows():
        sku = old_row[clave]
        new_row = df_new[df_new[clave] == sku]
        if not new_row.empty:
            new_row = new_row.iloc[0]
            for col in df_old.columns:
                if col == clave or col not in df_new.columns:
                    continue
                val_old = old_row[col]
                val_new = new_row[col]

                if pd.isna(val_old) and pd.isna(val_new):
                    continue
                elif pd.isna(val_old) != pd.isna(val_new) or val_old != val_new:
                    diferencias.append((sku, col, val_old, val_new))

    cambios_dict = defaultdict(lambda: {'Actualizar': {}, f'No elegido en ({hoy})': {}})

    for sku, col, val_old, val_new in diferencias:
        if col == 'Atributos':
            cambios_act, cambios_no = comparar_dicts_json(val_old, val_new)
            if cambios_act:  # Solo si hay cambios dentro del JSON
                cambios_dict[sku]['Actualizar'][col] = json.dumps(cambios_act, ensure_ascii=False)
                cambios_dict[sku][f'No elegido en ({hoy})'][col] = json.dumps(cambios_no, ensure_ascii=False)
        else:
            cambios_dict[sku]['Actualizar'][col] = clean_value(val_new)
            cambios_dict[sku][f'No elegido en ({hoy})'][col] = clean_value(val_old)

    df_update = pd.DataFrame([
        {
            'SKU': sku,
            'Actualizar': cambios['Actualizar'],
            f'No elegido en ({hoy})': cambios[f'No elegido en ({hoy})']
        }
        for sku, cambios in cambios_dict.items()
    ])

    # Detectar productos añadidos y eliminados
    old_skus = set(df_old[clave])
    new_skus = set(df_new[clave])
    nuevos_skus = list(new_skus - old_skus)
    eliminados_skus = list(old_skus - new_skus)

    df_quantity = pd.DataFrame({
        'Producto Añadido': nuevos_skus + [None] * max(len(eliminados_skus) - len(nuevos_skus), 0),
        'Producto Eliminado': eliminados_skus + [None] * max(len(nuevos_skus) - len(eliminados_skus), 0)
    })

    df_import = df_new[df_new[clave].isin(nuevos_skus)]

    # Reescribir hoja Products
    with pd.ExcelWriter(ruta_archivo, engine='openpyxl', mode='w') as writer:
        df_new.to_excel(writer, sheet_name='Products', index=False)

    # Guardar hojas adicionales
    with pd.ExcelWriter(ruta_archivo, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_update.to_excel(writer, sheet_name='Update', index=False)
        df_quantity.to_excel(writer, sheet_name='Quantity', index=False)
        df_import.to_excel(writer, sheet_name='Import', index=False)

    print("✅ Archivo actualizado con las hojas: Products, Update, Quantity, Import.")
    event_manager.emit('operation_completed', f"Archivo fusionado final generado: {ruta_archivo}")
