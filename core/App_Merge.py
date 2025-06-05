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
    Función que contiene toda la lógica de comparación y generación del HTML,
    idéntica a lo que era _run_comparison_task en la clase Controller.
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
        #html_output_dir = os.path.join('data', 'common', 'html_comparison')
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
        start_comparison_server(event_manager, user_selections_path, port=5050)

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

def start_comparison_server(event_manager, selections_file_path: str, port: int = 5050):
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
                return jsonify({"status": "success", "message": "Selecciones guardadas."})
            except Exception as e_save:
                event_manager.emit('progress_update', f"Error guardando selecciones: {e_save}")
                return jsonify({"status": "error", "message": str(e_save)}), 500

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

def execute_final_merge_file_generation(
    event_manager,
    excel_path: str,
    selections_path: str,
    output_filename: str
):
    """
    Inicia la generación del archivo fusionado final (Excel o JSON),
    idéntico a execute_final_merge_file_generation en Controller.
    """
    print(
        f"App_Merge: execute_final_merge_file_generation con excel='{excel_path}', selections='{selections_path}', output='{output_filename}'"
    )
    if is_critical_operation_active():
        event_manager.emit('operation_error', "Ya hay una operación en curso.")
        return

    start_long_operation(run_final_merge_task, event_manager, excel_path, selections_path, output_filename)


def run_final_merge_task(
    event_manager,
    excel_path: str,
    selections_path: str,
    output_filename: str
):
    """
    Contiene toda la lógica para generar el archivo fusionado final.
    Equivale a _run_final_merge_task en Controller.
    """
    print("App_Merge: run_final_merge_task iniciada...")
    try:
        event_manager.emit('progress_update', "Iniciando generación de archivo fusionado final...")

        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Archivo Excel base no encontrado: {excel_path}")
        if not os.path.exists(selections_path):
            raise FileNotFoundError(f"Archivo de selecciones no encontrado: {selections_path}")

        final_data_list_of_dicts = []
        try:
            from services.utils_comparison import ComparisonHandler

            handler = ComparisonHandler(
                excel_path=excel_path,
                skip_skus_path=None,
                selections_path=selections_path,
                html_output_path=None,
                event_manager=event_manager
            )
            if stop_current_operation.is_set():
                return
            final_data_list_of_dicts = handler.generate_merged_data_from_selections()
        except ImportError:
            event_manager.emit('operation_error', "Error crítico: No se pudo importar ComparisonHandler.")
            raise
        except Exception as e_handler_final:
            event_manager.emit(
                'operation_error',
                f"Error con ComparisonHandler en gen. final: {str(e_handler_final)}"
            )
            raise

        # Guardar el resultado
        if not final_data_list_of_dicts:
            event_manager.emit('progress_update', "No hay datos fusionados para guardar.")
        elif output_filename.endswith(".xlsx"):
            pd.DataFrame(final_data_list_of_dicts).to_excel(output_filename, index=False)
            event_manager.emit(
                'progress_update',
                f"Archivo Excel fusionado guardado en: {output_filename}"
            )
        elif output_filename.endswith(".json"):
            with open(output_filename, 'w', encoding='utf-8') as f_out:
                json.dump(final_data_list_of_dicts, f_out, indent=2, ensure_ascii=False)
            event_manager.emit(
                'progress_update',
                f"Archivo JSON fusionado guardado en: {output_filename}"
            )
        else:
            raise ValueError("Formato de salida no soportado. Use .xlsx o .json.")

        event_manager.emit('operation_completed', f"Archivo fusionado final generado: {output_filename}")

    except FileNotFoundError as e_fnf_final:
        if not stop_current_operation.is_set():
            event_manager.emit('operation_error', f"Archivo no encontrado: {str(e_fnf_final)}")
    except Exception as e_final:
        if not stop_current_operation.is_set():
            event_manager.emit('operation_error', f"Error generando archivo final: {str(e_final)}")


# ----------------------------------------------------------------
# Lógica de “Merge → comparar dos Excels fusionados + HTML + reporte”
# ----------------------------------------------------------------

def execute_excel_merge_comparison_and_launch_html(
    event_manager,
    excel_path_old: str,
    excel_path_new: str,
    user_selections_path: str,
    output_report_filename: str
):
    """
    Inicia la comparación de dos Excels fusionados y arranca el servidor Flask
    para el HTML interactivo “diff”. Luego, si el usuario guarda, genera un reporte Excel.
    Equivale a execute_excel_merge_comparison_and_launch_html en Controller.
    """
    print(
        f"App_Merge: execute_excel_merge_comparison_and_launch_html "
        f"con excel_old='{excel_path_old}', excel_new='{excel_path_new}', "
        f"selections='{user_selections_path}', output_report='{output_report_filename}'"
    )
    if is_critical_operation_active():
        event_manager.emit('operation_error', "Ya hay otra operación en curso.")
        return

    start_long_operation(
        run_excel_merge_comparison_task,
        event_manager,
        excel_path_old,
        excel_path_new,
        user_selections_path,
        output_report_filename
    )


def run_excel_merge_comparison_task(
    event_manager,
    excel_path_old: str,
    excel_path_new: str,
    user_selections_path: str,
    output_report_filename: str
):
    """
    Equivale a _run_excel_merge_comparison_task en Controller:
    - Compara dos Excels fusionados
    - Genera HTML interactivo
    - Arranca Flask en puerto 5051 (diff)
    - Cuando el usuario POST a /save_selections, genera el reporte Excel final
    """
    print(
        f"App_Merge: run_excel_merge_comparison_task con excel_old='{excel_path_old}', "
        f"excel_new='{excel_path_new}', selections='{user_selections_path}', "
        f"output_report='{output_report_filename}'"
    )
    try:
        event_manager.emit('progress_update', "Iniciando comparación entre Excels fusionados y generación HTML...")

        # Crear carpeta para selecciones
        os.makedirs(os.path.dirname(user_selections_path), exist_ok=True)

        # Directorio de salida para el HTML
        html_output_dir = os.path.expanduser("~/Documents/SMI Files/Data/data/common/html_comparison")
        # html_output_dir = os.path.join('data', 'common', 'html_comparison')
        os.makedirs(html_output_dir, exist_ok=True)
        html_file_name = "excel_diff_comparison_interactive.html"
        html_file_path = os.path.join(html_output_dir, html_file_name)

        handler = None
        try:
            from services.excel_diff_handler import ExcelDiffHandler

            handler = ExcelDiffHandler(
                excel_path_old=excel_path_old,
                excel_path_new=excel_path_new,
                selections_path=user_selections_path,
                html_output_path=html_file_path,
                event_manager=event_manager
            )
            if stop_current_operation.is_set():
                return
            handler.process_and_generate_html_for_diff()
        except ImportError:
            event_manager.emit('operation_error', "Error crítico: No se pudo importar ExcelDiffHandler.")
            raise
        except Exception as e_handler:
            event_manager.emit('operation_error', f"Error con ExcelDiffHandler: {str(e_handler)}")
            raise

        if stop_current_operation.is_set():
            return

        # Arrancamos Flask en puerto 5051 (diff)
        event_manager.emit('progress_update', "HTML generado. Iniciando servidor local (puerto 5051 para diff)...")
        start_comparison_server(event_manager, user_selections_path, port=5051)

        time.sleep(0.5)
        if stop_current_operation.is_set():
            stop_comparison_server(event_manager)
            return

        event_manager.emit('progress_update', f"Abriendo {html_file_path} en el navegador...")
        webbrowser.open(f"file://{os.path.abspath(html_file_path)}")
        event_manager.emit(
            'progress_update',
            "Proceso de comparación iniciado. Revisa tu navegador. Cierra la pestaña del HTML después de guardar para continuar."
        )

        # Esperamos hasta que el servidor se cierre o se cancele
        while not comparison_server_stop_event.is_set() and not stop_current_operation.is_set():
            if not (comparison_server_thread and comparison_server_thread.is_alive()):
                event_manager.emit('operation_error', "El servidor Flask (diff) se detuvo inesperadamente.")
                break
            time.sleep(0.2)

        if stop_current_operation.is_set():
            event_manager.emit('progress_update', "Tarea de comparación de Excels Merge cancelada.")
            return

        if comparison_server_stop_event.is_set():
            event_manager.emit('progress_update', "Selecciones guardadas. Procediendo a generar el reporte Excel...")
            if handler:
                handler.generate_diff_report_excel(output_report_filename)
                event_manager.emit(
                    'operation_completed',
                    f"Reporte de comparación de Excels Merge generado: {output_report_filename}"
                )
            else:
                event_manager.emit('operation_error', "No se pudo generar el reporte: el manejador no estaba disponible.")
        else:
            event_manager.emit(
                'progress_update',
                "La comparación de Excels Merge no se completó con guardado de selecciones. No se generó el reporte."
            )

    except FileNotFoundError as e_fnf:
        if not stop_current_operation.is_set():
            event_manager.emit('operation_error', f"Archivo no encontrado: {str(e_fnf)}")
    except Exception as e:
        import traceback
        traceback.print_exc()
        if not stop_current_operation.is_set():
            event_manager.emit('operation_error', f"Error en tarea de comparación de Excels Merge: {str(e)}")

