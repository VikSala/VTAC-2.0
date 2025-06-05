# services/excel_diff_handler.py
import pandas as pd
import json
import os
import html
import ast
import datetime


class ExcelDiffHandler:
    def __init__(self, excel_path_old, excel_path_new, selections_path, html_output_path, event_manager=None):
        self.excel_path_old = excel_path_old
        self.excel_path_new = excel_path_new
        self.selections_path = selections_path
        self.html_output_path = html_output_path
        self.event_manager = event_manager

        self.data_old = pd.DataFrame()
        self.data_new = pd.DataFrame()

        self.comparison_results = {}
        self.user_selections = {}

        self.product_identifier_col_config = 'default_code'
        self.product_identifier_col = self.product_identifier_col_config

        self.display_name_col_config = 'Name'
        self.display_name_col = self.display_name_col_config
        self.display_name_col_old = self.display_name_col_config

        self.structured_spec_field = 'Specifications'
        self.image_urls_field = 'Image_Urls'
        self.website_description_field = 'website_description'
        self.pdf_urls_field = 'Pdf_Urls'

    def _emit_progress(self, message):
        if self.event_manager:
            self.event_manager.emit('progress_update', f"[ExcelDiffHandler] {message}")
        else:
            print(f"[ExcelDiffHandler] {message}")

    def _load_excels(self):
        # Cargar Excel Antiguo
        self._emit_progress(f"Cargando Excel antiguo: {self.excel_path_old}")
        if not os.path.exists(self.excel_path_old):
            self._emit_progress(
                f"Advertencia: Archivo Excel antiguo no encontrado: {self.excel_path_old}. Se asumirá que no hay datos antiguos.")
            self.data_old = pd.DataFrame()
        else:
            try:
                self.data_old = pd.read_excel(self.excel_path_old, sheet_name=0, dtype=str).fillna('')
            except Exception as e:
                raise ValueError(f"Error al leer Excel antiguo '{self.excel_path_old}': {e}")

        self.identifier_for_old_df = None
        if not self.data_old.empty:
            if self.product_identifier_col_config in self.data_old.columns:
                self.identifier_for_old_df = self.product_identifier_col_config
            elif 'Name' in self.data_old.columns:
                self.identifier_for_old_df = 'Name'
                self._emit_progress(
                    f"'{self.product_identifier_col_config}' no encontrado en Excel antiguo, usando 'Name' como identificador para el antiguo.")
            else:
                self._emit_progress(
                    f"Advertencia: Ni '{self.product_identifier_col_config}' ni 'Name' encontrados en Excel antiguo. No se podrá indexar para comparación detallada.")

            if self.display_name_col_config in self.data_old.columns:
                self.display_name_col_old = self.display_name_col_config
            elif 'Name' in self.data_old.columns:
                self.display_name_col_old = 'Name'
            elif self.identifier_for_old_df:
                self.display_name_col_old = self.identifier_for_old_df
            elif len(self.data_old.columns) > 0:
                self.display_name_col_old = self.data_old.columns[0]
            else:
                self.display_name_col_old = "Producto (Antiguo)"

        # Cargar Excel Nuevo
        self._emit_progress(f"Cargando Excel nuevo: {self.excel_path_new}")
        if not os.path.exists(self.excel_path_new):
            raise FileNotFoundError(f"Archivo Excel nuevo no encontrado: {self.excel_path_new}")
        try:
            self.data_new = pd.read_excel(self.excel_path_new, sheet_name=0, dtype=str).fillna('')
        except Exception as e:
            raise ValueError(f"Error al leer Excel nuevo '{self.excel_path_new}': {e}")

        if self.product_identifier_col_config in self.data_new.columns:
            self.product_identifier_col = self.product_identifier_col_config
        elif 'Name' in self.data_new.columns:
            self.product_identifier_col = 'Name'
            self._emit_progress(
                f"'{self.product_identifier_col_config}' no encontrado en Excel nuevo, usando 'Name' como identificador principal para la comparación.")
        else:
            raise ValueError(
                f"Columna identificador crítica ('{self.product_identifier_col_config}' o 'Name') no encontrada en Excel nuevo. No se puede continuar.")

        if self.display_name_col_config in self.data_new.columns:
            self.display_name_col = self.display_name_col_config
        elif 'Name' in self.data_new.columns:
            self.display_name_col = 'Name'
        elif self.product_identifier_col in self.data_new.columns:
            self.display_name_col = self.product_identifier_col
        elif len(self.data_new.columns) > 0:
            self.display_name_col = self.data_new.columns[0]
        else:
            self.display_name_col = "Producto (Nuevo)"

        self._emit_progress(
            f"Identificador principal para la comparación (del Excel nuevo): '{self.product_identifier_col}'.")
        self._emit_progress(f"Nombre a mostrar principal (del Excel nuevo): '{self.display_name_col}'.")
        if self.identifier_for_old_df:
            self._emit_progress(f"Identificador para Excel antiguo: '{self.identifier_for_old_df}'.")
            self._emit_progress(f"Nombre a mostrar para Excel antiguo: '{self.display_name_col_old}'.")

        self._emit_progress("Excels cargados.")

    def _parse_field_to_dict(self, value_str):
        if not value_str or not isinstance(value_str, str): return {}
        try:
            loaded_json = json.loads(value_str)
            if isinstance(loaded_json, dict): return loaded_json
            if isinstance(loaded_json, list) and len(loaded_json) == 1 and isinstance(loaded_json[0], dict): return \
            loaded_json[0]
            return {}
        except json.JSONDecodeError:
            try:
                evaluated = ast.literal_eval(value_str)
                if isinstance(evaluated, dict): return evaluated
                if isinstance(evaluated, list) and len(evaluated) == 1 and isinstance(evaluated[0], dict): return \
                evaluated[0]
                return {}
            except:
                return {}

    def _parse_urls_from_string(self, urls_str):
        if not urls_str or not isinstance(urls_str, str): return []
        try:
            evaluated = ast.literal_eval(urls_str)
            if isinstance(evaluated, list):
                return [str(url).strip() for url in evaluated if
                        isinstance(url, str) and url.strip().startswith('http')]
        except:
            pass
        urls = []
        potential_urls_comma = urls_str.split(',')
        cleaned_urls_comma = [url.strip().strip("'").strip('"') for url in potential_urls_comma]
        for p_url in cleaned_urls_comma:
            if p_url.startswith('http') and p_url not in urls: urls.append(p_url)
        if not urls and 'http' in urls_str:
            potential_urls_space = urls_str.split()
            cleaned_urls_space = [url.strip().strip("'").strip('"') for url in potential_urls_space]
            for p_url_space in cleaned_urls_space:
                if p_url_space.startswith('http') and p_url_space not in urls: urls.append(p_url_space)
        return urls

    def _load_user_selections(self):
        if self.selections_path and os.path.exists(self.selections_path):
            self._emit_progress(f"Cargando selecciones previas desde: {self.selections_path}")
            try:
                with open(self.selections_path, 'r', encoding='utf-8') as f:
                    self.user_selections = json.load(f)
                self._emit_progress(f"Selecciones de usuario previas cargadas ({len(self.user_selections)}).")
            except Exception as e:
                self._emit_progress(f"Advertencia: Error al cargar selecciones previas ({e}). Se continuará sin ellas.")
                self.user_selections = {}
        else:
            self.user_selections = {}

    def _compare_data(self):
        self._emit_progress("Comparando datos entre los dos Excels...")
        dict_old = {}
        if not self.data_old.empty and self.identifier_for_old_df:
            try:
                dict_old = self.data_old.set_index(self.identifier_for_old_df).to_dict(orient='index')
            except KeyError:
                self._emit_progress(
                    f"Error al indexar Excel antiguo con '{self.identifier_for_old_df}'. Se procederá sin datos antiguos indexados."); dict_old = {}
        elif not self.data_old.empty:
            self._emit_progress(
                "Excel antiguo no tiene un identificador claro. No se podrá comparar producto a producto con el antiguo.")
        dict_new = self.data_new.set_index(self.product_identifier_col).to_dict(orient='index')
        all_ids = set(dict_old.keys()) | set(dict_new.keys())
        all_columns_set_old = set(self.data_old.columns) if not self.data_old.empty else set()
        all_columns_set_new = set(self.data_new.columns)
        all_columns_set = all_columns_set_old | all_columns_set_new

        columns_to_compare_set = all_columns_set.copy()
        # No quitar self.display_name_col de columns_to_compare_set si queremos que sea un campo comparable
        # Solo quitar los identificadores si no son también el display_name_col
        if self.product_identifier_col in columns_to_compare_set and self.product_identifier_col != self.display_name_col:
            columns_to_compare_set.remove(self.product_identifier_col)
        if self.identifier_for_old_df and self.identifier_for_old_df != self.product_identifier_col and \
                self.identifier_for_old_df in columns_to_compare_set and self.identifier_for_old_df != self.display_name_col_old:
            columns_to_compare_set.remove(self.identifier_for_old_df)

        ordered_columns_to_compare = sorted(list(columns_to_compare_set))

        for prod_id in sorted(list(all_ids)):
            prod_data_old_row = dict_old.get(prod_id, {})
            prod_data_new_row = dict_new.get(prod_id, {})
            display_name_val_new = prod_data_new_row.get(self.display_name_col, prod_id if prod_data_new_row else '')
            display_name_val_old = prod_data_old_row.get(self.display_name_col_old,
                                                         prod_id if prod_data_old_row else '')
            display_name = display_name_val_new if prod_data_new_row else display_name_val_old
            if not display_name: display_name = prod_id
            is_new_product = not bool(prod_data_old_row) and bool(prod_data_new_row)
            is_removed_product = bool(prod_data_old_row) and not bool(prod_data_new_row)
            product_comparison_entry = {'_displayName': display_name, '_is_new': is_new_product,
                                        '_is_removed': is_removed_product}
            has_any_relevant_change_for_product_display = False

            # Comparamos explícitamente el display_name_col si no es el mismo que el product_identifier_col
            if self.display_name_col != self.product_identifier_col:
                name_val_old = str(prod_data_old_row.get(self.display_name_col_old, ''))  # Usar display_name_col_old
                name_val_new = str(prod_data_new_row.get(self.display_name_col, ''))  # Usar display_name_col
                name_has_diff = name_val_old != name_val_new
                if name_has_diff: has_any_relevant_change_for_product_display = True
                product_comparison_entry[self.display_name_col] = {  # Usar self.display_name_col como clave
                    'val_old': name_val_old, 'val_new': name_val_new,
                    'has_difference': name_has_diff,
                    'chosen_source': self.user_selections.get(str(prod_id), {}).get(self.display_name_col),
                }

            for col_name in ordered_columns_to_compare:
                # Si col_name es el display_name_col y ya lo procesamos, saltar
                if col_name == self.display_name_col and self.display_name_col != self.product_identifier_col:
                    continue

                val_old_str = str(prod_data_old_row.get(col_name, ''))
                val_new_str = str(prod_data_new_row.get(col_name, ''))
                field_comparison_details = {
                    'val_old': val_old_str, 'val_new': val_new_str, 'has_difference': False,
                    'chosen_source': self.user_selections.get(str(prod_id), {}).get(col_name),
                }
                if col_name == self.structured_spec_field:
                    specs_old_dict, specs_new_dict = self._parse_field_to_dict(val_old_str), self._parse_field_to_dict(
                        val_new_str)
                    field_comparison_details.update(
                        {'val_old_parsed_dict': specs_old_dict, 'val_new_parsed_dict': specs_new_dict})
                    all_sub_keys, any_sub_spec_diff, sub_differences_map = set(specs_old_dict.keys()) | set(
                        specs_new_dict.keys()), False, {}
                    for sub_key in sorted(list(all_sub_keys)):
                        sub_val_old, sub_val_new = str(specs_old_dict.get(sub_key, '')), str(
                            specs_new_dict.get(sub_key, ''))
                        sub_is_different = sub_val_old != sub_val_new
                        if sub_is_different: any_sub_spec_diff = True
                        sub_differences_map[sub_key] = {'val_old': sub_val_old, 'val_new': sub_val_new,
                                                        'is_different': sub_is_different}
                    field_comparison_details.update(
                        {'sub_differences': sub_differences_map, 'has_difference': any_sub_spec_diff})
                elif col_name == self.image_urls_field:
                    urls_old_list, urls_new_list = self._parse_urls_from_string(
                        val_old_str), self._parse_urls_from_string(val_new_str)
                    set_old, set_new = set(urls_old_list), set(urls_new_list)
                    images_added, images_removed = sorted(list(set_new - set_old)), sorted(list(set_old - set_new))
                    is_different_urls = bool(images_added or images_removed)
                    field_comparison_details.update({'val_old_parsed': urls_old_list, 'val_new_parsed': urls_new_list,
                                                     'images_added': images_added, 'images_removed': images_removed,
                                                     'has_difference': is_different_urls})
                elif col_name == self.pdf_urls_field:
                    urls_old_list, urls_new_list = self._parse_urls_from_string(
                        val_old_str), self._parse_urls_from_string(val_new_str)
                    is_different_urls = set(urls_old_list) != set(urls_new_list)
                    field_comparison_details.update(
                        {'has_difference': is_different_urls, 'val_old_parsed': urls_old_list,
                         'val_new_parsed': urls_new_list})
                else:
                    field_comparison_details['has_difference'] = val_old_str != val_new_str
                if field_comparison_details['has_difference']: has_any_relevant_change_for_product_display = True
                product_comparison_entry[col_name] = field_comparison_details
            if has_any_relevant_change_for_product_display or is_new_product or is_removed_product:
                self.comparison_results[str(prod_id)] = product_comparison_entry
        self._emit_progress("Comparación de datos finalizada.")

    def get_complete_css(self):
        # ... (sin cambios) ...
        return """
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding:0; background-color: #f4f7f6; color: #333; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 0.9em; box-shadow: 0 1px 3px rgba(0,0,0,0.1); background-color: white; }
        th, td { border: 1px solid #ddd; padding: 8px 10px; text-align: left; vertical-align: top; }
        th { background-color: #f0f0f0; color: #333; font-weight: 600; }
        .sku-page { display: none; border: 1px solid #ccc; border-radius: .25rem; padding: 15px; margin: 0 15px 20px 15px; background-color: #fff; }
        .sku-page.active { display: block; }
        #sku-navigation-container { padding: 10px 0; background-color: #343a40; color:white; position: sticky; top: 0; z-index: 1001; width:100%;}
        #sku-navigation { text-align: center; }
        #sku-navigation button { padding: 8px 15px; font-size: 1em; margin: 0 10px; cursor: pointer; background-color: #FF8C00; color:white; border:none; border-radius:4px; }
        #sku-navigation button:disabled { opacity: 0.6; }
        #sku-nav-status { font-weight: bold; margin: 0 15px; }
        .sku-header { background-color: #FF8C00; color: white; padding: 10px 15px; margin-bottom:15px; border-radius: 3px 3px 0 0; font-size: 1.15em; }
        .sku-header.new-product { background-color: #28a745; }
        .sku-header.removed-product { background-color: #dc3545; }
        .field-name { font-weight: bold; width: 20%; background-color: #f9f9f9; }
        .selection-cell { width: 15%; text-align: center; }
        .value-cell { width: 32.5%; }
        .spec-list-container { max-height: 200px; overflow-y: auto; padding: 5px; }
        .spec-sub-item { display: block; margin-bottom: 4px; font-size: 0.95em; }
        .spec-sub-item .key { font-style: italic; color: #555; }
        .spec-sub-item .value { margin-left: 5px; }
        .specifications-table { width: 100%; margin-top: 5px; border: 1px solid #e0e0e0; border-collapse: separate; border-spacing:0; }
        .specifications-table th { background-color: #f7f7f7; font-size: 0.9em; padding: 6px; border-bottom: 1px solid #ddd; text-align:left; }
        .specifications-table td { font-size: 0.9em; padding: 6px; border-bottom: 1px solid #eee; text-align:left; }
        .specifications-table tr:last-child td { border-bottom: none; }
        .specifications-table .spec-key-selection { font-style: normal; color: #333; width: 60%;} 
        .specifications-table .spec-key-selection small { display:block; font-size:0.85em; color:#666; margin-top:2px;}
        .specifications-table .spec-choice-selection { width: 40%; text-align:center; } 
        .spec-choice input[type="radio"] { margin-right: 3px; vertical-align:middle; }
        .spec-choice label { margin-right: 8px; font-size:0.95em; vertical-align:middle; }
        .specifications-table tr.spec-row-diff td { background-color: #fffacd; }
        .image-list { display: flex; flex-wrap: wrap; gap: 8px; padding: 5px 0; align-items: flex-start; }
        .image-list-item { display: flex; flex-direction: column; align-items: center; border: 2px solid #ddd; padding: 5px; border-radius: 4px; background-color: #fdfdfd;}
        .image-list-item img { max-width: 100px; max-height: 100px; object-fit: contain; margin-bottom: 5px; }
        .image-list-item.added { border-color: #28a745; box-shadow: 0 0 4px 1px #28a745; }
        .image-list-item.removed { border-color: #dc3545; box-shadow: 0 0 4px 1px #dc3545; }
        .image-list-item.removed img { opacity: 0.5; } 
        .website-description-content { max-height: 200px; overflow-y: auto; padding: 8px; border: 1px solid #eee; background-color: #fdfdfd; border-radius:3px; }
        .website-description-content ul, .website-description-content ol { margin-left: 20px; padding-left: 15px; }
        .website-description-content li { margin-bottom: 5px; }
        .pdf-list a { display: inline-block; margin-bottom: 5px; color: #007bff; text-decoration: none; padding: 3px 5px; border: 1px solid transparent; border-radius:3px; }
        .pdf-list a:hover { text-decoration: underline; background-color: #eef; border-color: #cce;}
        #save-button-container { position: fixed; bottom: 0; left: 0; width: 100%; background-color: #343a40; padding: 12px 0; text-align: center; z-index: 1000; box-shadow: 0 -1px 4px rgba(0,0,0,0.15);}
        #save-button { padding: 10px 25px; font-size: 1.05em; background-color: #28a745; color: white; border: none; border-radius: 5px; cursor: pointer; }
        #status-message { text-align: center; margin-top: 6px; font-weight: bold; height:18px; color: #fff; font-size:0.9em;}
        h1#main-title {text-align:center; margin: 15px 0; color: #333;}
        .no-change-text { color: #777; font-style: italic; font-size:0.9em; }
        """

    def _generate_html_for_diff(self):
        # ... (código igual hasta la preparación de product_rows_html_parts) ...
        self._emit_progress(f"Generando archivo HTML en: {self.html_output_path}")
        if not self.comparison_results:
            html_content = "<!DOCTYPE html><html><head><title>Comparación</title></head><body><p>No se encontraron productos con diferencias, ni nuevos/eliminados para mostrar.</p></body></html>"
            with open(self.html_output_path, "w", encoding="utf-8") as f: f.write(html_content)
            return

        html_parts = [f"""
        <!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Comparación de Excels Merge</title><style>
        {self.get_complete_css()}
        </style></head><body><h1 id="main-title">Comparación de Excels Merge</h1>
        <div id="save-button-container"><button id="save-button" onclick="saveAllSelections()">Guardar Selecciones y Generar Reporte</button><div id="status-message"></div></div>
        <div id="sku-navigation-container"><div id="sku-navigation"><button id="prev-sku-btn" onclick="navigateSku(-1)" disabled>Anterior</button><span id="sku-nav-status"></span><button id="next-sku-btn" onclick="navigateSku(1)">Siguiente</button></div></div>
        <div id="skus-content-container" style="margin-bottom: 70px; margin-top: 10px;">
        """]
        actual_products_rendered = 0
        for prod_id_str in sorted(self.comparison_results.keys()):
            product_data = self.comparison_results[prod_id_str]
            display_name_header = product_data.get('_displayName', prod_id_str)  # Nombre para el H2
            is_new_product = product_data.get('_is_new', False)
            is_removed_product = product_data.get('_is_removed', False)

            if not (is_new_product or is_removed_product or any(
                    v.get('has_difference', False) for k, v in product_data.items() if not k.startswith('_'))):
                continue

            header_class = "new-product" if is_new_product else "removed-product" if is_removed_product else ""
            status_prefix = "NUEVO" if is_new_product else "ELIMINADO" if is_removed_product else "Modificado"

            product_rows_html_parts = []

            # Ordenar campos para la tabla HTML: 'Name' primero, luego el resto
            field_keys_for_table = [k for k in product_data.keys() if not k.startswith('_')]

            # self.display_name_col es el nombre real de la columna que actúa como "Name"
            # No mostrarlo como un campo si es el mismo que el identificador, ya que el ID está en el header
            name_field_to_prioritize = self.display_name_col if self.display_name_col != self.product_identifier_col else None

            ordered_fields_for_html_table = []
            if name_field_to_prioritize and name_field_to_prioritize in field_keys_for_table:
                name_field_details = product_data[name_field_to_prioritize]
                if is_new_product or is_removed_product or name_field_details.get('has_difference', False):
                    ordered_fields_for_html_table.append(name_field_to_prioritize)

            for key in sorted(field_keys_for_table):
                if key not in ordered_fields_for_html_table:  # Evitar duplicados y añadir el resto
                    field_details_temp = product_data[key]
                    if is_new_product or is_removed_product or field_details_temp.get('has_difference', False):
                        ordered_fields_for_html_table.append(key)

            for field_name in ordered_fields_for_html_table:
                field_details = product_data[field_name]
                # ... (resto de la lógica para generar la fila HTML, que es la misma que tenías) ...
                # Asegúrate de que esta lógica interna para val_old, val_new y la celda de selección
                # maneje correctamente el `field_name` que ahora es `self.display_name_col`
                # cuando ese es el caso. Ya debería hacerlo si `field_details` es correcto para `self.display_name_col`.

                current_field_row_parts = []
                escaped_field_name = html.escape(field_name)
                safe_js_field_name = ''.join(c if c.isalnum() else '_' for c in field_name)
                current_field_row_parts.append(f"<tr><td class='field-name'>{escaped_field_name}</td>")
                # VALOR ANTIGUO
                current_field_row_parts.append("<td class='value-cell val-old'>")
                if is_new_product:
                    current_field_row_parts.append("<span class='no-change-text'>N/A (Nuevo Producto)</span>")
                else:
                    if field_name == self.image_urls_field:
                        current_field_row_parts.append("<div class='image-list'>")
                        for img_url in field_details.get('val_old_parsed', []): current_field_row_parts.append(
                            f"<div class='image-list-item'><img src='{html.escape(img_url)}' alt='Img'></div>")
                        if not field_details.get('val_old_parsed'): current_field_row_parts.append(
                            "<span class='no-change-text'>-</span>")
                        current_field_row_parts.append("</div>")
                    elif field_name == self.structured_spec_field:
                        current_field_row_parts.append("<div class='spec-list-container'>")
                        specs_dict = field_details.get('val_old_parsed_dict', {})
                        if not specs_dict: current_field_row_parts.append("<span class='no-change-text'>-</span>")
                        for sk, sv in sorted(specs_dict.items()): current_field_row_parts.append(
                            f"<div class='spec-sub-item'><span class='key'>{html.escape(sk)}:</span> <span class='value'>{html.escape(str(sv))}</span></div>")
                        current_field_row_parts.append("</div>")
                    elif field_name == self.pdf_urls_field:
                        current_field_row_parts.append("<div class='pdf-list'>")
                        for pdf_url in field_details.get('val_old_parsed', []): current_field_row_parts.append(
                            f"<a href='{html.escape(pdf_url)}' target='_blank'>{html.escape(os.path.basename(pdf_url.split('?')[0]))}</a>")
                        if not field_details.get('val_old_parsed'): current_field_row_parts.append(
                            "<span class='no-change-text'>-</span>")
                        current_field_row_parts.append("</div>")
                    elif field_name == self.website_description_field:
                        current_field_row_parts.append(
                            f"<div class='website-description-content'>{field_details.get('val_old', '-')}</div>")
                    else:
                        current_field_row_parts.append(html.escape(field_details.get('val_old', '-')))
                current_field_row_parts.append("</td>")
                # VALOR NUEVO
                current_field_row_parts.append("<td class='value-cell val-new'>")
                if is_removed_product:
                    current_field_row_parts.append("<span class='no-change-text'>N/A (Eliminado)</span>")
                else:
                    if field_name == self.image_urls_field:
                        current_field_row_parts.append("<div class='image-list'>")
                        val_new_parsed_list = field_details.get('val_new_parsed', [])
                        for img_url in val_new_parsed_list:
                            img_class = "image-list-item added" if img_url in field_details.get('images_added',
                                                                                                []) else "image-list-item"
                            current_field_row_parts.append(
                                f"<div class='{img_class}'><img src='{html.escape(img_url)}' alt='Img Nueva'></div>")
                        if not is_new_product:
                            for img_url in field_details.get('images_removed', []): current_field_row_parts.append(
                                f"<div class='image-list-item removed'><img src='{html.escape(img_url)}' alt='Img Eliminada'></div>")
                        if not val_new_parsed_list and (is_new_product or not field_details.get('images_removed',
                                                                                                [])): current_field_row_parts.append(
                            "<span class='no-change-text'>-</span>")
                        current_field_row_parts.append("</div>")
                    elif field_name == self.structured_spec_field:
                        current_field_row_parts.append("<div class='spec-list-container'>")
                        specs_dict = field_details.get('val_new_parsed_dict', {})
                        if not specs_dict: current_field_row_parts.append("<span class='no-change-text'>-</span>")
                        for sk, sv in sorted(specs_dict.items()): current_field_row_parts.append(
                            f"<div class='spec-sub-item'><span class='key'>{html.escape(sk)}:</span> <span class='value'>{html.escape(str(sv))}</span></div>")
                        current_field_row_parts.append("</div>")
                    elif field_name == self.pdf_urls_field:
                        current_field_row_parts.append("<div class='pdf-list'>")
                        for pdf_url in field_details.get('val_new_parsed', []): current_field_row_parts.append(
                            f"<a href='{html.escape(pdf_url)}' target='_blank'>{html.escape(os.path.basename(pdf_url.split('?')[0]))}</a>")
                        if not field_details.get('val_new_parsed'): current_field_row_parts.append(
                            "<span class='no-change-text'>-</span>")
                        current_field_row_parts.append("</div>")
                    elif field_name == self.website_description_field:
                        current_field_row_parts.append(
                            f"<div class='website-description-content'>{field_details.get('val_new', '-')}</div>")
                    else:
                        current_field_row_parts.append(html.escape(field_details.get('val_new', '-')))
                current_field_row_parts.append("</td>")
                # CELDA DE SELECCIÓN
                current_field_row_parts.append("<td class='selection-cell'>")
                if is_removed_product or is_new_product or not field_details.get('has_difference'):
                    current_field_row_parts.append(
                        f"<span class='no-change-text'>({'Eliminado' if is_removed_product else ('Nuevo' if is_new_product else 'sin cambios')})</span>")
                elif field_name == self.structured_spec_field:
                    current_field_row_parts.append(
                        f"<table class='specifications-table' id='spec-table-{html.escape(prod_id_str)}-{safe_js_field_name}'><thead><tr><th class='spec-key-selection'>Subcategoría (Valores)</th><th class='spec-choice-selection'>Elegir</th></tr></thead><tbody>")
                    raw_user_choice_for_specs, user_choice_for_field_dict = field_details.get('chosen_source'), {}
                    if isinstance(raw_user_choice_for_specs,
                                  dict): user_choice_for_field_dict = raw_user_choice_for_specs
                    sub_diffs_to_show_in_selection = {k: v for k, v in field_details.get('sub_differences', {}).items()
                                                      if v['is_different']}
                    if not sub_diffs_to_show_in_selection:
                        current_field_row_parts.append(
                            "<tr><td colspan='2' class='no-change-text'>(Sin dif. específicas)</td></tr>")
                    else:
                        for sub_key, sub_vals in sorted(sub_diffs_to_show_in_selection.items()):
                            safe_sub_key = ''.join(c if c.isalnum() else '_' for c in sub_key)
                            chosen_sub_src = user_choice_for_field_dict.get(sub_key, 'new')
                            cb_old_id, cb_new_id, radio_name = f"spec-{html.escape(prod_id_str)}-{safe_js_field_name}-{safe_sub_key}-old", f"spec-{html.escape(prod_id_str)}-{safe_js_field_name}-{safe_sub_key}-new", f"spec-radio-{html.escape(prod_id_str)}-{safe_js_field_name}-{safe_sub_key}"
                            current_field_row_parts.append(
                                f"<tr class='spec-row-diff'><td class='spec-key-selection'>{html.escape(sub_key)}<small>Ant: {html.escape(sub_vals['val_old'])}<br>Nue: {html.escape(sub_vals['val_new'])}</small></td>")
                            current_field_row_parts.append("<td class='spec-choice-selection'>")
                            current_field_row_parts.append(
                                f"<input type='radio' name='{radio_name}' id='{cb_old_id}' value='old' {'checked' if chosen_sub_src == 'old' else ''} data-sku='{html.escape(prod_id_str)}' data-field='{escaped_field_name}' data-subkey='{html.escape(sub_key)}' onchange='recordSelection(this)'><label for='{cb_old_id}'>A</label>&nbsp;")
                            current_field_row_parts.append(
                                f"<input type='radio' name='{radio_name}' id='{cb_new_id}' value='new' {'checked' if chosen_sub_src == 'new' else ''} data-sku='{html.escape(prod_id_str)}' data-field='{escaped_field_name}' data-subkey='{html.escape(sub_key)}' onchange='recordSelection(this)'><label for='{cb_new_id}'>N</label>")
                            current_field_row_parts.append("</td></tr>")
                    current_field_row_parts.append("</tbody></table>")
                elif field_name == self.image_urls_field or field_name == self.pdf_urls_field:
                    chosen_set_src = field_details.get('chosen_source', 'new')
                    radio_old_id, radio_new_id, radio_group = f"radio-set-{html.escape(prod_id_str)}-{safe_js_field_name}-old", f"radio-set-{html.escape(prod_id_str)}-{safe_js_field_name}-new", f"radio-set-group-{html.escape(prod_id_str)}-{safe_js_field_name}"
                    current_field_row_parts.append(
                        f"<input type='radio' name='{radio_group}' id='{radio_old_id}' value='old' {'checked' if chosen_set_src == 'old' else ''} data-sku='{html.escape(prod_id_str)}' data-field='{escaped_field_name}' onchange='recordSelection(this)'><label for='{radio_old_id}'>Usar Antiguo</label><br>")
                    current_field_row_parts.append(
                        f"<input type='radio' name='{radio_group}' id='{radio_new_id}' value='new' {'checked' if chosen_set_src == 'new' else ''} data-sku='{html.escape(prod_id_str)}' data-field='{escaped_field_name}' onchange='recordSelection(this)'><label for='{radio_new_id}'>Usar Nuevo</label>")
                else:
                    chosen_src = field_details.get('chosen_source', 'new')
                    radio_old_id, radio_new_id, radio_group_name = f"radio-{html.escape(prod_id_str)}-{safe_js_field_name}-old", f"radio-{html.escape(prod_id_str)}-{safe_js_field_name}-new", f"radio-group-{html.escape(prod_id_str)}-{safe_js_field_name}"
                    current_field_row_parts.append(
                        f"<input type='radio' name='{radio_group_name}' id='{radio_old_id}' value='old' {'checked' if chosen_src == 'old' else ''} data-sku='{html.escape(prod_id_str)}' data-field='{escaped_field_name}' onchange='recordSelection(this)'><label for='{radio_old_id}'>Antiguo</label><br>")
                    current_field_row_parts.append(
                        f"<input type='radio' name='{radio_group_name}' id='{radio_new_id}' value='new' {'checked' if chosen_src == 'new' else ''} data-sku='{html.escape(prod_id_str)}' data-field='{escaped_field_name}' onchange='recordSelection(this)'><label for='{radio_new_id}'>Nuevo</label>")
                current_field_row_parts.append("</td></tr>")
                product_rows_html_parts.append("".join(current_field_row_parts))

            if product_rows_html_parts:
                active_class_for_div = 'active' if actual_products_rendered == 0 else ''
                html_parts.append(
                    f"<div class='sku-page {active_class_for_div}' id='sku-page-{html.escape(prod_id_str)}'>")
                html_parts.append(
                    f"<h2 class='sku-header {header_class}'>{status_prefix} - {html.escape(display_name_header)} (ID: {html.escape(prod_id_str)})</h2><table><thead><tr><th class='field-name'>Campo</th><th class='value-cell'>Valor Antiguo</th><th class='value-cell'>Valor Nuevo</th><th class='selection-cell'>Selección</th></tr></thead><tbody>")
                html_parts.extend(product_rows_html_parts)
                html_parts.append("</tbody></table></div>")
                actual_products_rendered += 1
            elif not product_rows_html_parts and (is_new_product or is_removed_product):
                active_class_for_div = 'active' if actual_products_rendered == 0 else ''
                html_parts.append(
                    f"<div class='sku-page {active_class_for_div}' id='sku-page-{html.escape(prod_id_str)}'>")
                html_parts.append(
                    f"<h2 class='sku-header {header_class}'>{status_prefix} - {html.escape(display_name_header)} (ID: {html.escape(prod_id_str)})</h2>")
                html_parts.append(
                    f"<p style='padding:10px;'>Este producto es {status_prefix.lower()} y no tiene campos detallados para mostrar/comparar.</p></div>")
                actual_products_rendered += 1
        # ... (resto del JS y cierre del HTML sin cambios) ...
        initial_selections_json = json.dumps(self.user_selections)
        script_js = f"""
            <script>
                const selections = JSON.parse('{initial_selections_json}'); 
                const structuredSpecFieldJS = "{self.structured_spec_field}";
                const imageUrlsFieldJS = "{self.image_urls_field}";
                const pdfUrlsFieldJS = "{self.pdf_urls_field}";
                let currentSkuIndex = 0; 
                let skuPages = [];       
                let prevSkuBtn, nextSkuBtn, skuNavStatus;

                function recordSelection(element) {{
                    const sku = element.dataset.sku; 
                    const field = element.dataset.field;
                    const subkey = element.dataset.subkey; 

                    if (!selections[sku]) selections[sku] = {{}};

                    if (field === structuredSpecFieldJS && subkey) {{ 
                        if (typeof selections[sku][field] !== 'object' || selections[sku][field] === null) {{
                            selections[sku][field] = {{}}; 
                        }}
                        if (element.type === 'radio' && element.checked) {{
                            selections[sku][field][subkey] = element.value; 
                        }}
                    }} else if (field === imageUrlsFieldJS || field === pdfUrlsFieldJS) {{ 
                        if (element.type === 'radio' && element.checked) {{
                            selections[sku][field] = element.value; 
                        }}
                    }}
                    else if (element.type === 'radio' && element.checked) {{ 
                        selections[sku][field] = element.value; 
                    }}
                }}

                document.addEventListener('DOMContentLoaded', () => {{
                    skuPages = Array.from(document.querySelectorAll('.sku-page'));
                    prevSkuBtn = document.getElementById('prev-sku-btn');
                    nextSkuBtn = document.getElementById('next-sku-btn');
                    skuNavStatus = document.getElementById('sku-nav-status');
                    initializeSkuPagination();
                }});

                function initializeSkuPagination() {{
                    if (skuPages.length > 0) {{
                        showSkuPage(0);
                        if (skuPages.length === 1) {{ 
                            if (prevSkuBtn) prevSkuBtn.style.display = 'none';
                            if (nextSkuBtn) nextSkuBtn.style.display = 'none';
                            if (skuNavStatus) skuNavStatus.textContent = "Producto 1 de 1";
                        }} else {{
                            if (prevSkuBtn) prevSkuBtn.style.display = 'inline-block'; 
                            if (nextSkuBtn) nextSkuBtn.style.display = 'inline-block';
                        }}
                    }} else {{ 
                        const navContainer = document.getElementById('sku-navigation-container');
                        if (navContainer) navContainer.style.display = 'none';
                        const skusContentContainer = document.getElementById('skus-content-container');
                        if (skusContentContainer) {{
                            skusContentContainer.innerHTML = "<p style='text-align:center; font-size:1.2em; margin-top:30px;'>No se encontraron productos con diferencias que mostrar, ni productos nuevos o eliminados.</p>";
                        }}
                         const saveButtonContainer = document.getElementById('save-button-container');
                         if(saveButtonContainer) saveButtonContainer.style.display = 'none'; 
                    }}
                }}
                function showSkuPage(index) {{
                    if (index < 0 || index >= skuPages.length || skuPages.length === 0) return;
                    const activePage = document.querySelector('.sku-page.active'); 
                    if(activePage) activePage.classList.remove('active');

                    skuPages[index].classList.add('active'); 
                    currentSkuIndex = index; 
                    updateNavigationControls();
                    if (skuPages[currentSkuIndex]) {{ 
                        const targetElement = skuPages[currentSkuIndex];
                        const navContainer = document.getElementById('sku-navigation-container');
                        const navHeight = navContainer ? navContainer.offsetHeight : 0;
                        const elementPosition = targetElement.getBoundingClientRect().top + window.pageYOffset;
                        const offsetPosition = elementPosition - navHeight - 15; 

                        window.scrollTo({{ top: offsetPosition, behavior: 'smooth' }});
                    }}
                }}
                function navigateSku(direction) {{
                    const newIndex = currentSkuIndex + direction;
                    showSkuPage(newIndex);
                }}
                function updateNavigationControls() {{
                    if (!skuPages || skuPages.length === 0) {{ 
                        if(prevSkuBtn) prevSkuBtn.disabled = true;
                        if(nextSkuBtn) nextSkuBtn.disabled = true;
                        if(skuNavStatus) skuNavStatus.textContent = "Producto 0 de 0";
                        return;
                    }}
                    if (prevSkuBtn) prevSkuBtn.disabled = currentSkuIndex === 0;
                    if (nextSkuBtn) nextSkuBtn.disabled = currentSkuIndex >= skuPages.length - 1;
                    if (skuNavStatus) {{
                        skuNavStatus.textContent = "Producto " + (currentSkuIndex + 1) + " de " + skuPages.length;
                    }}
                }}
                async function saveAllSelections() {{
                    const saveButton = document.getElementById('save-button'); const statusDiv = document.getElementById('status-message');
                    saveButton.disabled = true; saveButton.textContent = 'Guardando...'; statusDiv.textContent = 'Enviando...'; statusDiv.style.color = '#ffc107';
                    try {{
                        const response = await fetch('http://localhost:5051/save_selections', {{ method: 'POST', headers: {{ 'Content-Type': 'application/json' }}, body: JSON.stringify(selections) }});
                        const result = await response.json();
                        if (response.ok && result.status === 'success') {{ statusDiv.textContent = '¡Guardado con éxito! El reporte se generará.'; statusDiv.style.color = '#28a745'; saveButton.textContent = 'Guardado'; }}
                        else {{ statusDiv.textContent = 'Error: ' + (result.message || 'Desconocido'); statusDiv.style.color = '#dc3545'; saveButton.disabled = false; saveButton.textContent = 'Reintentar Guardar'; }}
                    }} catch (error) {{ statusDiv.textContent = 'Error de red: ' + error.message; statusDiv.style.color = '#dc3545'; saveButton.disabled = false; saveButton.textContent = 'Reintentar Guardar'; }}
                }}
            </script>
        """
        html_parts.append(f"</div>{script_js}</body></html>")
        final_html_content = "".join(html_parts)
        try:
            with open(self.html_output_path, "w", encoding="utf-8") as f:
                f.write(final_html_content)
            self._emit_progress(f"Archivo HTML para diff generado en: {self.html_output_path}")
        except Exception as e:
            self._emit_progress(f"Error al escribir archivo HTML: {e}"); raise

    def process_and_generate_html_for_diff(self):
        # ... (sin cambios) ...
        self._emit_progress("Iniciando proceso de comparación de Excels y generación HTML.")
        self._load_excels()
        self._load_user_selections()
        self._compare_data()
        self._generate_html_for_diff()
        self._emit_progress("Proceso de comparación de Excels y generación HTML completado.")

    def generate_diff_report_excel(self, output_excel_path):
        self._emit_progress(f"Generando reporte Excel de diferencias en: {output_excel_path}")
        self._load_user_selections()
        products_sheet_data, update_sheet_data, added_products_names, removed_products_names, import_sheet_data = [], [], [], [], []
        import_columns_order = ['Image', 'Name', 'standard_price', 'Image_Urls', 'website_description',
                                'Specifications', 'Additional Information', 'default_code', 'barcode', 'weight',
                                'Website Published', 'description_sale']
        products_columns_order = list(self.data_new.columns) if not self.data_new.empty else import_columns_order
        dict_old_indexed, dict_new_indexed = {}, {}
        if not self.data_old.empty and self.identifier_for_old_df:
            try:
                dict_old_indexed = self.data_old.set_index(self.identifier_for_old_df).to_dict(orient='index')
            except KeyError:
                self._emit_progress(f"Error al indexar Excel antiguo con '{self.identifier_for_old_df}' para reporte.")
        if not self.data_new.empty and self.product_identifier_col:
            try:
                dict_new_indexed = self.data_new.set_index(self.product_identifier_col).to_dict(orient='index')
            except KeyError:
                self._emit_progress(
                    f"Error crítico al indexar Excel nuevo con '{self.product_identifier_col}' para reporte.")

        for prod_id, comparison_details_for_prod in self.comparison_results.items():
            prod_name_for_report = comparison_details_for_prod.get('_displayName', prod_id)
            is_new_prod, is_removed_prod = comparison_details_for_prod.get('_is_new',
                                                                           False), comparison_details_for_prod.get(
                '_is_removed', False)
            user_choices_for_sku = self.user_selections.get(str(prod_id), {})
            current_product_row_for_products_sheet = {}

            if is_new_prod:
                added_products_names.append(prod_name_for_report)
                new_prod_data_from_new_excel = dict_new_indexed.get(prod_id, {})
                import_row = {}
                for col_template_name in import_columns_order:
                    if col_template_name == 'default_code':
                        import_row[col_template_name] = prod_id
                    elif col_template_name == 'Name':
                        import_row[col_template_name] = prod_name_for_report
                    else:
                        import_row[col_template_name] = new_prod_data_from_new_excel.get(col_template_name, '')
                import_sheet_data.append(import_row)
                current_product_row_for_products_sheet = {col: new_prod_data_from_new_excel.get(col, '') for col in
                                                          products_columns_order}
                current_product_row_for_products_sheet[self.product_identifier_col] = prod_id  # Asegurar el ID correcto
                current_product_row_for_products_sheet[
                    self.display_name_col] = prod_name_for_report  # Asegurar el Nombre correcto
            elif is_removed_prod:
                removed_products_names.append(prod_name_for_report)
                old_prod_data_from_old_excel = dict_old_indexed.get(prod_id, {})
                current_product_row_for_products_sheet = {col: old_prod_data_from_old_excel.get(col, '') for col in
                                                          products_columns_order}
                id_col_to_use = self.identifier_for_old_df if self.identifier_for_old_df else self.product_identifier_col
                name_col_to_use = self.display_name_col_old if self.display_name_col_old else self.display_name_col
                current_product_row_for_products_sheet[id_col_to_use] = prod_id
                current_product_row_for_products_sheet[name_col_to_use] = prod_name_for_report
                for col in products_columns_order:
                    if col not in current_product_row_for_products_sheet: current_product_row_for_products_sheet[
                        col] = ''
            else:  # Producto modificado
                base_data_for_modified = dict_new_indexed.get(prod_id, {}).copy()
                final_product_data_for_products_sheet = {self.product_identifier_col: prod_id}
                chosen_name_value = prod_name_for_report
                name_field_comp_details = comparison_details_for_prod.get(self.display_name_col)
                if name_field_comp_details and name_field_comp_details.get('has_difference', False):
                    user_choice_for_name = user_choices_for_sku.get(self.display_name_col, 'new')
                    chosen_name_value = name_field_comp_details['val_new'] if user_choice_for_name == 'new' else \
                    name_field_comp_details['val_old']
                final_product_data_for_products_sheet[self.display_name_col] = chosen_name_value

                for field_name in products_columns_order:
                    if field_name == self.product_identifier_col or field_name == self.display_name_col: continue
                    field_comp_details = comparison_details_for_prod.get(field_name)
                    if field_comp_details and field_comp_details.get('has_difference', False):
                        user_choice_for_field = user_choices_for_sku.get(field_name)
                        if field_name == self.structured_spec_field:
                            chosen_value_dict, original_specs_old, original_specs_new = {}, self._parse_field_to_dict(
                                field_comp_details['val_old']), self._parse_field_to_dict(field_comp_details['val_new'])
                            all_sub_keys, spec_sub_choices = set(original_specs_old.keys()) | set(
                                original_specs_new.keys()), user_choice_for_field if isinstance(user_choice_for_field,
                                                                                                dict) else {}
                            for sub_key in all_sub_keys:
                                chosen_sub_version = spec_sub_choices.get(sub_key)
                                if chosen_sub_version == 'old':
                                    chosen_value_dict[sub_key] = original_specs_old.get(sub_key, '')
                                elif chosen_sub_version == 'new':
                                    chosen_value_dict[sub_key] = original_specs_new.get(sub_key, '')
                                else:
                                    general_field_choice = user_choices_for_sku.get(field_name,
                                                                                    'new') if not isinstance(
                                        user_choice_for_field, dict) else 'new'
                                    chosen_value_dict[sub_key] = original_specs_old.get(sub_key,
                                                                                        '') if general_field_choice == 'old' else original_specs_new.get(
                                        sub_key, '')
                            final_product_data_for_products_sheet[field_name] = json.dumps(chosen_value_dict,
                                                                                           ensure_ascii=False,
                                                                                           sort_keys=True)
                        elif field_name == self.image_urls_field or field_name == self.pdf_urls_field:
                            chosen_set_source = user_choice_for_field if user_choice_for_field in ['old',
                                                                                                   'new'] else 'new'
                            final_product_data_for_products_sheet[field_name] = json.dumps(field_comp_details.get(
                                'val_old_parsed' if chosen_set_source == 'old' else 'val_new_parsed', []),
                                                                                           ensure_ascii=False)
                        else:
                            chosen_version = user_choice_for_field if user_choice_for_field in ['old', 'new'] else 'new'
                            final_product_data_for_products_sheet[field_name] = field_comp_details[
                                'val_old' if chosen_version == 'old' else 'val_new']
                    else:
                        final_product_data_for_products_sheet[field_name] = base_data_for_modified.get(field_name, '')
                current_product_row_for_products_sheet = final_product_data_for_products_sheet

                updated_fields_for_report, not_chosen_fields_for_report, has_updates_for_this_row_in_update_sheet = {}, {}, False
                for field_name_update, field_data_from_comparison_update in comparison_details_for_prod.items():
                    if field_name_update.startswith('_') or not field_data_from_comparison_update.get('has_difference',
                                                                                                      False): continue
                    has_updates_for_this_row_in_update_sheet = True
                    user_choice_for_field_update = user_choices_for_sku.get(field_name_update)
                    if field_name_update == self.structured_spec_field:
                        chosen_specs_sub_values_update, not_chosen_specs_sub_values_update, spec_sub_choices_update = {}, {}, user_choice_for_field_update if isinstance(
                            user_choice_for_field_update, dict) else {}
                        for sub_key, sub_diff_details in field_data_from_comparison_update.get('sub_differences',
                                                                                               {}).items():
                            if not sub_diff_details['is_different']: continue
                            chosen_version_for_subkey_update = spec_sub_choices_update.get(sub_key, 'new')
                            if chosen_version_for_subkey_update == 'new':
                                chosen_specs_sub_values_update[sub_key], not_chosen_specs_sub_values_update[sub_key] = \
                                sub_diff_details['val_new'], sub_diff_details['val_old']
                            else:
                                chosen_specs_sub_values_update[sub_key], not_chosen_specs_sub_values_update[sub_key] = \
                                sub_diff_details['val_old'], sub_diff_details['val_new']
                        if chosen_specs_sub_values_update: updated_fields_for_report[field_name_update] = json.dumps(
                            chosen_specs_sub_values_update, ensure_ascii=False, sort_keys=True)
                        if not_chosen_specs_sub_values_update: not_chosen_fields_for_report[
                            field_name_update] = json.dumps(not_chosen_specs_sub_values_update, ensure_ascii=False,
                                                            sort_keys=True)
                    elif field_name_update == self.image_urls_field or field_name_update == self.pdf_urls_field:
                        chosen_set_source_update = user_choice_for_field_update if user_choice_for_field_update in [
                            'old', 'new'] else 'new'
                        val_old_parsed_update, val_new_parsed_update = field_data_from_comparison_update.get(
                            'val_old_parsed', []), field_data_from_comparison_update.get('val_new_parsed', [])
                        if chosen_set_source_update == 'new':
                            updated_fields_for_report[field_name_update], not_chosen_fields_for_report[
                                field_name_update] = json.dumps(val_new_parsed_update, ensure_ascii=False), json.dumps(
                                val_old_parsed_update, ensure_ascii=False)
                        else:
                            updated_fields_for_report[field_name_update], not_chosen_fields_for_report[
                                field_name_update] = json.dumps(val_old_parsed_update, ensure_ascii=False), json.dumps(
                                val_new_parsed_update, ensure_ascii=False)
                    else:
                        chosen_version_update = user_choice_for_field_update if user_choice_for_field_update in ['old',
                                                                                                                 'new'] else 'new'
                        if chosen_version_update == 'new':
                            updated_fields_for_report[field_name_update], not_chosen_fields_for_report[
                                field_name_update] = field_data_from_comparison_update['val_new'], \
                            field_data_from_comparison_update['val_old']
                        else:
                            updated_fields_for_report[field_name_update], not_chosen_fields_for_report[
                                field_name_update] = field_data_from_comparison_update['val_old'], \
                            field_data_from_comparison_update['val_new']
                if has_updates_for_this_row_in_update_sheet and updated_fields_for_report:
                    update_sheet_data.append({'Nombre': prod_name_for_report,
                                              'Actualizar': json.dumps(updated_fields_for_report, ensure_ascii=False,
                                                                       sort_keys=True),
                                              'No Elegido': json.dumps(not_chosen_fields_for_report, ensure_ascii=False,
                                                                       sort_keys=True) if not_chosen_fields_for_report else ''})
            if current_product_row_for_products_sheet: products_sheet_data.append(
                current_product_row_for_products_sheet)

        df_products = pd.DataFrame(products_sheet_data);
        df_products = df_products.reindex(columns=products_columns_order,
                                          fill_value='') if not df_products.empty else pd.DataFrame(
            columns=products_columns_order)
        formatted_date = datetime.date.today().strftime("%d/%m/%Y");
        update_columns = ['Nombre', 'Actualizar', f'No Elegido en ({formatted_date})']
        df_update = pd.DataFrame(update_sheet_data);
        df_update = df_update.rename(columns={'No Elegido': f'No Elegido en ({formatted_date})'}).reindex(
            columns=update_columns, fill_value='') if not df_update.empty else pd.DataFrame(columns=update_columns)
        max_len = max(len(added_products_names), len(removed_products_names));
        padded_added, padded_removed = added_products_names + [''] * (
                    max_len - len(added_products_names)), removed_products_names + [''] * (
                                                   max_len - len(removed_products_names))
        df_quantity = pd.DataFrame({'Producto Añadido': padded_added, 'Producto Eliminado': padded_removed});
        df_quantity = pd.DataFrame(columns=['Producto Añadido',
                                            'Producto Eliminado']) if df_quantity.empty and not added_products_names and not removed_products_names else df_quantity
        df_import = pd.DataFrame(import_sheet_data);
        df_import = df_import.reindex(columns=import_columns_order,
                                      fill_value='') if not df_import.empty else pd.DataFrame(
            columns=import_columns_order)
        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                df_products.to_excel(writer, sheet_name='Products', index=False)
                df_update.to_excel(writer, sheet_name='Update', index=False)
                df_quantity.to_excel(writer, sheet_name='Quantity', index=False)
                df_import.to_excel(writer, sheet_name='Import', index=False)
            self._emit_progress(f"Reporte Excel guardado en: {output_excel_path}")
        except Exception as e:
            self._emit_progress(f"Error al guardar el reporte Excel: {e}"); import \
                traceback; traceback.print_exc(); raise
