# services/utils_comparison.py
import pandas as pd
import json
import os
import html
import ast
import difflib


class ComparisonHandler:
    def __init__(self, excel_path, skip_skus_path, selections_path, html_output_path, event_manager=None):
        self.excel_path = excel_path
        self.skip_skus_path = skip_skus_path
        self.selections_path = selections_path
        self.html_output_path = html_output_path
        self.event_manager = event_manager

        self.data_es = pd.DataFrame()
        self.data_uk = pd.DataFrame()
        self.data_ita = pd.DataFrame()

        self.all_skus_to_process = []
        self.comparison_results = {}
        self.user_selections = {}
        self.all_columns_from_excel = set()
        self.sku_column_name = 'SKU'

        self.url_list_fields = ['Pdf_Urls', 'Video_Urls']
        self.structured_fields = ['specifications']

    def _emit_progress(self, message):
        if self.event_manager:
            self.event_manager.emit('progress_update', f"[ComparisonHandler] {message}")
        else:
            print(f"[ComparisonHandler] {message}")

    def _load_excel(self):
        # ... (como estaba en la última versión funcional)
        self._emit_progress(f"Leyendo datos del Excel: {self.excel_path}")
        if not os.path.exists(self.excel_path):
            raise FileNotFoundError(f"El archivo Excel no se encontró en: {self.excel_path}")
        try:
            xls = pd.ExcelFile(self.excel_path)
            sheet_names_found = []

            first_sheet_to_read_cols = None
            if 'ES' in xls.sheet_names:
                first_sheet_to_read_cols = 'ES'
            elif 'UK' in xls.sheet_names:
                first_sheet_to_read_cols = 'UK'
            elif 'ITA' in xls.sheet_names:
                first_sheet_to_read_cols = 'ITA'
            elif xls.sheet_names:
                first_sheet_to_read_cols = xls.sheet_names[0]

            if not first_sheet_to_read_cols:
                raise ValueError("El archivo Excel no contiene ninguna hoja.")

            df_cols_check = pd.read_excel(xls, first_sheet_to_read_cols, nrows=0)
            dtype_all_str = {col: str for col in df_cols_check.columns}

            for col in df_cols_check.columns:
                if col.upper() == 'SKU':
                    self.sku_column_name = col
                    break

            if 'ES' in xls.sheet_names:
                self.data_es = pd.read_excel(xls, 'ES', dtype=dtype_all_str).fillna('')
                sheet_names_found.append('ES')
                self.all_columns_from_excel.update(self.data_es.columns)
            if 'UK' in xls.sheet_names:
                self.data_uk = pd.read_excel(xls, 'UK', dtype=dtype_all_str).fillna('')
                sheet_names_found.append('UK')
                self.all_columns_from_excel.update(self.data_uk.columns)
            if 'ITA' in xls.sheet_names:
                self.data_ita = pd.read_excel(xls, 'ITA', dtype=dtype_all_str).fillna('')
                sheet_names_found.append('ITA')
                self.all_columns_from_excel.update(self.data_ita.columns)

            if not sheet_names_found:
                raise ValueError("Ninguna de las hojas esperadas (ES, UK, ITA) fue encontrada en el Excel.")
            if self.data_es.empty and self.data_uk.empty and self.data_ita.empty:
                self._emit_progress("Advertencia: Las hojas ES, UK, ITA están vacías o no contienen datos.")

            for df, name in [(self.data_es, 'ES'), (self.data_uk, 'UK'), (self.data_ita, 'ITA')]:
                if not df.empty and self.sku_column_name not in df.columns:
                    raise ValueError(f"La columna '{self.sku_column_name}' no se encontró en la hoja '{name}'.")

            if self.sku_column_name in self.all_columns_from_excel:
                self.all_columns_from_excel.remove(self.sku_column_name)

            self._emit_progress(f"Datos de Excel cargados para hojas: {', '.join(sheet_names_found)}.")
        except Exception as e:
            self._emit_progress(f"Error crítico al cargar Excel: {e}")
            raise

    def _load_skus_to_skip(self):
        # ... (como estaba en la última versión funcional)
        skipped_skus = set()
        if self.skip_skus_path and os.path.exists(self.skip_skus_path):
            self._emit_progress(f"Cargando SKUs a omitir desde: {self.skip_skus_path}")
            try:
                with open(self.skip_skus_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, dict) and "skus" in data and isinstance(data["skus"], list):
                        skipped_skus.update(map(str, data["skus"]))
                    elif isinstance(data, list):
                        skipped_skus.update(map(str, data))
                self._emit_progress(f"{len(skipped_skus)} SKUs cargados para omitir.")
            except Exception as e:
                self._emit_progress(
                    f"Advertencia: Error al cargar SKUs a omitir desde '{self.skip_skus_path}': {e}. Se continuará sin ellos.")
        return skipped_skus

    def _load_user_selections(self):
        # ... (como estaba en la última versión funcional)
        if self.selections_path and os.path.exists(self.selections_path):
            self._emit_progress(f"Cargando selecciones previas desde: {self.selections_path}")
            try:
                with open(self.selections_path, 'r', encoding='utf-8') as f:
                    self.user_selections = json.load(f)
                self._emit_progress(f"Selecciones de usuario previas cargadas ({len(self.user_selections)} SKUs).")
            except Exception as e:
                self._emit_progress(
                    f"Advertencia: Error al cargar selecciones previas desde '{self.selections_path}' (se continuará sin ellas): {e}")
                self.user_selections = {}
        else:
            self._emit_progress(f"No se encontró archivo de selecciones previas en '{self.selections_path}'.")
            self.user_selections = {}

    def _identify_all_skus_to_process(self):
        # ... (como estaba en la última versión funcional)
        skus_from_excel = set()
        id_col_name = self.sku_column_name
        if not self.data_es.empty and id_col_name in self.data_es.columns: skus_from_excel.update(
            self.data_es[id_col_name].astype(str).unique())
        if not self.data_uk.empty and id_col_name in self.data_uk.columns: skus_from_excel.update(
            self.data_uk[id_col_name].astype(str).unique())
        if not self.data_ita.empty and id_col_name in self.data_ita.columns: skus_from_excel.update(
            self.data_ita[id_col_name].astype(str).unique())

        skus_to_skip_set = self._load_skus_to_skip()

        self.all_skus_to_process = sorted(list(skus_from_excel - skus_to_skip_set))
        if not self.all_skus_to_process:
            self._emit_progress("Advertencia: No se identificaron SKUs para procesar.")
        else:
            self._emit_progress(f"Identificados {len(self.all_skus_to_process)} SKUs únicos para procesar.")

    def _parse_generic_urls(self, urls_str):
        # ... (como estaba en la última versión funcional)
        if not isinstance(urls_str, str) or not urls_str.strip():
            return []
        parsed_urls = []
        try:
            evaluated_data = ast.literal_eval(urls_str)
            if isinstance(evaluated_data, list):
                parsed_urls = [str(url).strip() for url in evaluated_data if
                               isinstance(url, str) and str(url).strip().startswith('http')]
            elif isinstance(evaluated_data, dict):
                for value in evaluated_data.values():
                    if isinstance(value, str) and value.strip().startswith('http'):
                        parsed_urls.append(value.strip())
        except (ValueError, SyntaxError, TypeError):
            if 'http' in urls_str:
                if not urls_str.strip().startswith('[') and not urls_str.strip().startswith('{'):
                    possible_urls_csv = [url.strip() for url in urls_str.split(',') if url.strip().startswith('http')]
                    if possible_urls_csv: return possible_urls_csv
                if urls_str.strip().startswith('http'): return [urls_str.strip()]
        return parsed_urls

    def _parse_specifications(self, spec_str: str) -> list:
        # ... (como estaba en la última versión funcional)
        parsed_specs = []
        if not spec_str or not isinstance(spec_str, str) or not spec_str.strip():
            return []
        try:
            evaluated_data = ast.literal_eval(spec_str)
            if isinstance(evaluated_data, dict):
                parsed_specs = [{'key': str(k).strip(), 'value': str(v).strip()} for k, v in evaluated_data.items() if
                                str(k).strip() or str(v).strip()]
            elif isinstance(evaluated_data, list):
                for idx, item in enumerate(evaluated_data):
                    item_str = str(item).strip()
                    if not item_str: continue
                    if ':' in item_str:
                        parts = item_str.split(':', 1)
                        parsed_specs.append({'key': parts[0].strip(), 'value': parts[1].strip()})
                    else:
                        parsed_specs.append({'key': f"Item_{idx + 1}", 'value': item_str})
            else:
                if str(evaluated_data).strip():
                    parsed_specs.append({'key': "Specification", 'value': str(evaluated_data).strip()})
        except (ValueError, SyntaxError, TypeError):
            lines = []
            raw_lines = []
            if '\n' in spec_str:
                raw_lines = spec_str.split('\n')
            elif ';' in spec_str:
                raw_lines = spec_str.split(';')
            elif ',' in spec_str and ':' not in spec_str:
                raw_lines = spec_str.split(',')

            if raw_lines:
                lines = [line.strip() for line in raw_lines if line.strip()]
            elif spec_str.strip():
                lines = [spec_str.strip()]

            for idx, line in enumerate(lines):
                if ':' in line:
                    parts = line.split(':', 1)
                    parsed_specs.append({'key': parts[0].strip(), 'value': parts[1].strip()})
                else:
                    parsed_specs.append({'key': f"Detail_{idx + 1}", 'value': line})
        return [spec for spec in parsed_specs if spec.get('key', '').strip() or spec.get('value', '').strip()]

    def _compare_products(self):
        # ... (Aquí va la lógica de _compare_products con la corrección para 'is_different')
        if not self.all_skus_to_process:
            self._emit_progress("No hay SKUs para comparar. Saltando la comparación.")
            return
        self._emit_progress("Iniciando comparación de productos por SKU...")
        total_skus = len(self.all_skus_to_process)
        id_col_name = self.sku_column_name

        for sku_idx, sku_val_str in enumerate(self.all_skus_to_process):
            if (sku_idx + 1) % 20 == 0 or sku_idx == 0:
                self._emit_progress(f"Comparando SKU {sku_idx + 1}/{total_skus}: {sku_val_str}")

            prod_comparison_data = {}
            row_es_series = self.data_es[self.data_es[id_col_name] == sku_val_str].iloc[0] if not self.data_es[
                self.data_es[id_col_name] == sku_val_str].empty else pd.Series(dtype=str)
            row_uk_series = self.data_uk[self.data_uk[id_col_name] == sku_val_str].iloc[0] if not self.data_uk[
                self.data_uk[id_col_name] == sku_val_str].empty else pd.Series(dtype=str)
            row_ita_series = self.data_ita[self.data_ita[id_col_name] == sku_val_str].iloc[0] if not self.data_ita[
                self.data_ita[id_col_name] == sku_val_str].empty else pd.Series(dtype=str)

            for field in sorted(list(self.all_columns_from_excel)):
                val_es_orig = str(row_es_series.get(field, ''))
                val_uk_orig = str(row_uk_series.get(field, ''))
                val_ita_orig = str(row_ita_series.get(field, ''))

                if field.lower() in self.structured_fields:
                    parsed_es = self._parse_specifications(val_es_orig)
                    parsed_uk = self._parse_specifications(val_uk_orig)
                    parsed_ita = self._parse_specifications(val_ita_orig)

                    map_es = {s['key']: s['value'] for s in parsed_es}
                    map_uk = {s['key']: s['value'] for s in parsed_uk}
                    map_ita = {s['key']: s['value'] for s in parsed_ita}

                    # Marcar 'is_different' en los items de las listas parseadas
                    for spec_list_country_idx, spec_list_country_data in enumerate(
                            [(parsed_es, "ES"), (parsed_uk, "UK"), (parsed_ita, "ITA")]):
                        current_parsed_list, current_country_code = spec_list_country_data
                        for spec_item in current_parsed_list:
                            key_to_check = spec_item['key']
                            value_to_check = spec_item['value']
                            spec_item['is_different'] = False  # Default

                            # Comparar con otras fuentes
                            has_conflict_or_uniqueness = False
                            # Fuentes que tienen CUALQUIER spec
                            sources_with_specs = []
                            if parsed_es: sources_with_specs.append(map_es)
                            if parsed_uk: sources_with_specs.append(map_uk)
                            if parsed_ita: sources_with_specs.append(map_ita)

                            # Fuentes que tienen ESTA clave
                            sources_with_this_key = []
                            if key_to_check in map_es: sources_with_this_key.append(map_es)
                            if key_to_check in map_uk: sources_with_this_key.append(map_uk)
                            if key_to_check in map_ita: sources_with_this_key.append(map_ita)

                            if len(sources_with_this_key) < len(sources_with_specs) and len(sources_with_specs) > 1:
                                # La clave no está presente en todas las fuentes que tienen specs, es una diferencia estructural
                                has_conflict_or_uniqueness = True
                            else:
                                # La clave está presente en todas las fuentes que tienen specs (o solo hay una fuente con specs)
                                # Ahora comprobar si los valores son diferentes
                                for other_map in sources_with_this_key:  # Iterar sobre los que tienen la clave
                                    if other_map[key_to_check] != value_to_check:
                                        has_conflict_or_uniqueness = True
                                        break

                            if has_conflict_or_uniqueness:
                                spec_item['is_different'] = True

                    any_sub_attr_is_different = any(
                        s.get('is_different', False) for s_list in [parsed_es, parsed_uk, parsed_ita] for s in s_list)
                    field_has_difference = any_sub_attr_is_different

                    chosen_items_list = self.user_selections.get(sku_val_str, {}).get(field, [])
                    # ... (resto de la lógica de chosen_items_list y prod_comparison_data[field] se mantiene como antes)
                    if isinstance(chosen_items_list, str) and chosen_items_list in ['ES', 'UK', 'ITA']:
                        default_source_country = chosen_items_list
                        source_parsed_list_for_migration = []
                        if default_source_country == 'ES':
                            source_parsed_list_for_migration = parsed_es
                        elif default_source_country == 'UK':
                            source_parsed_list_for_migration = parsed_uk
                        elif default_source_country == 'ITA':
                            source_parsed_list_for_migration = parsed_ita
                        chosen_items_list = [
                            {'key': s['key'], 'value': s['value'], 'source_country': default_source_country} for s in
                            source_parsed_list_for_migration]
                    elif not isinstance(chosen_items_list, list):
                        chosen_items_list = []

                    prod_comparison_data[field] = {
                        "ES_parsed": parsed_es, "UK_parsed": parsed_uk, "ITA_parsed": parsed_ita,
                        "ES": val_es_orig, "UK": val_uk_orig, "ITA": val_ita_orig,
                        "has_difference": field_has_difference,
                        "chosen_source": chosen_items_list
                    }
                elif field == 'Image_Urls' or field in self.url_list_fields:
                    # ... (como estaba en la última versión funcional)
                    list_es = self._parse_generic_urls(val_es_orig)
                    list_uk = self._parse_generic_urls(val_uk_orig)
                    list_ita = self._parse_generic_urls(val_ita_orig)
                    all_lists_tuples = [tuple(sorted(l)) for l in [list_es, list_uk, list_ita] if l]
                    has_difference_urls = len(set(all_lists_tuples)) > 1 if all_lists_tuples else False

                    chosen_val = self.user_selections.get(sku_val_str, {}).get(field)
                    if field == 'Image_Urls':
                        chosen_val = chosen_val if isinstance(chosen_val, list) else (
                            [chosen_val] if chosen_val else [])

                    prod_comparison_data[field] = {
                        "ES": list_es, "UK": list_uk, "ITA": list_ita,
                        "has_difference": has_difference_urls,
                        "chosen_source": chosen_val
                    }
                else:
                    # ... (como estaba en la última versión funcional)
                    distinct_values_present = {v for v in [val_es_orig, val_uk_orig, val_ita_orig] if v}
                    has_difference_txt = len(distinct_values_present) > 1
                    chosen_source_for_field = self.user_selections.get(sku_val_str, {}).get(field, None)
                    prod_comparison_data[field] = {
                        "ES": val_es_orig, "UK": val_uk_orig, "ITA": val_ita_orig,
                        "has_difference": has_difference_txt,
                        "chosen_source": chosen_source_for_field,
                    }
            self.comparison_results[sku_val_str] = prod_comparison_data
        self._emit_progress("Comparación de productos completada.")

    def _generate_html(self):
        # ... (CSS y HTML inicial como en la respuesta anterior) ...
        if not self.comparison_results:
            self._emit_progress("No hay resultados de comparación para generar HTML.")
            html_content = """<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Comparación de Productos</title></head><body><h1>Comparación de Productos</h1><p>No se encontraron productos para comparar o ninguno tiene diferencias.</p></body></html>"""
            with open(self.html_output_path, "w", encoding="utf-8") as f: f.write(html_content)
            self._emit_progress(f"HTML vacío generado en: {self.html_output_path}")
            return

        self._emit_progress(f"Generando archivo HTML en: {self.html_output_path}")

        html_parts = [f"""
        <!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Comparación Interactiva de Productos</title><style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding:0; background-color: #f4f7f6; color: #333; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 25px; font-size: 0.85em; box-shadow: 0 2px 5px rgba(0,0,0,0.1); background-color: white; }}
        th, td {{ border: 1px solid #ccc; padding: 10px; text-align: left; vertical-align: top; }}
        th {{ background-color: #e9ecef; color: #495057; font-weight: 600; }}
        .sku-page {{ display: none; border: 1px solid #dee2e6; border-radius: .25rem; padding: 15px; margin-bottom:20px; margin-left:15px; margin-right:15px; background-color: #fff; }}
        .sku-page.active {{ display: block; }}
        #sku-navigation-container {{ padding: 10px 0; background-color: #343a40; color:white; position: sticky; top: 0; z-index: 1001; box-shadow: 0 2px 4px rgba(0,0,0,0.1); width:100%;}}
        #sku-navigation {{ text-align: center; }}
        #sku-navigation button {{ padding: 8px 15px; font-size: 1em; margin: 0 10px; cursor: pointer; background-color: #007bff; color:white; border:none; border-radius:4px; }}
        #sku-navigation button:disabled {{ cursor: not-allowed; opacity: 0.65; background-color: #6c757d;}}
        #sku-nav-status {{ font-weight: bold; margin: 0 15px; }}

        .sku-header {{ background-color: #007bff; color: white; padding: 12px 15px; margin-top: 0; border-radius: 5px 5px 0 0; font-size: 1.2em; }}
        .field-name {{ font-weight: bold; min-width: 180px; background-color: #f8f9fa; }}
        .chosen-value-cell {{ background-color: #d1ecf1 !important; border-left: 3px solid #007bff; }}
        .country-val {{ padding: 8px; max-width: 300px; min-width:200px; word-wrap: break-word; }}
        .selection-group input[type="radio"], .selection-group input[type="checkbox"] {{ margin-right: 4px; vertical-align: middle; }}
        .selection-group label {{ margin-right: 12px; font-size: 0.95em; cursor: pointer; }}
        #save-button-container {{ position: fixed; bottom: 0; left: 0; width: 100%; background-color: #343a40; padding: 15px 0; text-align: center; box-shadow: 0 -2px 5px rgba(0,0,0,0.2); z-index: 1000; }}
        #save-button {{ padding: 12px 28px; font-size: 1.1em; background-color: #28a745; color: white; border: none; border-radius: 5px; cursor: pointer; }}
        #status-message {{ text-align: center; margin-top: 8px; font-weight: bold; height:20px; color: #fff; }}
        .image-container {{ display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }}
        .image-container img {{ max-width: 100px; max-height: 100px; border: 2px solid transparent; cursor: pointer; border-radius: 3px; }}
        .image-container img.selected-image {{ border-color: #007bff; box-shadow: 0 0 8px rgba(0,123,255,0.5); }}
        .url-list, .specifications-list ul {{ padding-left: 20px; margin:0; list-style-type: disc; }}
        .url-list li, .specifications-list li {{ margin-bottom: 4px; }}
        .url-list a {{ text-decoration: none; color: #007bff; }} .url-list a:hover {{ text-decoration: underline; }}

        .specifications-cell ul.specifications-sublist {{ list-style-type: none; padding-left: 0; margin: 0; }}
        .specifications-sublist li {{ margin-bottom: 6px; padding: 4px; border-radius: 3px; }}
        .specifications-sublist li:hover {{ background-color: #f0f8ff; }}
        .specifications-sublist input[type="checkbox"] {{ margin-right: 8px; vertical-align: middle; transform: scale(1.1); }}
        .specifications-sublist label {{ font-size: 0.9em; cursor: pointer; display: inline; }}
        .specifications-selection-display ul.specifications-chosen-list {{ list-style-type: none; padding-left: 0; margin: 5px 0 0 0; font-size: 0.9em; }}
        .specifications-chosen-list li {{ margin-bottom: 3px; background-color: #e9ecef; padding: 3px 6px; border-radius: 3px;}}
        .specifications-chosen-list em {{ font-style: normal; color: #007bff; font-weight: bold; }}
        h1#main-title {{text-align:center; margin-top: 15px; margin-bottom: 15px;}}
        </style></head><body><h1 id="main-title">Comparación Interactiva de Productos</h1>
        <div id="save-button-container"><button id="save-button" onclick="saveAllSelections()">Guardar Selecciones</button><div id="status-message"></div></div>

        <div id="sku-navigation-container">
          <div id="sku-navigation">
            <button id="prev-sku-btn" onclick="navigateSku(-1)" disabled>Anterior SKU</button>
            <span id="sku-nav-status"></span>
            <button id="next-sku-btn" onclick="navigateSku(1)">Siguiente SKU</button>
          </div>
        </div>
        <div id="skus-content-container" style="margin-bottom: 80px; margin-top: 15px;"> 
        """]

        skus_with_differences = []
        for sku_test, fields_data_test in self.comparison_results.items():
            if any(details.get('has_difference', False) for details in fields_data_test.values()):
                skus_with_differences.append(sku_test)

        sorted_skus_for_html = sorted(skus_with_differences)

        if not sorted_skus_for_html:
            html_parts.append("<!-- No SKUs with differences to display -->")

        for sku_idx_html, sku in enumerate(sorted_skus_for_html):
            fields_data = self.comparison_results[sku]
            active_class = 'active' if sku_idx_html == 0 else ''
            html_parts.append(f"<div class='sku-page {active_class}' id='sku-page-{html.escape(sku)}'>")
            html_parts.append(
                f"<h2 class='sku-header'>SKU: {html.escape(sku)}</h2><table><thead><tr><th>Campo</th><th>Valor ES</th><th>Valor UK</th><th>Valor ITA</th><th>Selección</th></tr></thead><tbody>")

            sorted_field_names = sorted(fields_data.keys())

            for field_name in sorted_field_names:
                values = fields_data[field_name]

                if not values.get('has_difference', False):
                    continue

                escaped_field_name = html.escape(field_name)
                safe_field_name_for_js = ''.join(c if c.isalnum() else '_' for c in field_name)
                unique_input_group_name = f"sel-{html.escape(sku)}-{safe_field_name_for_js}"

                html_parts.append(f"<tr><td class='field-name'>{escaped_field_name}</td>")

                if field_name == 'Image_Urls':
                    selected_image_urls = values.get('chosen_source', [])
                    if not isinstance(selected_image_urls, list): selected_image_urls = [
                        selected_image_urls] if selected_image_urls else []
                    for country_code in ["ES", "UK", "ITA"]:
                        html_parts.append(
                            f"<td class='country-val {country_code.lower()}-val image-container-cell' data-source='{country_code}'><div class='image-container' id='image-container-{html.escape(sku)}-{safe_field_name_for_js}-{country_code}'>")
                        image_urls_list_country = values[country_code]
                        if isinstance(image_urls_list_country, list):
                            for idx, img_url in enumerate(image_urls_list_country):
                                if img_url:
                                    is_img_selected = img_url in selected_image_urls
                                    img_checkbox_id = f"imgcb-{html.escape(sku)}-{safe_field_name_for_js}-{country_code}-{idx}"
                                    html_parts.append(
                                        f"<input type='checkbox' name='{unique_input_group_name}_{country_code}' value='{html.escape(img_url)}' id='{img_checkbox_id}' {'checked' if is_img_selected else ''} data-sku='{html.escape(sku)}' data-field='{escaped_field_name}' style='display:none;' onchange='recordSelection(this)'><img src='{html.escape(img_url)}' alt='Img {idx + 1}' onclick='toggleImageSelection(this, \"{img_checkbox_id}\")' class='{'selected-image' if is_img_selected else ''}'>")
                        html_parts.append("</div></td>")
                    html_parts.append(
                        f"<td class='selection-display' id='selection-display-img-{html.escape(sku)}-{safe_field_name_for_js}'>")
                    if selected_image_urls:
                        html_parts.append("Seleccionadas: ")
                        for img_url_sel in selected_image_urls: html_parts.append(
                            f"<img src='{html.escape(img_url_sel)}' style='max-width:50px; max-height:50px;'>")
                    else:
                        html_parts.append("Ninguna seleccionada")
                    html_parts.append("</td>")

                elif field_name.lower() in self.structured_fields:
                    parsed_specs_data = {"ES": values.get("ES_parsed", []), "UK": values.get("UK_parsed", []),
                                         "ITA": values.get("ITA_parsed", [])}
                    chosen_specs_from_data = values.get("chosen_source", [])

                    for country_code_spec in ["ES", "UK", "ITA"]:
                        html_parts.append(
                            f"<td class='country-val {country_code_spec.lower()}-val specifications-cell' data-source='{country_code_spec}'>")
                        country_specs_list = parsed_specs_data[country_code_spec]

                        sub_attrs_to_display = [s for s in country_specs_list if s.get('is_different', False)]

                        if sub_attrs_to_display:
                            html_parts.append("<ul class='specifications-sublist'>")
                            for idx_s, spec_item in enumerate(sub_attrs_to_display):
                                spec_key_html = html.escape(spec_item['key'])
                                spec_val_html = html.escape(spec_item['value'])

                                spec_key_for_group = spec_key_html.replace(' ', '_').replace(':', '-')
                                checkbox_group_name = f"specgroup-{html.escape(sku)}-{safe_field_name_for_js}-{spec_key_for_group}"
                                checkbox_id = f"cb-{html.escape(sku)}-{safe_field_name_for_js}-{country_code_spec}-{spec_key_for_group}-{idx_s}"

                                is_checked = any(
                                    cs['key'] == spec_item['key'] and cs['value'] == spec_item['value'] and cs[
                                        'source_country'] == country_code_spec
                                    for cs in chosen_specs_from_data
                                )
                                html_parts.append("<li>")
                                html_parts.append(f"<input type='checkbox' id='{checkbox_id}' "
                                                  f"name='{checkbox_group_name}' "  # Name para agrupar por clave
                                                  f"data-sku='{html.escape(sku)}' data-field='{escaped_field_name}' "
                                                  f"data-spec-key='{spec_key_html}' data-spec-value='{spec_val_html}' data-source-country='{country_code_spec}' "
                                                  f"{'checked' if is_checked else ''} onchange='handleSpecCheckboxChange(this)'>")  # Usar nuevo handler
                                html_parts.append(
                                    f"<label for='{checkbox_id}'><strong>{spec_key_html}:</strong> {spec_val_html}</label></li>")
                            html_parts.append("</ul>")
                        else:
                            html_parts.append("-(todos los subatributos son idénticos o no hay datos)-")
                        html_parts.append("</td>")

                    html_parts.append(
                        f"<td class='selection-display specifications-selection-display' id='selection-display-{html.escape(sku)}-{safe_field_name_for_js}'>")
                    if chosen_specs_from_data:
                        html_parts.append("Seleccionados:<ul class='specifications-chosen-list'>")
                        for cs in chosen_specs_from_data:
                            html_parts.append(
                                f"<li><strong>{html.escape(cs['key'])}:</strong> {html.escape(cs['value'])} <em>({html.escape(cs['source_country'])})</em></li>")
                        html_parts.append("</ul>")
                    else:
                        html_parts.append("Ninguno seleccionado")
                    html_parts.append("</td>")

                elif field_name in self.url_list_fields:
                    for country_code in ["ES", "UK", "ITA"]:
                        current_country_urls_list = values.get(country_code, [])
                        html_parts.append(
                            f"<td class='country-val {country_code.lower()}-val url-list-cell' data-source='{country_code}'>")
                        if current_country_urls_list:
                            html_parts.append("<ul class='url-list'>")
                            for idx, item_url in enumerate(current_country_urls_list):
                                if item_url:
                                    item_name = os.path.basename(
                                        item_url.split('?')[0]) if 'http' in item_url else item_url
                                    if not item_name or item_name == item_url: item_name = f"Enlace {idx + 1}"
                                    html_parts.append(
                                        f"<li><a href='{html.escape(item_url)}' target='_blank'>{html.escape(item_name)}</a></li>")
                            html_parts.append("</ul>")
                        html_parts.append("</td>")
                    html_parts.append("<td class='selection-group'>")
                    for src_country in ["ES", "UK", "ITA"]:
                        is_ch = str(values['chosen_source']) == src_country
                        r_id = f"radio-{unique_input_group_name}-{src_country}"
                        html_parts.append(
                            f"<input type='radio' name='{unique_input_group_name}' value='{src_country}' id='{r_id}' {'checked' if is_ch else ''} data-sku='{html.escape(sku)}' data-field='{escaped_field_name}' onchange='recordSelection(this)'><label for='{r_id}'>{src_country}</label>")
                    html_parts.append("</td>")

                else:
                    for country_code_txt in ["ES", "UK", "ITA"]:
                        current_val_txt = str(values[country_code_txt])
                        cell_html_content_txt = html.escape(current_val_txt)
                        html_parts.append(
                            f"<td class='country-val {country_code_txt.lower()}-val' data-source='{country_code_txt}'>{cell_html_content_txt}</td>")
                    html_parts.append("<td class='selection-group'>")
                    for src_country_txt in ["ES", "UK", "ITA"]:
                        is_ch_txt = str(values['chosen_source']) == src_country_txt
                        r_id_txt = f"radio-{unique_input_group_name}-{src_country_txt}"
                        html_parts.append(
                            f"<input type='radio' name='{unique_input_group_name}' value='{src_country_txt}' id='{r_id_txt}' {'checked' if is_ch_txt else ''} data-sku='{html.escape(sku)}' data-field='{escaped_field_name}' onchange='recordSelection(this)'><label for='{r_id_txt}'>{src_country_txt}</label>")
                    html_parts.append("</td>")
                html_parts.append("</tr>")
            html_parts.append("</tbody></table></div>")

        js_url_list_fields_array_str = json.dumps(self.url_list_fields)
        js_structured_fields_array_str = json.dumps(self.structured_fields)

        script_initial_declarations = f"""
                const selections = {{}};
                const urlListFieldsJS = {js_url_list_fields_array_str};
                const structuredFieldsJS = {js_structured_fields_array_str};

                let currentSkuIndex = 0; 
                let skuPages = [];       
                let prevSkuBtn = null;   
                let nextSkuBtn = null;
                let skuNavStatus = null;
        """

        html_parts.append(f"""
            </div> 
            <script>
                {script_initial_declarations}

                function handleSpecCheckboxChange(checkboxElement) {{
                    const groupName = checkboxElement.name;
                    if (checkboxElement.checked) {{
                        document.querySelectorAll(`input[type="checkbox"][name="${{groupName}}"]`).forEach(cb => {{
                            if (cb !== checkboxElement) {{
                                cb.checked = false;
                            }}
                        }});
                    }}
                    recordSelection(checkboxElement); 
                }}

                function initializeSkuPagination() {{
                    if (skuPages.length > 0) {{
                        updateNavigationControls(); 
                        if (skuPages.length === 1) {{
                            if(prevSkuBtn) prevSkuBtn.style.display = 'none';
                            if(nextSkuBtn) nextSkuBtn.style.display = 'none';
                        }}
                    }} else {{ 
                        const navContainer = document.getElementById('sku-navigation-container');
                        if(navContainer) navContainer.style.display = 'none';
                        const skusContentContainer = document.getElementById('skus-content-container');
                        if (skusContentContainer) {{
                             skusContentContainer.innerHTML = "<p style='text-align:center; font-size:1.2em; margin-top:30px;'>No se encontraron SKUs con diferencias para mostrar.</p>";
                        }}
                    }}
                }}

                function showSkuPage(index) {{
                    if (index < 0 || index >= skuPages.length) return;

                    if (skuPages[currentSkuIndex]) {{
                        skuPages[currentSkuIndex].classList.remove('active');
                    }}
                    if (skuPages[index]) {{
                        skuPages[index].classList.add('active');
                    }}
                    currentSkuIndex = index; 
                    updateNavigationControls();

                    if (skuPages[currentSkuIndex]) {{ 
                        const pageTop = skuPages[currentSkuIndex].offsetTop;
                        const navContainer = document.getElementById('sku-navigation-container');
                        const navHeight = navContainer ? navContainer.offsetHeight : 0;
                        window.scrollTo({{ top: pageTop - navHeight - 10, behavior: 'smooth' }});
                    }}
                }}

                function navigateSku(direction) {{
                    const newIndex = currentSkuIndex + direction; 
                    showSkuPage(newIndex);
                }}

                function updateNavigationControls() {{
                    if (!skuPages || skuPages.length === 0) {{ return; }}
                    if(prevSkuBtn) prevSkuBtn.disabled = currentSkuIndex === 0;
                    if(nextSkuBtn) nextSkuBtn.disabled = currentSkuIndex >= skuPages.length - 1; 
                    if(skuNavStatus && typeof skuNavStatus.textContent !== 'undefined') {{
                        skuNavStatus.textContent = "SKU " + (currentSkuIndex + 1) + " de " + skuPages.length;
                    }}
                }}

                function escapeHtml(unsafeText) {{
                    const div = document.createElement('div');
                    div.innerText = unsafeText;
                    return div.innerHTML;
                }}

                function recordSelection(element) {{
                    const sku = element.dataset.sku; 
                    const field = element.dataset.field;
                    if (!selections[sku]) selections[sku] = {{}};

                    if (field === 'Image_Urls') {{
                        const cbs = document.querySelectorAll(`input[type="checkbox"][data-sku="${{sku}}"][data-field="Image_Urls"]:checked`);
                        selections[sku][field] = Array.from(cbs).map(cb => cb.value);
                        updateImageSelectionDisplay(sku, field);
                    }} else if (structuredFieldsJS.includes(field.toLowerCase()) && element.type === 'checkbox') {{
                        const allCheckedSpecCheckboxesForField = document.querySelectorAll(`input[type="checkbox"][data-sku="${{sku}}"][data-field="${{field}}"]:checked`);
                        const currentSelectedItems = [];
                        allCheckedSpecCheckboxesForField.forEach(cb => {{
                            currentSelectedItems.push({{
                                key: cb.dataset.specKey,
                                value: cb.dataset.specValue,
                                source_country: cb.dataset.sourceCountry
                            }});
                        }});
                        selections[sku][field] = currentSelectedItems;
                        updateStructuredFieldSelectionDisplay(sku, field, currentSelectedItems);

                    }} else if (element.type === 'radio' && element.checked) {{ 
                        selections[sku][field] = element.value;
                        const tr = element.closest('tr');
                        if (tr) {{
                            tr.querySelectorAll('td.country-val').forEach(td => td.classList.remove('chosen-value-cell'));
                            const targetTd = tr.querySelector(`td.country-val[data-source="${{element.value}}"]`);
                            if (targetTd) targetTd.classList.add('chosen-value-cell');
                        }}
                    }}
                }}

                function toggleImageSelection(imgEl, cbId) {{ const cb = document.getElementById(cbId); if (cb) {{ cb.checked = !cb.checked; imgEl.classList.toggle('selected-image', cb.checked); recordSelection(cb); }} }}

                function updateImageSelectionDisplay(sku, field) {{
                    const safeId = field.replace(/[^a-zA-Z0-9_]/g, '_');
                    const dispId = `selection-display-img-${{sku}}-${{safeId}}`;
                    const dispCell = document.getElementById(dispId);
                    if (!dispCell) return;
                    const urls = selections[sku]?.[field] || [];
                    if (urls.length > 0) {{
                        dispCell.innerHTML = "Seleccionadas: " + urls.map(url => `<img src="${{escapeHtml(url)}}" style="max-width:50px; max-height:50px; margin-right:5px; vertical-align:middle; border:1px solid #ddd;">`).join('');
                    }} else dispCell.innerHTML = "Ninguna seleccionada";
                }}

                function updateStructuredFieldSelectionDisplay(sku, field, selectedItems) {{
                    const safeFieldName = field.replace(/[^a-zA-Z0-9_]/g, '_');
                    const displayCellId = `selection-display-${{sku}}-${{safeFieldName}}`;
                    const displayCell = document.getElementById(displayCellId);
                    if (!displayCell) return;

                    if (selectedItems && selectedItems.length > 0) {{
                        let html = "Seleccionados:<ul class='specifications-chosen-list'>";
                        selectedItems.forEach(cs => {{
                            html += `<li><strong>${{escapeHtml(cs.key)}}:</strong> ${{escapeHtml(cs.value)}} <em>(${{escapeHtml(cs.source_country)}})</em></li>`;
                        }});
                        html += "</ul>";
                        displayCell.innerHTML = html;
                    }} else {{
                        displayCell.innerHTML = "Ninguno seleccionado";
                    }}
                }}

                async function saveAllSelections() {{
                    const saveButton = document.getElementById('save-button'); const statusDiv = document.getElementById('status-message');
                    saveButton.disabled = true; saveButton.textContent = 'Guardando...'; statusDiv.textContent = 'Enviando...'; statusDiv.style.color = '#ffc107';
                    try {{
                        const response = await fetch('http://localhost:5050/save_selections', {{ method: 'POST', headers: {{ 'Content-Type': 'application/json' }}, body: JSON.stringify(selections) }});
                        const result = await response.json();
                        if (response.ok && result.status === 'success') {{ statusDiv.textContent = '¡Guardado con éxito!'; statusDiv.style.color = '#28a745'; saveButton.textContent = 'Guardado'; }}
                        else {{ statusDiv.textContent = 'Error: ' + (result.message || 'Desconocido'); statusDiv.style.color = '#dc3545'; saveButton.disabled = false; saveButton.textContent = 'Reintentar'; }}
                    }} catch (error) {{ statusDiv.textContent = 'Error de red: ' + error.message; statusDiv.style.color = '#dc3545'; saveButton.disabled = false; saveButton.textContent = 'Reintentar'; }}
                }}

                document.addEventListener('DOMContentLoaded', () => {{
                    console.log("--- DOMContentLoaded START (Paginación Reactivada) ---");
                    try {{
                        skuPages = Array.from(document.querySelectorAll('.sku-page')); 
                        prevSkuBtn = document.getElementById('prev-sku-btn');
                        nextSkuBtn = document.getElementById('next-sku-btn');
                        skuNavStatus = document.getElementById('sku-nav-status');

                        initializeSkuPagination(); 

                        document.querySelectorAll('input[type="radio"][data-sku][data-field], input[type="checkbox"][data-sku][data-field]').forEach(inp => {{
                            if (inp.checked) {{
                               // Para campos estructurados, el handleSpecCheckboxChange no se llamará en la carga inicial
                               // así que necesitamos llamar a recordSelection directamente.
                               // Para radio y otros checkboxes, esto también es seguro.
                               recordSelection(inp); 
                            }}
                        }});

                        const activePage = document.querySelector('.sku-page.active');
                        if (activePage) {{
                            activePage.querySelectorAll('table tbody tr').forEach(tr => {{
                                const firstRadio = tr.querySelector('input[type="radio"][data-field]:checked');
                                 if (firstRadio) {{
                                    const fieldName = firstRadio.dataset.field;
                                    if (!structuredFieldsJS.includes(fieldName.toLowerCase())) {{ 
                                        tr.querySelectorAll('td.country-val').forEach(td => td.classList.remove('chosen-value-cell'));
                                        const targetTd = tr.querySelector(`td.country-val[data-source="${{firstRadio.value}}"]`);
                                        if (targetTd) targetTd.classList.add('chosen-value-cell');
                                    }}
                                }}
                            }});
                        }}
                       console.log("--- DOMContentLoaded END (Paginación Reactivada) ---");
                    }} catch (e) {{
                        console.error("Error during DOMContentLoaded (Paginación Reactivada):", e.message, e.stack);
                        alert("JavaScript Initialization Error (Paginación Reactivada - consola para detalles): " + e.message);
                    }}
                }});
                console.log("--- SCRIPT END (Paginación Reactivada) ---");
            </script></body></html>""")

        final_html_content = "".join(html_parts)
        try:
            with open(self.html_output_path, "w", encoding="utf-8") as f:
                f.write(final_html_content)
            self._emit_progress(f"Archivo HTML interactivo generado en: {self.html_output_path}")
        except Exception as e_write:
            self._emit_progress(f"Error al escribir el archivo HTML: {e_write}")
            raise

    def process_and_generate_html(self):
        try:
            self._emit_progress("Iniciando proceso de comparación y generación HTML.")
            self._load_excel()
            self._load_user_selections()
            self._identify_all_skus_to_process()
            self._compare_products()
            self._generate_html()
            self._emit_progress("Proceso de comparación y generación HTML completado.")
        except Exception as e:
            self._emit_progress(f"Fallo el proceso principal de ComparisonHandler: {type(e).__name__} - {e}")
            import traceback
            tb_str = traceback.format_exc()
            self._emit_progress(f"Traceback: {tb_str}")
            try:
                if self.html_output_path:
                    with open(self.html_output_path, "w", encoding="utf-8") as f_err:
                        f_err.write(
                            f"<html><body><h1>Error en la Generación del HTML</h1><p>{html.escape(str(e))}</p><pre>{html.escape(tb_str)}</pre></body></html>")
                    self._emit_progress(f"HTML de error generado en: {self.html_output_path}")
            except Exception as e_html_err:
                self._emit_progress(f"No se pudo generar HTML de error: {e_html_err}")
            raise

    def generate_merged_data_from_selections(self):
        self._emit_progress("Iniciando generación de datos finales según selecciones.")
        self._load_excel()
        self._load_user_selections()

        if not self.user_selections:
            self._emit_progress(
                "Advertencia: No hay selecciones de usuario. El archivo final podría estar incompleto o usar valores por defecto/primera fuente si los campos no se tocaron.")

        merged_products_list = []
        id_col_name = self.sku_column_name

        skus_in_selections = set(self.user_selections.keys())
        excel_skus_set = set()
        if not self.data_es.empty and id_col_name in self.data_es.columns: excel_skus_set.update(
            self.data_es[id_col_name].astype(str).unique())
        if not self.data_uk.empty and id_col_name in self.data_uk.columns: excel_skus_set.update(
            self.data_uk[id_col_name].astype(str).unique())
        if not self.data_ita.empty and id_col_name in self.data_ita.columns: excel_skus_set.update(
            self.data_ita[id_col_name].astype(str).unique())

        all_unique_skus_for_final_merge = sorted(list(skus_in_selections | excel_skus_set))

        if not all_unique_skus_for_final_merge:
            self._emit_progress("No hay SKUs para generar datos fusionados.")
            return []

        self._emit_progress(f"Generando datos para {len(all_unique_skus_for_final_merge)} SKUs.")

        if not self.all_columns_from_excel and (
                not self.data_es.empty or not self.data_uk.empty or not self.data_ita.empty):
            temp_cols = set()
            if not self.data_es.empty: temp_cols.update(self.data_es.columns)
            if not self.data_uk.empty: temp_cols.update(self.data_uk.columns)
            if not self.data_ita.empty: temp_cols.update(self.data_ita.columns)
            if id_col_name in temp_cols: temp_cols.remove(id_col_name)
            self.all_columns_from_excel = temp_cols

        ordered_columns = [id_col_name] + sorted(list(self.all_columns_from_excel))

        for sku_val_str in all_unique_skus_for_final_merge:
            merged_product_row = {col: '' for col in ordered_columns}
            merged_product_row[id_col_name] = sku_val_str
            sku_specific_selections = self.user_selections.get(sku_val_str, {})

            row_es_dict = self.data_es[self.data_es[id_col_name] == sku_val_str].iloc[0].to_dict() if not self.data_es[
                self.data_es[id_col_name] == sku_val_str].empty else {}
            row_uk_dict = self.data_uk[self.data_uk[id_col_name] == sku_val_str].iloc[0].to_dict() if not self.data_uk[
                self.data_uk[id_col_name] == sku_val_str].empty else {}
            row_ita_dict = self.data_ita[self.data_ita[id_col_name] == sku_val_str].iloc[0].to_dict() if not \
            self.data_ita[self.data_ita[id_col_name] == sku_val_str].empty else {}

            for field_name in ordered_columns:
                if field_name == id_col_name: continue

                chosen_value_from_selection = sku_specific_selections.get(field_name)
                final_value_for_field = ''

                if field_name.lower() in self.structured_fields:
                    final_specs_dict = {}

                    original_parsed_es = self._parse_specifications(row_es_dict.get(field_name, ''))
                    original_parsed_uk = self._parse_specifications(row_uk_dict.get(field_name, ''))
                    original_parsed_ita = self._parse_specifications(row_ita_dict.get(field_name, ''))

                    map_orig_es = {s['key']: s['value'] for s in original_parsed_es}
                    map_orig_uk = {s['key']: s['value'] for s in original_parsed_uk}
                    map_orig_ita = {s['key']: s['value'] for s in original_parsed_ita}

                    all_original_keys = set(map_orig_es.keys()) | set(map_orig_uk.keys()) | set(map_orig_ita.keys())

                    if isinstance(chosen_value_from_selection, list):
                        for sel_item in chosen_value_from_selection:
                            final_specs_dict[sel_item['key']] = sel_item['value']

                    priority_sources_data = {'ES': map_orig_es, 'UK': map_orig_uk, 'ITA': map_orig_ita}
                    source_priority_order = ['ES', 'UK', 'ITA']

                    for key in all_original_keys:
                        if key in final_specs_dict:
                            continue

                        values_for_key_map = {}
                        sources_that_have_this_key = []
                        if key in map_orig_es:
                            values_for_key_map["ES"] = map_orig_es[key]
                            sources_that_have_this_key.append("ES")
                        if key in map_orig_uk:
                            values_for_key_map["UK"] = map_orig_uk[key]
                            sources_that_have_this_key.append("UK")
                        if key in map_orig_ita:
                            values_for_key_map["ITA"] = map_orig_ita[key]
                            sources_that_have_this_key.append("ITA")

                        valid_values = [v for v in values_for_key_map.values() if
                                        v is not None and str(v).strip() != ""]

                        if not valid_values:
                            continue

                        if len(set(valid_values)) == 1 and len(valid_values) == len(sources_that_have_this_key):
                            final_specs_dict[key] = valid_values[0]
                        elif valid_values:
                            for source_code in source_priority_order:
                                if key in priority_sources_data[source_code] and \
                                        priority_sources_data[source_code][key] is not None and \
                                        str(priority_sources_data[source_code][key]).strip() != "":
                                    final_specs_dict[key] = priority_sources_data[source_code][key]
                                    break

                    final_value_for_field = json.dumps(final_specs_dict,
                                                       ensure_ascii=False) if final_specs_dict else '{}'

                elif field_name == 'Image_Urls':
                    if isinstance(chosen_value_from_selection, list) and chosen_value_from_selection:
                        final_value_for_field = str(chosen_value_from_selection)
                    elif chosen_value_from_selection and not isinstance(chosen_value_from_selection, list):
                        final_value_for_field = str([chosen_value_from_selection])
                    else:
                        val_es_img = self._parse_generic_urls(row_es_dict.get(field_name, ''))
                        val_uk_img = self._parse_generic_urls(row_uk_dict.get(field_name, ''))
                        val_ita_img = self._parse_generic_urls(row_ita_dict.get(field_name, ''))
                        final_value_for_field = str(val_es_img or val_uk_img or val_ita_img or [])

                elif field_name in self.url_list_fields:
                    source_map = {'ES': row_es_dict, 'UK': row_uk_dict, 'ITA': row_ita_dict}
                    if chosen_value_from_selection in source_map:
                        final_value_for_field = str(
                            self._parse_generic_urls(source_map[chosen_value_from_selection].get(field_name, '')))
                    else:
                        val_es_list = self._parse_generic_urls(row_es_dict.get(field_name, ''))
                        val_uk_list = self._parse_generic_urls(row_uk_dict.get(field_name, ''))
                        val_ita_list = self._parse_generic_urls(row_ita_dict.get(field_name, ''))
                        final_value_for_field = str(val_es_list or val_uk_list or val_ita_list or [])

                elif chosen_value_from_selection in ["ES", "UK", "ITA"]:
                    source_map_txt = {'ES': row_es_dict, 'UK': row_uk_dict, 'ITA': row_ita_dict}
                    final_value_for_field = source_map_txt[chosen_value_from_selection].get(field_name, '')

                elif chosen_value_from_selection is not None and \
                        not field_name.lower() in self.structured_fields and \
                        not field_name in self.url_list_fields and \
                        field_name != 'Image_Urls':
                    final_value_for_field = chosen_value_from_selection

                else:
                    final_value_for_field = row_es_dict.get(field_name, '') or \
                                            row_uk_dict.get(field_name, '') or \
                                            row_ita_dict.get(field_name, '')

                merged_product_row[field_name] = str(final_value_for_field)
            merged_products_list.append(merged_product_row)

        self._emit_progress("Datos fusionados finales listos para ser guardados.")
        return merged_products_list

