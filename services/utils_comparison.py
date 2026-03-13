# services/utils_comparison.py
import pandas as pd
import json
import os
import html
import ast
from typing import Tuple, Dict
import unicodedata
from services.campos_odoo import ClavesExcel
from services.utils import Utils

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
        self.data_opt = pd.DataFrame()

        self.all_skus_to_process = []
        self.comparison_results = {}
        self.user_selections = {}
        self.all_columns_from_excel = set()
        self.sku_column_name = ClavesExcel.SKU.value

        self.url_list_fields = [ClavesExcel.DOCUMENTOS.value, ClavesExcel.VIDEOS.value]
        self.structured_fields = [ClavesExcel.ATRIBUTOS.value.lower()]

        self.skus_nuevos = []
        self.skus_eliminados = []

    def _emit_progress(self, message):
        if self.event_manager:
            self.event_manager.emit('progress_update', f"[ComparisonHandler] {message}")
        else:
            print(f"[ComparisonHandler] {message}")

    def _identify_new_and_deleted_products(self):
        scrap_skus = set()
        if not self.data_es.empty: scrap_skus.update(self.data_es[self.sku_column_name].astype(str).unique())
        if not self.data_uk.empty: scrap_skus.update(self.data_uk[self.sku_column_name].astype(str).unique())
        if not self.data_ita.empty: scrap_skus.update(self.data_ita[self.sku_column_name].astype(str).unique())

        opt_skus = set()
        if not self.data_opt.empty: opt_skus.update(self.data_opt[self.sku_column_name].astype(str).unique())

        self.skus_nuevos = sorted(list(scrap_skus - opt_skus))
        self.skus_eliminados = sorted(list(opt_skus - scrap_skus))

    def load_excel(self):
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
            elif 'OPT' in xls.sheet_names:
                first_sheet_to_read_cols = 'OPT'
            elif xls.sheet_names:
                first_sheet_to_read_cols = xls.sheet_names[0]

            if not first_sheet_to_read_cols:
                raise ValueError("El archivo Excel no contiene ninguna hoja.")

            df_cols_check = pd.read_excel(xls, first_sheet_to_read_cols, nrows=0)
            dtype_all_str = {col: str for col in df_cols_check.columns}

            for col in df_cols_check.columns:
                if col.upper() == self.sku_column_name:
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
            if 'OPT' in xls.sheet_names:
                self.data_opt = pd.read_excel(xls, 'OPT', dtype=dtype_all_str).fillna('')
                sheet_names_found.append('OPT')
                self.all_columns_from_excel.update(self.data_opt.columns)

            if not sheet_names_found:
                raise ValueError("Ninguna de las hojas esperadas (ES, UK, ITA, OPT) fue encontrada en el Excel.")
            if self.data_es.empty and self.data_uk.empty and self.data_ita.empty:
                self._emit_progress("Advertencia: Las hojas ES, UK, ITA, OPT están vacías o no contienen datos.")

            for df, name in [(self.data_es, 'ES'), (self.data_uk, 'UK'), (self.data_ita, 'ITA'), (self.data_opt, 'OPT')]:
                if not df.empty and self.sku_column_name not in df.columns:
                    raise ValueError(f"La columna '{self.sku_column_name}' no se encontró en la hoja '{name}'.")

            if self.sku_column_name in self.all_columns_from_excel:
                self.all_columns_from_excel.remove(self.sku_column_name)

            # Limpiar espacios en SKU en todas las hojas cargadas
            for df in [self.data_es, self.data_uk, self.data_ita, self.data_opt]:
                if not df.empty and self.sku_column_name in df.columns:
                    df[self.sku_column_name] = df[self.sku_column_name].astype(str).str.strip()

            # Guardar Excel con SKUs corregidos
            self.column_order = list(df_cols_check.columns)
            try:
                clean_excel_path = self.excel_path
                with pd.ExcelWriter(clean_excel_path, engine='openpyxl') as writer:
                    if not self.data_es.empty:
                        self.data_es.to_excel(writer, sheet_name='ES', index=False)
                    if not self.data_uk.empty:
                        self.data_uk.to_excel(writer, sheet_name='UK', index=False)
                    if not self.data_ita.empty:
                        self.data_ita.to_excel(writer, sheet_name='ITA', index=False)
                    if not self.data_opt.empty:
                        self.data_opt.to_excel(writer, sheet_name='OPT', index=False)

            except Exception as e:
                self._emit_progress(f"❌ Error al guardar Excel limpio: {e}")

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

    def load_user_selections(self):
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
                               isinstance(url, str) and (str(url).strip().endswith('jpg') or str(url).strip().startswith('http'))]
            elif isinstance(evaluated_data, dict):
                for value in evaluated_data.values():
                    if isinstance(value, str) and (value.strip().endswith('jpg') or value.strip().startswith('http')):
                        parsed_urls.append(value.strip())
        except (ValueError, SyntaxError, TypeError):
            if 'jpg' in urls_str or 'http' in urls_str:
                if not urls_str.strip().startswith('[') and not urls_str.strip().startswith('{'):
                    possible_urls_csv = [url.strip() for url in urls_str.split(',') if url.strip().endswith('jpg') or url.strip().startswith('http')]
                    if possible_urls_csv: return possible_urls_csv
                if urls_str.strip().endswith('jpg') or urls_str.strip().startswith('http'): return [urls_str.strip()]
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
                    parsed_specs.append({'key': ClavesExcel.ATRIBUTOS.value, 'value': str(evaluated_data).strip()})
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

        MAX_SKUS_TO_COMPARE = 20000

        for sku_idx, sku_val_str in enumerate(self.all_skus_to_process[:MAX_SKUS_TO_COMPARE]):
            if (sku_idx + 1) % 20 == 0 or sku_idx == 0:
                self._emit_progress(f"Comparando SKU {sku_idx + 1}/{total_skus}: {sku_val_str}")

            prod_comparison_data = {}
            row_es_series = self.data_es[self.data_es[id_col_name] == sku_val_str].iloc[0] if not self.data_es[
                self.data_es[id_col_name] == sku_val_str].empty else pd.Series(dtype=str)
            row_uk_series = self.data_uk[self.data_uk[id_col_name] == sku_val_str].iloc[0] if not self.data_uk[
                self.data_uk[id_col_name] == sku_val_str].empty else pd.Series(dtype=str)
            row_ita_series = self.data_ita[self.data_ita[id_col_name] == sku_val_str].iloc[0] if not self.data_ita[
                self.data_ita[id_col_name] == sku_val_str].empty else pd.Series(dtype=str)
            row_opt_series = self.data_opt[self.data_opt[id_col_name] == sku_val_str].iloc[0] if not self.data_opt[
                self.data_opt[id_col_name] == sku_val_str].empty else pd.Series(dtype=str)

            for field in self.column_order: #for field in sorted(list(self.all_columns_from_excel)):
                if field not in self.all_columns_from_excel:
                    continue
                val_es_orig = str(row_es_series.get(field, ''))
                val_uk_orig = str(row_uk_series.get(field, ''))
                val_ita_orig = str(row_ita_series.get(field, ''))
                val_opt_orig = str(row_opt_series.get(field, ''))

                if field.lower() in self.structured_fields:
                    parsed_es = self._parse_specifications(val_es_orig)
                    parsed_uk = self._parse_specifications(val_uk_orig)
                    parsed_ita = self._parse_specifications(val_ita_orig)
                    parsed_opt = self._parse_specifications(val_opt_orig)

                    map_es = {s['key']: s['value'] for s in parsed_es}
                    map_uk = {s['key']: s['value'] for s in parsed_uk}
                    map_ita = {s['key']: s['value'] for s in parsed_ita}
                    map_opt = {s['key']: s['value'] for s in parsed_opt}

                    # Marcar 'is_different' en los items de las listas parseadas
                    for spec_list_country_idx, spec_list_country_data in enumerate(
                            [(parsed_es, "ES"), (parsed_uk, "UK"), (parsed_ita, "ITA"), (parsed_opt, "OPT")]):
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
                            if parsed_opt: sources_with_specs.append(map_opt)

                            # Fuentes que tienen ESTA clave
                            sources_with_this_key = []
                            if key_to_check in map_es: sources_with_this_key.append(map_es)
                            if key_to_check in map_uk: sources_with_this_key.append(map_uk)
                            if key_to_check in map_ita: sources_with_this_key.append(map_ita)
                            if key_to_check in map_opt: sources_with_this_key.append(map_opt)

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
                    if isinstance(chosen_items_list, str) and chosen_items_list in ['ES', 'UK', 'ITA', 'OPT']:
                        default_source_country = chosen_items_list
                        source_parsed_list_for_migration = []
                        if default_source_country == 'ES':
                            source_parsed_list_for_migration = parsed_es
                        elif default_source_country == 'UK':
                            source_parsed_list_for_migration = parsed_uk
                        elif default_source_country == 'ITA':
                            source_parsed_list_for_migration = parsed_ita
                        elif default_source_country == 'OPT':
                            source_parsed_list_for_migration = parsed_opt
                        chosen_items_list = [
                            {'key': s['key'], 'value': s['value'], 'source_country': default_source_country} for s in
                            source_parsed_list_for_migration]
                    elif not isinstance(chosen_items_list, list):
                        chosen_items_list = []

                    prod_comparison_data[field] = {
                        "ES_parsed": parsed_es, "UK_parsed": parsed_uk, "ITA_parsed": parsed_ita, "OPT_parsed": parsed_opt,
                        "ES": val_es_orig, "UK": val_uk_orig, "ITA": val_ita_orig, "OPT": val_opt_orig,
                        "has_difference": field_has_difference,
                        "chosen_source": chosen_items_list
                    }
                elif field == ClavesExcel.GALERIA.value or field == ClavesExcel.VIDEOS.value:#field in self.url_list_fields:
                    list_es = self._parse_generic_urls(val_es_orig)
                    list_uk = self._parse_generic_urls(val_uk_orig)
                    list_ita = self._parse_generic_urls(val_ita_orig)
                    list_opt = self._parse_generic_urls(val_opt_orig)

                    all_lists_tuples = [tuple(sorted(l)) for l in [list_es, list_uk, list_ita, list_opt] if l]
                    has_difference_urls = len(set(all_lists_tuples)) > 1 if all_lists_tuples else False

                    chosen_val = self.user_selections.get(sku_val_str, {}).get(field)
                    if field == ClavesExcel.GALERIA.value:
                        chosen_val = chosen_val if isinstance(chosen_val, list) else (
                            [chosen_val] if chosen_val else [])

                    prod_comparison_data[field] = {
                        "ES": list_es, "UK": list_uk, "ITA": list_ita, "OPT": list_opt,
                        "has_difference": has_difference_urls,
                        "chosen_source": chosen_val
                    }
                elif field == ClavesExcel.DOCUMENTOS.value:
                    def parse_pdf_dict(url_str):
                        try:
                            parsed_dict = ast.literal_eval(url_str)
                            if isinstance(parsed_dict, dict):
                                return {str(k).strip(): str(v).strip() for k, v in parsed_dict.items() if v.strip()}
                        except Exception:
                            pass
                        return {}

                    dict_es = parse_pdf_dict(val_es_orig)
                    dict_uk = parse_pdf_dict(val_uk_orig)
                    dict_ita = parse_pdf_dict(val_ita_orig)
                    dict_opt = parse_pdf_dict(val_opt_orig)

                    # Comparación por claves y enlaces
                    sets_comp = [
                        set((k, v) for k, v in d.items())
                        for d in [dict_es, dict_uk, dict_ita, dict_opt] if d
                    ]
                    has_difference_urls = len(set(frozenset(s) for s in sets_comp)) > 1 if sets_comp else False

                    chosen_val = self.user_selections.get(sku_val_str, {}).get(field, {})

                    prod_comparison_data[field] = {
                        "ES": dict_es, "UK": dict_uk, "ITA": dict_ita, "OPT": dict_opt,
                        "has_difference": has_difference_urls,
                        "chosen_source": chosen_val if isinstance(chosen_val, dict) else {}
                    }
                else:
                    distinct_values_present = {v for v in [val_es_orig, val_uk_orig, val_ita_orig, val_opt_orig] if v}
                    has_difference_txt = len(distinct_values_present) > 1
                    chosen_source_for_field = self.user_selections.get(sku_val_str, {}).get(field, None)
                    prod_comparison_data[field] = {
                        "ES": val_es_orig, "UK": val_uk_orig, "ITA": val_ita_orig, "OPT": val_opt_orig,
                        "has_difference": has_difference_txt,
                        "chosen_source": chosen_source_for_field,
                    }
            self.comparison_results[sku_val_str] = prod_comparison_data
        self._emit_progress("Comparación de productos completada.")

    def _generate_html(self, to_compare):
        COUNTRY_DISPLAY_ORDER = ["OPT", "UK", "ITA", "ES"]
        css_init = [f"""
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
                

                <div id="sku-navigation-container">
                  <div id="sku-navigation">
                    <button id="prev-sku-btn" onclick="navigateSku(-1)" disabled>Anterior SKU</button>
                    <span id="sku-nav-status"></span>
                    <button id="next-sku-btn" onclick="navigateSku(1)">Siguiente SKU</button>
                  </div>
                </div>
                <div id="skus-content-container" style="margin-bottom: 40px; padding: 15px; margin-left: 15px; margin-right: 15px;"> 
                """]
        TEST = False    #AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
        html_parts = css_init
        if TEST:
            html_parts.append("""
            <div id="save-button-container">
                <button id="save-button" onclick="saveNewEdits()">Guardar Cambios</button>
                <div id="status-message"></div>
            </div>""")
        else:
            html_parts.append("""<div id="save-button-container"><button id="save-button" onclick="saveAllSelections()">Guardar Selecciones</button><div id="status-message"></div></div>""")

        if TEST:#if not self.comparison_results:
            df_new_products = Utils.detectar_skus_unicos(self.excel_path)
            if TEST:#if df_new_products:
                def _generate_html_nuevos(df_nuevos):
                    self._emit_progress("Generando HTML para nuevos productos detectados...")

                    for idx, (_, row) in enumerate(df_nuevos.iterrows()):
                        sku = str(row.get("SKU", f"SKU_{idx}"))
                        html_parts.append(
                            f"<div class='sku-page {'active' if idx == 0 else ''}' id='sku-page-{html.escape(sku)}'>")
                        html_parts.append(
                            f"<h2 class='sku-header'>SKU: {html.escape(sku)}</h2><table><thead><tr><th>Campo</th><th>Valor</th><th>Edición</th></tr></thead><tbody>")

                        for campo in row.index:
                            if campo.upper() == "SKU":
                                continue
                            valor = str(row[campo]) if pd.notna(row[campo]) else ''
                            safe_id = f"{html.escape(sku)}__{html.escape(campo)}"

                            if campo == ClavesExcel.GALERIA.value:
                                try:
                                    urls = ast.literal_eval(valor)
                                    if not isinstance(urls, list):
                                        urls = [valor]
                                except Exception:
                                    urls = [valor]

                                html_parts.append(
                                    f"<tr><td>{html.escape(campo)}</td><td colspan='2'><table class='sub-table'>")

                                for i, url in enumerate(urls):
                                    src = html.escape(url.strip())
                                    img_src = "http://143.47.53.74:8070/" + src if "http" not in src else src
                                    input_id = f"{safe_id}_{i}"
                                    html_parts.append(
                                        f"<tr><td style='width:50%;'><img src='{img_src}' style='max-width:80px; max-height:80px;'></td>"
                                        f"<td style='width:50%;'>"
                                        f"<input type='text' id='{input_id}' value='{img_src}' style='width:100%;' "
                                        f"onchange='recordEditImagen(\"{html.escape(sku)}\", \"{html.escape(campo)}\", {i}, this.value)'>"
                                        f"</td></tr>")
                                html_parts.append("</table></td></tr>")

                            elif campo == ClavesExcel.IMAGEN.value:

                                src = valor.strip()
                                img_url = "http://143.47.53.74:8070/" + src if "http" not in src else src
                                safe_img_id = f"{safe_id}_img"
                                html_parts.append(f"<tr><td>{html.escape(campo)}</td>")

                                # Columna Valor → Imagen visible
                                if img_url:
                                    html_parts.append(
                                        f"<td><img src='{html.escape(img_url)}' style='max-width:120px; max-height:120px;'></td>")
                                else:
                                    html_parts.append("<td>–</td>")
                                # Columna Edición → Input de URL
                                html_parts.append(
                                    f"<td><input type='text' id='{safe_img_id}' value='{html.escape(img_url)}' "
                                    f"onchange='recordEdit(\"{html.escape(sku)}\", \"{html.escape(campo)}\", this.value)' style='width:100%;'></td></tr>")

                            elif campo == ClavesExcel.DESCRIPCION_WEB.value:
                                html_parts.append(f"<tr><td>{html.escape(campo)}</td><td>{html.escape(valor)}</td><td>")
                                html_parts.append(
                                    f"<textarea id='{safe_id}' "
                                    f"oninput='autoResize(this); recordEdit(\"{html.escape(sku)}\", \"{html.escape(campo)}\", this.value)' "
                                    f"style='width:100%; min-height:80px; resize:vertical;'>{html.escape(valor)}</textarea>"
                                )
                                html_parts.append("</td></tr>")

                            else:
                                rendered_as_list = False
                                try:
                                    parsed_val = ast.literal_eval(valor)
                                    if isinstance(parsed_val, list):
                                        rendered_as_list = True
                                        html_parts.append(
                                            f"<tr><td>{html.escape(campo)}</td><td colspan='2'><table class='sub-table'>")
                                        for idx, item in enumerate(parsed_val):
                                            item_val = html.escape(str(item))
                                            input_id = f"{safe_id}_{idx}"
                                            html_parts.append(
                                                f"<tr><td style='width:50%;'>{item_val}</td><td style='width:50%;'>"
                                                f"<input type='text' id='{input_id}' value='{item_val}' style='width:100%;' "
                                                f"onchange='recordEditArray(\"{html.escape(sku)}\", \"{html.escape(campo)}\", {idx}, this.value)'></td></tr>")
                                        html_parts.append("</table></td></tr>")

                                    elif isinstance(parsed_val, dict):
                                        '''if campo == ClavesExcel.ATRIBUTOS.value:
                                            docs_dict = ast.literal_eval(row[ClavesExcel.DOCUMENTOS.value])
                                            if isinstance(docs_dict, dict):
                                                docs_items = set(
                                                    (k.strip(), v.strip()) for k, v in docs_dict.items() if k and v)
                                                parsed_val = {
                                                    k: v for k, v in parsed_val.items()
                                                    if (k.strip(), v.strip()) not in docs_items
                                                }'''

                                        rendered_as_list = True
                                        html_parts.append(
                                            f"<tr><td>{html.escape(campo)}</td><td colspan='2'><table class='sub-table'>")
                                        for k, v in parsed_val.items():
                                            k_html = html.escape(str(k))
                                            v_html = html.escape(str(v))
                                            full_string = f"{k_html}: {v_html}"
                                            input_id = f"{safe_id}_{k_html}"
                                            html_parts.append(
                                                f"<tr><td style='width:50%;'><b>{k_html}</b>: {v_html}</td><td style='width:50%;'>"
                                                f"<input type='text' id='{input_id}' value='{full_string}' style='width:100%; font-weight: bolder;' "
                                                f"onchange='recordEditDictLine(\"{html.escape(sku)}\", \"{html.escape(campo)}\", \"{k_html}\", this.value)'></td></tr>")
                                        html_parts.append("</table></td></tr>")
                                except Exception:
                                    rendered_as_list = False

                                if not rendered_as_list:
                                    html_parts.append(f"<tr><td>{html.escape(campo)}</td><td>{html.escape(valor)}</td>"
                                                      f"<td><input type='text' id='{safe_id}' value='{html.escape(valor)}' "
                                                      f"onchange='recordEdit(\"{html.escape(sku)}\", \"{html.escape(campo)}\", this.value)' style='width:95%'></td></tr>")

                        html_parts.append("</tbody></table></div>")

                    html_parts.append(f"""
                        </div>
                        <script>
                        const skuPages = Array.from(document.querySelectorAll('.sku-page'));
                        let currentSkuIndex = 0;
                        let prevSkuBtn = document.getElementById('prev-sku-btn');
                        let nextSkuBtn = document.getElementById('next-sku-btn');
                        let skuNavStatus = document.getElementById('sku-nav-status');

                        const edits = {{}};

                        function recordEdit(sku, campo, nuevoValor) {{
                            if (!edits[sku]) edits[sku] = {{}};
                            edits[sku][campo] = nuevoValor;
                        }}
                        
                        function recordEditImagen(sku, campo, index, nuevoValor) {{
                            const inputs = document.querySelectorAll(`[id^="${{sku}}__${{campo}}_"]`);
                            const arrayFinal = [];
                        
                            inputs.forEach(input => {{
                                arrayFinal.push(input.value.trim());
                            }});
                        
                            if (!edits[sku]) edits[sku] = {{}};
                            edits[sku][campo] = arrayFinal;
                        }}
                        
                        function autoResize(textarea) {{
                            textarea.style.height = 'auto';
                            textarea.style.height = (textarea.scrollHeight + 2) + 'px';
                        }}

                        function showSkuPage(index) {{
                            if (index < 0 || index >= skuPages.length) return;
                            skuPages[currentSkuIndex].classList.remove('active');
                            skuPages[index].classList.add('active');
                            currentSkuIndex = index;
                            updateNavigationControls();
                            const pageTop = skuPages[currentSkuIndex].offsetTop;
                            const navHeight = document.getElementById('sku-navigation-container').offsetHeight;
                            window.scrollTo({{ top: pageTop - navHeight - 10, behavior: 'smooth' }});
                        }}

                        function navigateSku(direction) {{
                            showSkuPage(currentSkuIndex + direction);
                        }}

                        function updateNavigationControls() {{
                            prevSkuBtn.disabled = currentSkuIndex === 0;
                            nextSkuBtn.disabled = currentSkuIndex >= skuPages.length - 1;
                            skuNavStatus.textContent = "SKU " + (currentSkuIndex + 1) + " de " + skuPages.length;
                        }}

                        function saveNewEdits() {{
                            const btn = document.getElementById('save-button');
                            const status = document.getElementById('status-message');
                            btn.disabled = true;
                            status.textContent = "Guardando...";
                            status.style.color = "#ffc107";

                            fetch('http://localhost:5050/save_new_edits', {{
                                method: 'POST',
                                headers: {{ 'Content-Type': 'application/json' }},
                                body: JSON.stringify(edits)
                            }})
                            .then(res => res.json())
                            .then(result => {{
                                if (result.status === 'success') {{
                                    status.textContent = "✅ Cambios guardados con éxito";
                                    status.style.color = "lightgreen";
                                }} else {{
                                    status.textContent = "❌ Error: " + (result.message || "desconocido");
                                    status.style.color = "red";
                                    btn.disabled = false;
                                }}
                            }})
                            .catch(err => {{
                                status.textContent = "❌ Error de red: " + err.message;
                                status.style.color = "red";
                                btn.disabled = false;
                            }});
                        }}
                        
                        function recordEditArray(sku, campo, index, nuevoValor) {{
                            const inputs = document.querySelectorAll(`[id^="${{sku}}__${{campo}}_"]`);
                            const arrayFinal = [];
                        
                            inputs.forEach(input => {{
                                arrayFinal.push(input.value.trim());
                            }});
                        
                            if (!edits[sku]) edits[sku] = {{}};
                            edits[sku][campo] = arrayFinal;
                        }}
                        
                        function recordEditDictLine(sku, campo, originalKey, lineaTexto) {{
                            const prefix = `${{sku}}__${{campo}}_`;
                            const inputs = document.querySelectorAll(`[id^="${{prefix}}"]`);
                        
                            const dictFinal = {{}};
                            inputs.forEach(input => {{
                                const parts = input.value.split(':');
                                if (parts.length >= 2) {{
                                    const k = parts[0].trim();
                                    const v = parts.slice(1).join(':').trim();
                                    if (k) dictFinal[k] = v;
                                }}
                            }});
                        
                            if (!edits[sku]) edits[sku] = {{}};
                            edits[sku][campo] = dictFinal;
                        }}


                        showSkuPage(0);  // Inicial
                        </script>
                        """)
                    html_parts.append(f"""
                    <script>
                    document.addEventListener('DOMContentLoaded', () => {{
                        const ajustarAnchoColumnas = () => {{
                            const filas = document.querySelectorAll('.sku-page table tr');
                            filas.forEach(tr => {{
                                const tds = tr.querySelectorAll('td');
                                if (tds.length === 3) {{
                                    tds[0].style.width = "10%";
                                    tds[1].style.width = "45%";
                                    tds[2].style.width = "45%";
                                }}
                            }});

                            const cabeceras = document.querySelectorAll('.sku-page table thead tr th');
                            if (cabeceras.length === 3) {{
                                cabeceras[0].style.width = "25%";
                                cabeceras[1].style.width = "37.5%";
                                cabeceras[2].style.width = "37.5%";
                            }}

                            const inputs = document.querySelectorAll('input[type="text"]');
                            inputs.forEach(input => {{
                                input.style.width = "100%";
                                input.style.boxSizing = "border-box";
                                input.style.padding = "6px 8px";
                            }});
                            
                            document.querySelectorAll('textarea').forEach(autoResize);
                        }};

                        ajustarAnchoColumnas();
                    }});
                    </script></body></html>
                    """)

                    try:
                        with open(self.html_output_path, "w", encoding="utf-8") as f:
                            f.write("".join(html_parts))
                        self._emit_progress(f"✅ HTML de nuevos productos generado en: {self.html_output_path}")
                    except Exception as e:
                        self._emit_progress(f"❌ Error al guardar HTML de nuevos productos: {str(e)}")

                _generate_html_nuevos(df_new_products)
                return
            else:
                self._emit_progress("No hay resultados de comparación para generar HTML.")
                html_content = """<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Comparación de Productos</title></head><body><h1>Comparación de Productos</h1><p>No se encontraron productos para comparar o ninguno tiene diferencias.</p></body></html>"""
                with open(self.html_output_path, "w", encoding="utf-8") as f: f.write(html_content)
                self._emit_progress(f"HTML vacío generado en: {self.html_output_path}")
                return

        self._emit_progress(f"Generando archivo HTML en: {self.html_output_path}")

        html_parts = css_init

        skus_with_differences = []
        for sku_test, fields_data_test in self.comparison_results.items():
            if any(details.get('has_difference', False) for details in fields_data_test.values()):
                skus_with_differences.append(sku_test)

        #sorted_skus_for_html = sorted(skus_with_differences)
        MAX_SKUS_TO_RENDER = 500
        sorted_skus_for_html = sorted(skus_with_differences)[:MAX_SKUS_TO_RENDER]

        if not sorted_skus_for_html:
            html_parts.append("<!-- No SKUs with differences to display -->")

        if max(len(self.skus_nuevos), len(self.skus_eliminados)) > 0: sorted(skus_with_differences)

        for sku_idx_html, sku in enumerate(sorted_skus_for_html):

            fields_data = self.comparison_results[sku]
            active_class = 'active' if sku_idx_html == 0 else ''

            html_parts.append(f"<div class='sku-page {active_class}' id='sku-page-{html.escape(sku)}'>")
            html_parts.append(
                f"""<h2 class='sku-header'>SKU: {html.escape(sku)}</h2><table><thead style="position: sticky; top: 60px; z-index: 10; background: white;"><tr><th>Campo</th><th>Valor OPT</th><th>Valor UK</th><th>Valor ITA</th><th>Valor ES</th><th>Selección</th></tr></thead><tbody>""")

            sorted_field_names = fields_data.keys()#sorted(fields_data.keys())
            for field_name in sorted_field_names:
                values = fields_data[field_name]

                if not values.get('has_difference', False):
                    continue

                escaped_field_name = html.escape(field_name)
                safe_field_name_for_js = ''.join(c if c.isalnum() else '_' for c in field_name)
                unique_input_group_name = f"sel-{html.escape(sku)}-{safe_field_name_for_js}"

                html_parts.append(f"<tr><td class='field-name'>{escaped_field_name}</td>")

                if field_name == ClavesExcel.GALERIA.value:
                    selected_image_urls = values.get('chosen_source', [])

                    if not isinstance(selected_image_urls, list): selected_image_urls = [
                        selected_image_urls] if selected_image_urls else []
                    for country_code in COUNTRY_DISPLAY_ORDER:
                        html_parts.append(
                            f"<td class='country-val {country_code.lower()}-val image-container-cell' data-source='{country_code}'><div class='image-container' id='image-container-{html.escape(sku)}-{safe_field_name_for_js}-{country_code}'>")
                        image_urls_list_country = values[country_code]
                        if isinstance(image_urls_list_country, list):
                            for idx, img_url in enumerate(image_urls_list_country):
                                opt_url = "http://143.47.53.74:8070/" if country_code == "OPT" else ""
                                if img_url:
                                    img_url = opt_url + img_url
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

                elif field_name == ClavesExcel.IMAGEN.value:
                    for country_code_img in COUNTRY_DISPLAY_ORDER:
                        opt_port = "http://143.47.53.74:8070/" if country_code_img == "OPT" else ""
                        url_img = str(values.get(country_code_img, '')).strip()

                        if url_img:
                            url_img = opt_port + url_img
                            html_parts.append(
                                f"<td class='country-val {country_code_img.lower()}-val' data-source='{country_code_img}'>")
                            html_parts.append(
                                f"<img src='{html.escape(url_img)}' alt='Imagen {country_code_img}' style='max-width:100px; max-height:100px;'>")
                            html_parts.append("</td>")
                        else:
                            html_parts.append(
                                f"<td class='country-val {country_code_img.lower()}-val' data-source='{country_code_img}'>-</td>")
                    html_parts.append("<td class='selection-group'>")
                    for src_country_img in COUNTRY_DISPLAY_ORDER:
                        is_ch_img = str(values['chosen_source']) == src_country_img #True if src_country_img == "OPT" else
                        r_id_img = f"radio-{unique_input_group_name}-{src_country_img}"
                        html_parts.append(
                            f"<input type='radio' name='{unique_input_group_name}' value='{src_country_img}' id='{r_id_img}' {'checked' if is_ch_img else ''} data-sku='{html.escape(sku)}' data-field='{escaped_field_name}' onchange='recordSelection(this)'><label for='{r_id_img}'>{src_country_img}</label>")
                    html_parts.append("</td>")
                    html_parts.append("</tr>")

                elif field_name.lower() in self.structured_fields:
                    parsed_specs_data = {"ES": values.get("ES_parsed", []), "UK": values.get("UK_parsed", []),
                                         "ITA": values.get("ITA_parsed", []), "OPT": values.get("OPT_parsed", [])}
                    chosen_specs_from_data = values.get("chosen_source", [])

                    def normalizar(valor: str) -> str:
                        """Normaliza un string para comparación insensible a mayúsculas y tildes."""
                        valor = unicodedata.normalize('NFKD', valor).encode('ascii', 'ignore').decode('ascii')
                        return valor.strip().lower()

                    def rellenar_diccionarios_ordenados(
                            d1: dict, d2: dict, d3: dict, d4: dict,
                            nuevo_1: dict, nuevo_2: dict, nuevo_3: dict, nuevo_4: dict
                    ) -> Tuple[dict, dict, dict, dict]:

                        diccionarios = [d1, d2, d3, d4]
                        nuevos = [nuevo_1, nuevo_2, nuevo_3, nuevo_4]

                        # Mapa de valor normalizado → lista de (i, clave_original, valor_original)
                        valor_norm_a_origen = {}
                        for i, dic in enumerate(diccionarios):
                            for clave, valor in dic.items():
                                valor_norm = normalizar(valor)
                                if valor_norm not in valor_norm_a_origen:
                                    valor_norm_a_origen[valor_norm] = []
                                valor_norm_a_origen[valor_norm].append((i, clave, valor))  # mantenemos el valor real

                        # Primero: valores coincidentes
                        for _, origenes in valor_norm_a_origen.items():
                            if len(origenes) > 1:
                                for i, clave, valor in origenes:
                                    nuevos[i][clave] = valor

                        # Luego: valores únicos
                        for _, origenes in valor_norm_a_origen.items():
                            if len(origenes) == 1:
                                i, clave, valor = origenes[0]
                                nuevos[i][clave] = valor

                        return nuevos[0], nuevos[1], nuevos[2], nuevos[3]

                    # Paso 1: obtener los diccionarios
                    dict_es = {s['key']: s['value'] for s in parsed_specs_data["ES"]}
                    dict_uk = {s['key']: s['value'] for s in parsed_specs_data["UK"]}
                    dict_ita = {s['key']: s['value'] for s in parsed_specs_data["ITA"]}
                    dict_opt = {s['key']: s['value'] for s in parsed_specs_data["OPT"]}

                    # Paso 2: estructuras vacías ordenadas
                    ES, UK, ITA, OPT = {}, {}, {}, {}

                    # Paso 3: rellenar con los valores coincidentes primero
                    ES, UK, ITA, OPT = rellenar_diccionarios_ordenados(dict_es, dict_uk, dict_ita, dict_opt, ES, UK,
                                                                       ITA, OPT)

                    # Paso 4: Creamos estructura alineada por valor
                    grupo_por_valor = {}
                    for pais, dic in zip(["OPT", "UK", "ITA", "ES"], [OPT, UK, ITA, ES]):
                        for key, val in dic.items():
                            norm = normalizar(val)
                            if norm not in grupo_por_valor:
                                grupo_por_valor[norm] = {}
                            grupo_por_valor[norm][pais] = (key, val)

                    def limpiar_grupo_por_valor(grupo_por_valor: Dict[str, Dict[str, Tuple[str, any]]]) -> Dict[str, Dict[str, Tuple[str, any]]]:
                        grupo_limpio = {}

                        '''docs_dict_por_pais = fields_data.get(ClavesExcel.DOCUMENTOS.value, {})
                        docs_items = set()

                        for pais, dic in docs_dict_por_pais.items():
                            if isinstance(dic, dict):
                                for k, v in dic.items():
                                    if k and v:
                                        docs_items.add((k.strip(), v.strip()))'''

                        for norm_valor, subgrupo in grupo_por_valor.items():
                            nuevo_subgrupo = {}
                            ya_vistos = set()

                            for pais, (clave, valor) in subgrupo.items():
                                par = (clave.strip(), valor.strip())
                                #if par in docs_items: continue  # ❌ Ya existe como documento → omitimos
                                if par not in ya_vistos:
                                    nuevo_subgrupo[pais] = (clave, valor)
                                    ya_vistos.add(par)

                            if nuevo_subgrupo:
                                grupo_limpio[norm_valor] = nuevo_subgrupo

                        return grupo_limpio

                    grupo_por_valor = limpiar_grupo_por_valor(grupo_por_valor)

                    # Render alineado por valor compartido entre países
                    for valor_norm in grupo_por_valor:
                        fila_valores = grupo_por_valor[valor_norm]

                        html_parts.append("<tr>")  # ← nueva fila de tabla

                        for pais in ["", "OPT", "UK", "ITA", "ES"]:
                            html_parts.append(
                                f"<td class='country-val {pais.lower()}-val specifications-cell' data-source='{pais}'>")

                            if pais in fila_valores:
                                clave, valor = fila_valores[pais]

                                # Generar input checkbox si el valor está marcado
                                is_checked = any(
                                    cs['key'] == clave and cs['value'] == valor and cs['source_country'] == pais
                                    for cs in chosen_specs_from_data
                                )

                                safe_key = html.escape(clave).replace(' ', '_').replace(':', '-')
                                checkbox_id = f"cb-{html.escape(sku)}-{safe_field_name_for_js}-{pais}-{safe_key}"

                                html_parts.append("<ul class='specifications-sublist'>")
                                html_parts.append("<li>")
                                html_parts.append(f"<input type='checkbox' id='{checkbox_id}' "
                                                  f"name='specgroup-{html.escape(sku)}-{safe_field_name_for_js}-{safe_key}' "
                                                  f"data-sku='{html.escape(sku)}' data-field='{escaped_field_name}' "
                                                  f"data-spec-key='{html.escape(clave)}' data-spec-value='{html.escape(valor)}' data-source-country='{pais}' "
                                                  f"{'checked' if is_checked else ''} onchange='handleSpecCheckboxChange(this)'>")
                                html_parts.append(
                                    f"<label for='{checkbox_id}'><strong>{html.escape(clave)}:</strong> {html.escape(valor)}</label>")
                                html_parts.append("</li>")
                                html_parts.append("</ul>")
                            else:
                                html_parts.append("–")

                            html_parts.append("</td>")

                        # Columna final para visualización de selección (puede mantenerse en blanco o mostrar algo)
                        html_parts.append(
                            f"<td class='selection-display specifications-selection-display' id='selection-display-{html.escape(sku)}-{safe_field_name_for_js}'>"
                        )
                        if chosen_specs_from_data:
                            html_parts.append("Seleccionados:<ul class='specifications-chosen-list'>")
                            for cs in chosen_specs_from_data:
                                html_parts.append(
                                    f"<li><strong>{html.escape(cs['key'])}:</strong> {html.escape(cs['value'])} <em>({html.escape(cs['source_country'])})</em></li>")
                            html_parts.append("</ul>")
                        else:
                            html_parts.append("")#Ninguno seleccionado
                        html_parts.append("</td></tr>")

                elif field_name == ClavesExcel.COSTE.value: continue

                elif field_name in self.url_list_fields:
                    for country_code in COUNTRY_DISPLAY_ORDER:
                        current_country_urls_list = values.get(country_code, [])
                        #print(current_country_urls_list)
                        html_parts.append(
                            f"<td class='country-val {country_code.lower()}-val url-list-cell' data-source='{country_code}'>")
                        if current_country_urls_list:
                            html_parts.append("<ul class='url-list'>")
                            for key, url in current_country_urls_list.items():
                                link_text = html.escape(key)
                                html_parts.append(
                                    f"<li>{link_text}: <a href='{html.escape(url)}' target='_blank'>{url}</a></li>")
                            '''for idx, item_url in enumerate(current_country_urls_list):
                                if item_url:
                                    item_name = os.path.basename(
                                        item_url.split('?')[0]) if 'http' in item_url else item_url
                                    if not item_name or item_name == item_url: item_name = f"Enlace {idx + 1}"
                                    html_parts.append(
                                        f"<li><a href='{html.escape(item_url)}' target='_blank'>{html.escape(item_name)}</a></li>")'''
                            html_parts.append("</ul>")
                        html_parts.append("</td>")
                    html_parts.append("<td class='selection-group'>")
                    for src_country in COUNTRY_DISPLAY_ORDER:
                        is_ch = str(values['chosen_source']) == src_country
                        r_id = f"radio-{unique_input_group_name}-{src_country}"
                        html_parts.append(
                            f"<input type='radio' name='{unique_input_group_name}' value='{src_country}' id='{r_id}' {'checked' if is_ch else ''} data-sku='{html.escape(sku)}' data-field='{escaped_field_name}' onchange='recordSelection(this)'><label for='{r_id}'>{src_country}</label>")
                    html_parts.append("</td>")

                else:
                    for country_code_txt in COUNTRY_DISPLAY_ORDER:
                        current_val_txt = str(values[country_code_txt])
                        cell_html_content_txt = html.escape(current_val_txt)
                        html_parts.append(
                            f"<td class='country-val {country_code_txt.lower()}-val' data-source='{country_code_txt}'>{cell_html_content_txt}</td>")
                    html_parts.append("<td class='selection-group'>")
                    for src_country_txt in COUNTRY_DISPLAY_ORDER:
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
                    let footer = document.getElementById('save-button-container');
                    
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
                
                async function deduplicateImagesIfNeeded(sku) {{
                    const imageContainers = document.querySelectorAll(`[id^="image-container-${{sku}}-"][id$="-OPT"]`);
                    if (!imageContainers.length) return;
                
                    let imagesByCountry = {{ OPT: [], UK: [], ITA: [], ES: [] }};
                    for (let country of ["OPT", "UK", "ITA", "ES"]) {{
                        const imgs = document.querySelectorAll(`#image-container-${{sku}}-Galería-${{country}} img`); //Image_Urls o Galería?
                        imagesByCountry[country] = Array.from(imgs).map(img => img.src);
                    }}
                
                    const totalImages = Object.values(imagesByCountry).reduce((acc, list) => acc + list.length, 0);
                    if (totalImages === 0) return;  // Nada que hacer si no hay imágenes
                
                    try {{
                        const response = await fetch("http://localhost:5050/process_images", {{
                            method: "POST",
                            headers: {{ "Content-Type": "application/json" }},
                            body: JSON.stringify({{ sku, images_by_country: imagesByCountry }})
                        }});
                
                        const result = await response.json();
                        if (result.status === "ok") {{
                            for (let country of ["OPT", "UK", "ITA", "ES"]) {{
                                const container = document.getElementById(`image-container-${{sku}}-Galería-${{country}}`); //Image_Urls o Galería?
                                if (!container) continue;
                
                                container.innerHTML = "";
                                const selectedUrls = selections[sku]?.Galería || [];  //Image_Urls o Galería?
                
                                result.filtered[country].forEach((url, idx) => {{
                                    const cbId = `imgcb-${{sku}}-Galería-${{country}}-${{idx}}`; //Image_Urls o Galería?
                                    const isChecked = selectedUrls.includes(url);
                
                                    const inputEl = document.createElement("input");
                                    inputEl.type = "checkbox";
                                    inputEl.name = `sel-${{sku}}-Galería{{country}}`; //Image_Urls o Galería?
                                    inputEl.value = url;
                                    inputEl.id = cbId;
                                    inputEl.style.display = "none";
                                    inputEl.dataset.sku = sku;
                                    inputEl.dataset.field = "Galería";  // ✅ crucial para funcionamiento correcto //Image_Urls o Galería?
                                    if (isChecked) inputEl.checked = true;
                                    inputEl.onchange = () => recordSelection(inputEl);
                
                                    const imgEl = document.createElement("img");
                                    imgEl.src = url;
                                    imgEl.alt = `Img ${{idx + 1}}`;
                                    imgEl.onclick = () => toggleImageSelection(imgEl, cbId);
                                    if (isChecked) imgEl.classList.add("selected-image");
                
                                    container.appendChild(inputEl);
                                    container.appendChild(imgEl);
                
                                    if (isChecked) {{
                                        recordSelection(inputEl);  // ✅ actualiza la columna “Seleccionadas”
                                    }}
                                }});
                            }}
                        }}
                    }} catch (err) {{
                        console.warn("❌ Error en deduplicación de imágenes:", err);
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
                    
                    deduplicateImagesIfNeeded(skuPages[index].id.replace('sku-page-', ''));
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
                    
                    //Image_Urls o Galería?
                    if (field === 'Galería') {{
                        const cbs = document.querySelectorAll(`input[type="checkbox"][data-sku="${{sku}}"][data-field="Galería"]:checked`); //Image_Urls o Galería?
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
                
                function processPresenceSelections() {{
                    const skusToProcess = Object.entries(selections)
                        .filter(([sku, sel]) => sel["_presence"] === "scrap" || sel["_presence"] === "local")
                        .map(([sku, sel]) => ({{ sku, tipo: sel["_presence"] }}));
                
                    if (skusToProcess.length === 0) {{
                        alert("No hay selecciones de tipo scrap o local para procesar.");
                        return;
                    }}
                
                    fetch('http://localhost:5050/process_presence', {{
                        method: 'POST',
                        headers: {{ 'Content-Type': 'application/json' }},
                        body: JSON.stringify(skusToProcess)
                    }})
                    .then(res => res.json())
                    .then(result => {{
                        if (result.status === "success") {{
                            alert("Productos procesados correctamente.");
                            location.reload();
                        }} else {{
                            alert("Error al procesar: " + result.message);
                        }}
                    }})
                    .catch(err => {{
                        alert("Error de red: " + err.message);
                    }});
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
            self.load_excel()
            self.load_user_selections()
            self._identify_new_and_deleted_products()
            self._identify_all_skus_to_process()
            self._compare_products()
            self._generate_html(False)
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

    def apply_user_selections_to_excel(self):
        self.load_excel()
        #self.load_user_selections()#Cambiado
        if not self.user_selections:
            self._emit_progress("No hay selecciones de usuario para aplicar.")
            return

        id_col = self.sku_column_name
        hojas = {
            'ES': self.data_es.copy(),
            'UK': self.data_uk.copy(),
            'ITA': self.data_ita.copy(),
            'OPT': self.data_opt.copy()
        }

        for sku, campos_seleccionados in self.user_selections.items():
            sku = str(sku)
            if not sku: continue

            for campo, origen in campos_seleccionados.items():
                if campo == "_presence":
                    continue  # Esto se maneja por separado en otro proceso

                # --- CASO 1: Arrays de URLs / imágenes ---
                if isinstance(origen, (list, dict)):
                    if campo in self.url_list_fields or campo == ClavesExcel.GALERIA.value:
                        val = json.dumps(origen, ensure_ascii=False)
                        for nombre_hoja, df in hojas.items():
                            if sku in df[id_col].astype(str).values and campo in df.columns:
                                hojas[nombre_hoja].loc[df[id_col].astype(str) == sku, campo] = val
                    continue

                # --- CASO 2: Campos estructurados (specifications, etc.) ---
                if isinstance(origen, list):
                    # Convertimos la lista de dicts en un solo dict plano
                    #resultado = {item["key"]: item["value"] for item in origen if "key" in item and "value" in item}
                    resultado = {
                        str(item["key"]): item["value"]
                        for item in origen
                        if isinstance(item, dict) and "key" in item and "value" in item
                    }
                    val = json.dumps(resultado, ensure_ascii=False)

                    for nombre_hoja, df in hojas.items():
                        if sku in df[id_col].astype(str).values and campo in df.columns:
                            hojas[nombre_hoja].loc[df[id_col].astype(str) == sku, campo] = val
                    continue

                # --- CASO 3: Campos normales seleccionados de una hoja ---
                hoja_origen = hojas.get(origen)
                if hoja_origen is None or campo not in hoja_origen.columns:
                    self._emit_progress(f"⚠ No se encontró la hoja '{origen}' o el campo '{campo}' para el SKU '{sku}'")
                    continue

                val = ''
                fila = hoja_origen[hoja_origen[id_col].astype(str) == sku]
                if not fila.empty:
                    val = fila[campo].values[0]

                # Aplicar valor a todas las hojas donde aparezca el SKU
                for nombre_hoja, df in hojas.items():
                    if sku in df[id_col].astype(str).values and campo in df.columns:
                        hojas[nombre_hoja].loc[df[id_col].astype(str) == sku, campo] = val

        # Guardar todas las hojas en el Excel original
        try:
            with pd.ExcelWriter(self.excel_path, engine='openpyxl', mode='w') as writer:
                for nombre_hoja, df in hojas.items():
                    df.to_excel(writer, sheet_name=nombre_hoja, index=False)
            self._emit_progress("✔ Se aplicaron las selecciones del usuario al Excel con múltiples hojas.")
        except Exception as e:
            self._emit_progress(f"❌ Error al guardar el Excel: {str(e)}")

    def apply_direct_edits_to_excel(self):
        """
        Aplica cambios directos al Excel usando self.user_selections.
        Los valores ya son definitivos: string, lista o diccionario.
        """
        edits = self.user_selections
        self.load_excel()

        if not edits:
            self._emit_progress("No hay ediciones directas para aplicar.")
            return

        id_col = self.sku_column_name
        hojas = {
            'ES': self.data_es.copy(),
            'UK': self.data_uk.copy(),
            'ITA': self.data_ita.copy(),
            'OPT': self.data_opt.copy()
        }

        for sku, campos_editados in edits.items():
            sku = str(sku)
            if not sku:
                continue

            for campo, nuevo_valor in campos_editados.items():
                # Serializar listas o diccionarios
                if isinstance(nuevo_valor, (list, dict)):
                    val_str = json.dumps(nuevo_valor, ensure_ascii=False)
                else:
                    val_str = str(nuevo_valor) if nuevo_valor is not None else ''

                for nombre_hoja, df in hojas.items():
                    if sku in df[id_col].astype(str).values and campo in df.columns:
                        hojas[nombre_hoja].loc[df[id_col].astype(str) == sku, campo] = val_str

        # Duplicar productos nuevos editados a las demás hojas
        for sku in edits.keys():
            sku = str(sku)
            hojas_presencia = [
                nombre_hoja for nombre_hoja, df in hojas.items()
                if sku in df[id_col].astype(str).values
            ]

            if len(hojas_presencia) == 1:
                hoja_origen = hojas_presencia[0]
                fila_origen = hojas[hoja_origen][hojas[hoja_origen][id_col].astype(str) == sku]

                for destino in hojas.keys():
                    if destino != hoja_origen:
                        hojas[destino] = pd.concat([hojas[destino], fila_origen], ignore_index=True)
                self._emit_progress(f"✔ SKU {sku} copiado desde {hoja_origen} al resto de hojas.")

        try:
            with pd.ExcelWriter(self.excel_path, engine='openpyxl', mode='w') as writer:
                for nombre_hoja, df in hojas.items():
                    df.to_excel(writer, sheet_name=nombre_hoja, index=False)
            self._emit_progress("✔ Ediciones directas aplicadas al Excel.")
        except Exception as e:
            self._emit_progress(f"❌ Error al guardar Excel con ediciones directas: {str(e)}")
