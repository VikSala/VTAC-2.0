import base64
import pandas as pd
from utils import Utils

'''def traducir_columnas_excel_filtradas(
    ruta_excel,
    hoja='Products',
    columnas_a_traducir=('Nombre', 'Descripción Web', 'Atributos', 'Categoría'),
    columna_filtro='Atributos',
    patron='"Sku":',
    idioma_destino='es',
    delay_segundos=1.5
):
    from deep_translator import GoogleTranslator
    print(f"📄 Abriendo archivo: {ruta_excel}, hoja: {hoja}")
    df = pd.read_excel(ruta_excel, sheet_name=hoja)

    # Verificamos que las columnas existan
    columnas_faltantes = [col for col in columnas_a_traducir if col not in df.columns]
    if columnas_faltantes:
        print(f"❌ Faltan las siguientes columnas: {columnas_faltantes}")
        return

    traductor = GoogleTranslator(source='auto', target=idioma_destino)

    # Filtrar filas que contienen el patrón en la columna filtro
    indices_a_traducir = df[df[columna_filtro].astype(str).str.contains(patron, na=False)].index
    print(f"🔎 Se encontraron {len(indices_a_traducir)} filas con el patrón '{patron}' en '{columna_filtro}'.")

    for i, idx in enumerate(indices_a_traducir, 1):
        print(f"🔁 Traduciendo fila {i}/{len(indices_a_traducir)} (índice {idx})...")
        for col in columnas_a_traducir:
            original = str(df.at[idx, col]) if pd.notna(df.at[idx, col]) else ""
            try:
                traducido = traductor.translate(original)
                df.at[idx, col] = traducido
                print(f"  - {col}: {original[:30]} → {traducido[:30]}")
            except Exception as e:
                print(f"  ⚠️ Error traduciendo columna {col} en fila {idx}: {e}")
        time.sleep(delay_segundos)

    # Guardar sobrescribiendo la hoja original
    with pd.ExcelWriter(ruta_excel, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=hoja, index=False)

    print(f"✅ Traducción completada. Cambios guardados en '{ruta_excel}', hoja '{hoja}'.")

traducir_columnas_excel_filtradas(
    ruta_excel=ruta,
    hoja='Sheet1',
    idioma_destino='es',
    delay_segundos=2
)'''

'''def comparar_imagenes_por_url(url1, url2, umbral_similitud=5):
    """
    Compara dos imágenes por URL usando perceptual hash (phash).
    Imprime el valor de diferencia y el grado de similitud.
    umbral_similitud: valores típicos
        - 0: exactamente iguales
        - 1-5: muy parecidas
        - 6+: diferentes
    """
    from PIL import Image
    import imagehash
    import requests
    from io import BytesIO

    def descargar_y_hash(url):
        if "http" not in url:
            url = "http://143.47.53.74:8070/" + url
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        img = Image.open(BytesIO(response.content))
        return imagehash.phash(img)

    try:
        h1 = descargar_y_hash(url1)
        h2 = descargar_y_hash(url2)
        diferencia = h1 - h2

        print(f"Diferencia de hash: {diferencia}")

        if diferencia == 0:
            print(f"✅ Las imágenes son IGUALES: {diferencia}")
        elif diferencia <= umbral_similitud:
            print(f"🟡 Las imágenes son PARECIDAS: {diferencia}")
        else:
            print(f"❌ Las imágenes son DIFERENTES: {diferencia}")

    except Exception as e:
        print(f"⚠️ Error comparando imágenes: {e}")
'''

ruta = Utils.seleccionar_excel()


def rellenar_valor_y_url_debug(ruta_excel):
    hojas = pd.read_excel(ruta_excel, sheet_name=None)
    df_todos = hojas["TODOS"]
    hojas_datos = {k: v for k, v in hojas.items() if k in ["ES", "UK", "ITA", "OPT"]}

    valores = []
    urls = []

    for idx, row in df_todos.iterrows():
        valor = ""
        url = ""
        print(f"\n--- Procesando fila {idx} ---")

        try:
            dic = ast.literal_eval(row["UNICOS"])
            if isinstance(dic, dict):
                # Obtener primer item del diccionario
                hoja_key, atributo = next(iter(dic.items()))
                print(f"Hoja destino: {hoja_key}, atributo buscado: {atributo}")

                if hoja_key in hojas_datos:
                    df_hoja = hojas_datos[hoja_key]

                    # Buscar fila que contenga ese atributo como clave en el dict de Specifications
                    for fila_idx, fila in df_hoja.iterrows():
                        specs_raw = fila.get("Specifications", "")
                        try:
                            specs = ast.literal_eval(specs_raw) if isinstance(specs_raw, str) else {}
                        except Exception as e:
                            print(f"Error parseando Specifications: {e}")
                            continue

                        if isinstance(specs, dict) and atributo in specs:
                            valor = specs[atributo]
                            url = fila.get("x_url", "")
                            print(f"✔ Encontrado en fila {fila_idx} de hoja {hoja_key}")
                            print(f"VALOR = {valor}")
                            print(f"URL = {url}")
                            break
                    else:
                        print(f"✘ No se encontró '{atributo}' en Specifications de hoja {hoja_key}")
                else:
                    print(f"✘ Hoja '{hoja_key}' no encontrada.")
            else:
                print("✘ UNICOS no es un diccionario válido.")
        except Exception as e:
            print(f"✘ Error evaluando UNICOS: {e}")

        valores.append(valor)
        urls.append(url)

    # Añadir resultados a la hoja TODOS
    df_todos["VALOR"] = valores
    df_todos["URL"] = urls

    # Guardar archivo actualizado
    salida = ruta_excel.replace(".xlsx", "_con_valores_y_urls_DEBUG.xlsx")
    with pd.ExcelWriter(salida, engine='openpyxl') as writer:
        for hoja, df in hojas.items():
            if hoja == "TODOS":
                df_todos.to_excel(writer, sheet_name=hoja, index=False)
            else:
                df.to_excel(writer, sheet_name=hoja, index=False)

    print(f"\n✅ Archivo guardado: {salida}")
    return salida

from pathlib import Path

def diff_unique_simple(
    excel_file: str | Path,
    origen: str | int | None = None,   # hoja con las columnas NEW y OLD
    hoja_diff: str = "DIFF",           # nombre de la hoja de salida
    salida: str | Path | None = None   # archivo resultante (None = sobrescribe)
) -> None:
    """
    Crea/actualiza una hoja `hoja_diff` con todos los valores únicos que
    aparecen solo en NEW o solo en OLD (comparación exacta, sin repetidos).
    """
    excel_file = Path(excel_file)
    salida = excel_file if salida is None else Path(salida)

    # 1) Leer datos y validar columnas
    df = pd.read_excel(excel_file, sheet_name=origen)
    if {"NEW", "OLD"} - set(df.columns):
        raise ValueError("La hoja debe contener las columnas 'NEW' y 'OLD'.")

    # 2) Diferencia simétrica (valores únicos)
    new_vals = set(df["NEW"].dropna().astype(str))
    old_vals = set(df["OLD"].dropna().astype(str))
    diff_vals = sorted(new_vals ^ old_vals)          # ⊕ = diferencia simétrica

    diff_df = pd.DataFrame({"DIFF": diff_vals})

    # 3) Guardar la hoja DIFF (manejo de append/replace)
    modo = "a" if salida.exists() else "w"
    writer_kwargs = dict(path=salida, engine="openpyxl", mode=modo)
    if modo == "a":
        writer_kwargs["if_sheet_exists"] = "replace"

    with pd.ExcelWriter(**writer_kwargs) as writer:
        diff_df.to_excel(writer, sheet_name=hoja_diff, index=False)

    print(f"Hoja '{hoja_diff}' escrita con {len(diff_df)} valor(es) únicos.")

    # --- OPCIONAL: añadir DIFF a la misma hoja ORIGINAL -----------------
    # Si en lugar de una hoja nueva prefirieras tener la columna DIFF
    # junto a NEW y OLD, bastaría con esto (y luego volver a guardar `df`):
    #
    # df["DIFF"] = df.apply(
    #     lambda r: r["NEW"] if r["NEW"] not in old_vals
    #               else (r["OLD"] if r["OLD"] not in new_vals else pd.NA),
    #     axis=1
    # )
    # df.to_excel(salida, sheet_name=origen, index=False)


def transferir_videos_documentos(excel_path):
    # Cargar ambas hojas
    df1 = pd.read_excel(excel_path, sheet_name='Sheet1')
    df2 = pd.read_excel(excel_path, sheet_name='Sheet2')

    # Nos aseguramos de que las columnas SKU sean de tipo string (por si acaso hay diferencias de tipo)
    df1['SKU'] = df1['SKU'].astype(str).str.strip()
    df2['SKU'] = df2['SKU'].astype(str).str.strip()

    # Seleccionamos solo las columnas necesarias de Sheet2
    df2_reducido = df2[['SKU', 'Vídeos', 'Documentos']]

    # Hacemos la unión (merge) usando SKU como clave
    df_actualizado = df1.merge(df2_reducido, on='SKU', how='left', suffixes=('', '_sheet2'))

    # Si las columnas ya existen en df1, las sobrescribimos solo si encontramos valores nuevos
    for col in ['Vídeos', 'Documentos']:
        if f'{col}_sheet2' in df_actualizado.columns:
            df_actualizado[col] = df_actualizado[f'{col}_sheet2'].combine_first(df_actualizado[col])
            df_actualizado.drop(columns=[f'{col}_sheet2'], inplace=True)

    return df_actualizado


def genera_sheet3_unicos(path_excel: str,
                         sheet1: str = "Sheet1",
                         sheet2: str = "Sheet2",
                         salida: str = "Sheet3",
                         col_sku: str = "SKU") -> None:
    # Leer la columna SKU de cada hoja
    s1 = (
        pd.read_excel(path_excel, sheet_name=sheet1, usecols=[col_sku])
        [col_sku]
        .dropna()
        .astype(str)
    )
    s2 = (
        pd.read_excel(path_excel, sheet_name=sheet2, usecols=[col_sku])
        [col_sku]
        .dropna()
        .astype(str)
    )

    # SKU que están en Sheet1 pero no en Sheet2
    exclusivos = (
        s1[~s1.isin(s2)]
        .drop_duplicates()
        .reset_index(drop=True)
        .to_frame(name=col_sku)  # convertir en DataFrame para exportar
    )

    # Escribirlos en Sheet3 (sobrescribe si ya existe)
    with pd.ExcelWriter(path_excel, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
        exclusivos.to_excel(writer, sheet_name=salida, index=False)

    print(f"✅ {len(exclusivos)} SKU(s) exclusivos de '{sheet1}' escritos en '{salida}'")


def duplica_sku_a_hoja(path_excel: str,
                       hoja_origen: str = "Sheet1",
                       hoja_salida: str = "Duplicados",
                       col_sku: str = "SKU") -> None:
    """
    Detecta SKU duplicados en `hoja_origen` y los escribe, con su nº de repeticiones,
    en la hoja `hoja_salida` del mismo libro.
    """
    # 1 ─ Leer únicamente la columna SKU
    df = pd.read_excel(path_excel, sheet_name=hoja_origen, usecols=[col_sku])
    df.columns = df.columns.str.strip()          # quita espacios alrededor del encabezado

    # 2 ─ Contar repeticiones y filtrar valores > 1
    dupes = (
        df[col_sku]
        .dropna()
        .astype(str)
        .value_counts()                          # Series: índice = SKU, valor = repeticiones
        .loc[lambda s: s > 1]                    # solo los duplicados
        .rename_axis(col_sku)                    # pone nombre al índice
        .reset_index(name="Repeticiones")        # lo convierte en DataFrame con dos columnas
        .sort_values(col_sku, kind="stable")     # orden alfabético por SKU
    )

    if dupes.empty:
        print("✅ No se encontraron SKU duplicados.")
        return

    # 3 ─ Escribir la hoja de salida (la reemplaza si ya existe)
    with pd.ExcelWriter(path_excel, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
        dupes.to_excel(writer, sheet_name=hoja_salida, index=False)

    print(f"✅ {len(dupes)} SKU duplicado(s) escritos en la hoja '{hoja_salida}'")


def crear_hoja_resultado(excel_path, hoja_resultado="Resultado"):
    """
    Lee un Excel con hojas: Madrid, Bulgaria, Odoo, VS, VD, VN
    y crea la hoja 'Resultado' con las columnas: VS, VN, VD, VZ,
    Novedades VN, Novedades VS, Novedades VD, Novedades VZ.
    """
    excel_path = Path(excel_path)

    # --- Helpers ---
    def norm_sku_series(s):
        if s is None:
            return []
        s = s.astype(str).str.strip()
        s = s[s.ne("") & s.ne("nan")]
        return list(dict.fromkeys(s.tolist()))

    def to_num(series):
        return pd.to_numeric(series, errors="coerce")

    # --- Leer hojas ---
    xls = pd.ExcelFile(excel_path)
    VS_df = pd.read_excel(xls, "VS", dtype={"SKU": str}, usecols=["SKU"])
    VD_df = pd.read_excel(xls, "VD", dtype={"SKU": str}, usecols=["SKU"])
    VN_df = pd.read_excel(xls, "VN", dtype={"SKU": str}, usecols=["SKU"])
    Madrid_df  = pd.read_excel(xls, "Madrid", dtype={"SKU": str}, usecols=["SKU", "STOCK"])
    Odoo_df    = pd.read_excel(xls, "Odoo",  dtype={"SKU": str}, usecols=["SKU", "STOCK"])
    Bulgaria_df = pd.read_excel(
        xls, "Bulgaria", dtype={"SKU": str}, usecols=["SKU", "STOCK", "UNDELIVERED ORDER"]
    )

    # --- Normalizar ---
    VS_skus = norm_sku_series(VS_df["SKU"])
    VD_skus = norm_sku_series(VD_df["SKU"])
    VN_skus = norm_sku_series(VN_df["SKU"])

    for df in (Madrid_df, Odoo_df, Bulgaria_df):
        df["SKU"] = df["SKU"].astype(str).str.strip()

    Madrid_df["STOCK"]   = to_num(Madrid_df["STOCK"])
    Odoo_df["STOCK"]     = to_num(Odoo_df["STOCK"])
    Bulgaria_df["STOCK"] = to_num(Bulgaria_df["STOCK"])

    bul_undel = Bulgaria_df["UNDELIVERED ORDER"].astype(str).str.strip()
    bul_undel = bul_undel.where(~bul_undel.isin(["", "nan"]), other=pd.NA)

    # --- VS final = VS - VD ---
    VD_set = set(VD_skus)
    VS_final = [sku for sku in VS_skus if sku not in VD_set]

    # --- VN final = VN - VS ---
    VS_set = set(VS_skus)
    VN_final = [sku for sku in VN_skus if sku not in VS_set]

    # --- Novedades ---
    Odoo_skus = set(norm_sku_series(Odoo_df["SKU"]))
    Novedades_VN = [sku for sku in VN_final if sku not in Odoo_skus]
    # el resto se calculan más abajo tras VD/VZ

    # --- VD/VZ ---
    madrid_stock = dict(zip(Madrid_df["SKU"], Madrid_df["STOCK"]))
    odoo_stock   = dict(zip(Odoo_df["SKU"],   Odoo_df["STOCK"]))
    bul_stock    = dict(zip(Bulgaria_df["SKU"], Bulgaria_df["STOCK"]))
    bul_undel_map = dict(zip(Bulgaria_df["SKU"], bul_undel))

    VD_final, VZ_final = [], []


    for sku in VD_skus:
        enviar_a_vz = True

        # Bulgaria
        b_stock = bul_stock.get(sku, pd.NA)
        b_undel = bul_undel_map.get(sku, pd.NA)
        stock_cond = (pd.isna(b_stock) or (pd.notna(b_stock) and b_stock < 1))
        undel_cond = pd.isna(b_undel)
        if stock_cond and undel_cond:
            enviar_a_vz = True
        else: enviar_a_vz = False

        # Madrid
        m_stock = madrid_stock.get(sku, pd.NA)
        if pd.notna(m_stock) and m_stock > 0:
            enviar_a_vz = False

        # Odoo
        o_stock = odoo_stock.get(sku, pd.NA)
        if not (pd.notna(o_stock) and o_stock < 1):
            enviar_a_vz = False

        (VZ_final if enviar_a_vz else VD_final).append(sku)

    # --- Novedades para todas las listas ---
    Novedades_VS = [sku for sku in VS_final if sku not in Odoo_skus]
    Novedades_VD = [sku for sku in VD_final if sku not in Odoo_skus]
    Novedades_VZ = [sku for sku in VZ_final if sku not in Odoo_skus]

    # --- Construcción hoja Resultado ---
    max_len = max(
        len(VS_final), len(VN_final), len(VD_final), len(VZ_final),
        len(Novedades_VN), len(Novedades_VS), len(Novedades_VD), len(Novedades_VZ)
    )

    def pad(lst, n): return lst + [None] * (n - len(lst))

    resultado_df = pd.DataFrame({
        "VS": pad(VS_final, max_len),
        "VN": pad(VN_final, max_len),
        "VD": pad(VD_final, max_len),
        "VZ": pad(VZ_final, max_len),
        "Novedades VN": pad(Novedades_VN, max_len),
        "Novedades VS": pad(Novedades_VS, max_len),
        "Novedades VD": pad(Novedades_VD, max_len),
        "Novedades VZ": pad(Novedades_VZ, max_len),
    })

    # --- Escribir/actualizar hoja Resultado ---
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        resultado_df.to_excel(writer, sheet_name=hoja_resultado, index=False)

    return {
        "VS": VS_final, "VN": VN_final, "VD": VD_final, "VZ": VZ_final,
        "Novedades VN": Novedades_VN,
        "Novedades VS": Novedades_VS,
        "Novedades VD": Novedades_VD,
        "Novedades VZ": Novedades_VZ,
    }


#region Fusion
import ast
import json
from typing import Dict, Any, Tuple, List, Set


def merge_atributos_prioridad(excel_path: str, output_path: str = None) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, List[str]]]:
    """
    Fusiona 'Atributos' entre hojas ES, ITA, UK con prioridad ES > ITA > UK,
    y aplica el orden de atributos (General → Datos del producto → Información adicional).

    Lógica de fusión (por SKU existente en la hoja):
      - Para ES: añade claves que existan en ITA/UK y no existan en ES.
                 Luego SOLO esas claves añadidas se sincronizan hacia ITA y UK (si el SKU existe allí).
      - Para ITA (SKUs no presentes en ES): añade claves desde UK que no existan en ITA.
                 Luego SOLO esas añadidas se sincronizan hacia UK.
      - Para UK (SKUs no presentes en ES ni ITA): no hay hojas de menor prioridad.

    No crea filas nuevas; solo modifica 'Atributos' donde ya exista el SKU.
    Devuelve (df_es, df_ita, df_uk, log_cambios).
    """
    # === ORDEN DE ATRIBUTOS =======================================================
    PRODUCT_DATA_ORDER = [
        "SKU",
        "Código de producto",
        "EAN Código",
        "Peso del artículo",
        "Volumen del artículo",
        "Unidades por caja",
    ]

    ADDITIONAL_INFO_ORDER = [
        "Marca",
        "Garantía",
        "Comentarios",
        "Certificados",
        "Clase de eficiencia energética",
        "Rendimiento energético",
        "Dimensiones gráficos",
    ]

    PRODUCT_DATA_SET = set(PRODUCT_DATA_ORDER)
    ADDITIONAL_INFO_SET = set(ADDITIONAL_INFO_ORDER)

    # === HELPERS =================================================================
    def _parse_attr(cell) -> Dict[str, Any]:
        """Convierte la celda (string dict / json) a dict; vacío si NaN o error."""
        if isinstance(cell, dict):
            return cell
        if not isinstance(cell, str) or not cell.strip():
            return {}
        try:
            return ast.literal_eval(cell)
        except Exception:
            try:
                return json.loads(cell)
            except Exception:
                return {}

    def _dump_attr(d: Dict[str, Any]) -> str:
        """Serializa dict a string; preserva acentos."""
        return json.dumps(d, ensure_ascii=False)

    def _ensure_columns(df: pd.DataFrame, sheet_name: str):
        if "SKU" not in df.columns or "Atributos" not in df.columns:
            raise ValueError(f"❌ La hoja '{sheet_name}' debe tener columnas 'SKU' y 'Atributos'.")

    def _build_index(df: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
        """Devuelve {sku: atributos_dict}."""
        idx = {}
        for _, row in df.iterrows():
            sku = str(row["SKU"]).strip()
            if not sku:
                continue
            idx[sku] = _parse_attr(row["Atributos"])
        return idx

    def _order_attributes(attrs: Dict[str, Any]) -> Dict[str, Any]:
        """
        Orden:
          1) Información general (cualquier clave no listada) [mantiene orden de aparición]
          2) Datos del producto (PRODUCT_DATA_ORDER)
          3) Información adicional (ADDITIONAL_INFO_ORDER)
        """
        # General = todo lo que no pertenezca a las listas definidas
        general_keys = [k for k in attrs.keys() if k not in PRODUCT_DATA_SET and k not in ADDITIONAL_INFO_SET]
        ordered = {}

        # 1) General (en su orden actual)
        for k in general_keys:
            ordered[k] = attrs[k]

        # 2) Datos del producto (en el orden explícito)
        for k in PRODUCT_DATA_ORDER:
            if k in attrs:
                ordered[k] = attrs[k]

        # 3) Información adicional (en el orden explícito)
        for k in ADDITIONAL_INFO_ORDER:
            if k in attrs:
                ordered[k] = attrs[k]

        return ordered

    def _write_back(df: pd.DataFrame, attr_index: Dict[str, Dict[str, Any]]):
        """Vuelca (ordenado) el dict a la columna 'Atributos' del df."""
        # Pre-ordena todos los dicts para mantener consistencia
        ordered_idx = {sku: _order_attributes(attrs) for sku, attrs in attr_index.items()}
        serialized = {sku: _dump_attr(attrs) for sku, attrs in ordered_idx.items()}

        def _serialize_row(row):
            sku = str(row["SKU"]).strip()
            if not sku:
                return _dump_attr({})
            # Si el SKU no está en el índice (fila residual), ordena lo que hay
            if sku not in serialized:
                raw = _parse_attr(row["Atributos"])
                raw = _order_attributes(raw)
                return _dump_attr(raw)
            return serialized[sku]

        df["Atributos"] = df.apply(_serialize_row, axis=1)

    # === FUNCIÓN PRINCIPAL ========================================================
    # Leer hojas
    es_df  = pd.read_excel(excel_path, sheet_name="ES")
    ita_df = pd.read_excel(excel_path, sheet_name="IT")
    uk_df  = pd.read_excel(excel_path, sheet_name="UK")

    # Validación
    _ensure_columns(es_df, "ES")
    _ensure_columns(ita_df, "IT")
    _ensure_columns(uk_df, "UK")

    # Índices {sku: dict}
    es_idx  = _build_index(es_df)
    ita_idx = _build_index(ita_df)
    uk_idx  = _build_index(uk_df)

    present_sets = {
        "ES": set(es_idx.keys()),
        "IT": set(ita_idx.keys()),
        "UK": set(uk_idx.keys()),
    }
    processed_skus: Set[str] = set()
    log_cambios: Dict[str, List[str]] = {"ES": [], "IT": [], "UK": []}

    def _merge_from_lower(base_name: str, base_idx: Dict[str, Dict[str, Any]],
                          lower_name: str, lower_idx: Dict[str, Dict[str, Any]],
                          sku: str):
        """
        Para un SKU:
          - Añade a 'base' las claves presentes en 'lower' y ausentes en 'base'.
          - Propaga SOLO esas claves añadidas hacia 'lower' (para alinear).
          - Registra en log las claves añadidas/sincronizadas.
        """
        base_attrs  = base_idx.get(sku, {})
        lower_attrs = lower_idx.get(sku)
        if lower_attrs is None:
            return

        added_keys = []
        for k, v in lower_attrs.items():
            if k not in base_attrs:
                base_attrs[k] = v
                added_keys.append(k)

        if added_keys:
            # Propaga SOLO las recién añadidas
            for k in added_keys:
                lower_attrs[k] = base_attrs[k]

            # Vuelve a guardar (sin ordenar aún; el orden se aplica al volcar)
            base_idx[sku]  = base_attrs
            lower_idx[sku] = lower_attrs

            log_cambios[base_name].append(f"{sku}: añadidas desde {lower_name} -> {added_keys}")
            log_cambios[lower_name].append(f"{sku}: sincronizadas desde {base_name} -> {added_keys}")

        # 1) ES como base: comparar con ITA y UK
        for sku in sorted(present_sets["ES"]):
            # Añadimos primero desde ITA (más prioridad que UK) y luego desde UK
            _merge_from_lower("ES", es_idx, "IT", ita_idx, sku)
            _merge_from_lower("ES", es_idx, "UK", uk_idx, sku)
            processed_skus.add(sku)

        # 2) ITA como base para SKUs no procesados: comparar solo con UK
        for sku in sorted(present_sets["IT"] - processed_skus):
            _merge_from_lower("IT", ita_idx, "ES", es_idx, sku)
            _merge_from_lower("IT", ita_idx, "UK", uk_idx, sku)
            processed_skus.add(sku)

        # 3) UK como base para SKUs no procesados: no hay hojas de menor prioridad
        # (nada que hacer salvo marcarlos como procesados)
        for sku in sorted(present_sets["UK"] - processed_skus):
            processed_skus.add(sku)

    # Volcado ordenado
    _write_back(es_df,  es_idx)
    _write_back(ita_df, ita_idx)
    _write_back(uk_df,  uk_idx)

    # Guardar si se pide
    if output_path:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            es_df.to_excel(writer,  index=False, sheet_name="ES")
            ita_df.to_excel(writer, index=False, sheet_name="IT")
            uk_df.to_excel(writer,  index=False, sheet_name="UK")

    return es_df, ita_df, uk_df, log_cambios

# Funde atributos priorizando ES > IT > UK
#merge_atributos_prioridad(ruta, output_path=ruta)

from typing import Optional

def merge_es_ita_uk_to_all(excel_path: str, output_path: Optional[str] = None):
    """
    Crea/actualiza las hojas:
      - ALL: ES + (ITA sin SKUs ya presentes) + (UK sin SKUs ya presentes)
      - NEW: copia idéntica al resultado de ALL en esta ejecución
    Si ya existía ALL en el archivo de entrada:
      - OLD: copia del ALL anterior

    Conserva el resto de hojas originales.
    """
    # --- Cargar todas las hojas ---
    xls = pd.ExcelFile(excel_path)
    sheet_names = set(xls.sheet_names)
    sheets = {name: xls.parse(name) for name in xls.sheet_names}

    # --- Validaciones básicas ---
    required = {"ES", "IT", "UK"}
    if not required.issubset(sheet_names):
        missing = ", ".join(sorted(required - sheet_names))
        raise ValueError(f"❌ Faltan hojas requeridas: {missing}")

    for name in ("ES", "IT", "UK"):
        if "SKU" not in sheets[name].columns:
            raise ValueError(f"❌ La hoja '{name}' no contiene la columna 'SKU'.")

    # --- Normalización SKU ---
    def _norm_sku(s):
        if pd.isna(s):
            return ""
        return str(s).strip()

    es  = sheets["ES"].copy()
    ita = sheets["IT"].copy()
    uk  = sheets["UK"].copy()

    es["__SKU__"]  = es["SKU"].apply(_norm_sku)
    ita["__SKU__"] = ita["SKU"].apply(_norm_sku)
    uk["__SKU__"]  = uk["SKU"].apply(_norm_sku)

    es  = es[es["__SKU__"]  != ""].copy()
    ita = ita[ita["__SKU__"] != ""].copy()
    uk  = uk[uk["__SKU__"]  != ""].copy()

    # --- Construcción de ALL (prioridad ES > ITA > UK) ---
    all_df = es.copy()
    seen = set(all_df["__SKU__"].tolist())

    ita_new = ita[~ita["__SKU__"].isin(seen)].copy()
    all_df = pd.concat([all_df, ita_new], ignore_index=True, sort=False)
    seen.update(ita_new["__SKU__"].tolist())

    uk_new = uk[~uk["__SKU__"].isin(seen)].copy()
    all_df = pd.concat([all_df, uk_new], ignore_index=True, sort=False)
    seen.update(uk_new["__SKU__"].tolist())

    # Limpiar columna auxiliar
    for df in (all_df, es, ita, uk):
        if "__SKU__" in df.columns:
            df.drop(columns=["__SKU__"], inplace=True)

    # --- Preparar hojas a escribir ---
    # NEW: igual que el ALL resultante
    new_df = all_df.copy()

    # OLD: si existe ALL original, lo guardamos
    old_df = sheets["ALL"].copy() if "ALL" in sheets else None

    # --- Construir listado de salida preservando hojas originales ---
    # Empezamos con todas las hojas originales
    out_sheets = {name: df.copy() for name, df in sheets.items()}

    # Sobrescribir siempre ES/ITA/UK con su versión limpiada (respetando contenido)
    out_sheets["ES"]  = es
    out_sheets["IT"] = ita
    out_sheets["UK"]  = uk

    # OLD si había ALL
    if old_df is not None:
        out_sheets["OLD"] = old_df

    # ALL (nuevo) y NEW (nuevo)
    out_sheets["ALL"] = all_df
    out_sheets["NEW"] = new_df

    # Evitar duplicados molestos en nombres (por si existían OLD/NEW previas)
    # Simplemente sobrescribimos.

    # --- Guardado ---
    if output_path is None:
        if excel_path.lower().endswith(".xlsx"):
            output_path = excel_path[:-5] + "_with_ALL_NEW_OLD.xlsx"
        else:
            output_path = excel_path + "_with_ALL_NEW_OLD.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        for name, df in out_sheets.items():
            df.to_excel(writer, index=False, sheet_name=name)

    # --- Métricas ---
    print(f"✅ Guardado en: {output_path}")
    print(f"   Filas ES: {len(es)}, ITA añadidas: {len(ita_new)}, UK añadidas: {len(uk_new)}, TOTAL ALL/NEW: {len(all_df)}")
    if old_df is not None:
        print(f"   OLD creado a partir del ALL anterior ({len(old_df)} filas).")

    return out_sheets["ALL"], out_sheets["NEW"], out_sheets.get("OLD")

#merge_es_ita_uk_to_all(ruta, ruta)
#endregion

def norm_atr(excel):
    import utils_excel
    utils_excel.normalizar_atributos_excel(excel, "it", excel)

import pandas as pd
from pathlib import Path

def fusionar_all_odoo(excel_path, hoja_all="ALL", hoja_odoo="ODOO",
                      hoja_fusion="FUSION", hoja_novedades="NOVEDADES"):
    """
    - Vuelca todo ODOO a FUSION.
    - Para SKUs coincidentes entre ALL y ODOO, copia 'Atributos' de ALL hacia FUSION
      (solo si en ALL hay dato).
    - NOVEDADES: filas de ALL cuyo SKU no aparece en ODOO.
    No crea columnas auxiliares en los DataFrames.
    """
    excel_path = Path(excel_path)
    xls = pd.ExcelFile(excel_path)

    df_all  = pd.read_excel(xls, hoja_all)
    df_odoo = pd.read_excel(xls, hoja_odoo)

    if "SKU" not in df_all.columns or "SKU" not in df_odoo.columns:
        raise ValueError("Ambas hojas deben tener la columna 'SKU'.")

    # Helper de normalización (no toca columnas originales)
    def norm_series(s: pd.Series) -> pd.Series:
        return (
            s.astype(str)
             .str.strip()
             .str.upper()
             .where(lambda x: ~x.isin(["", "NAN", "NONE"]))
        )

    # Índices normalizados (Series temporales, no se escriben)
    all_idx  = norm_series(df_all["SKU"])
    odoo_idx = norm_series(df_odoo["SKU"])

    # --- FUSION: copia de ODOO
    fusion_df = df_odoo.copy()

    # Map de atributos desde ALL (primera aparición no nula)
    if "Atributos" in df_all.columns:
        # construir Serie clave→atributos sin alterar df_all
        s_attrs = (
            df_all.loc[~df_all["Atributos"].isna(), ["SKU", "Atributos"]]
                 .assign(_K=norm_series(df_all.loc[~df_all["Atributos"].isna(), "SKU"]))
                 .drop_duplicates(subset="_K")
                 .set_index("_K")["Atributos"]
        )

        # asegurar columna Atributos en FUSION
        if "Atributos" not in fusion_df.columns:
            fusion_df["Atributos"] = pd.NA

        # clave normalizada temporal para FUSION (sin escribir columna)
        key_fusion = norm_series(fusion_df["SKU"])
        # alinear valores por índice normalizado
        attrs_from_all = s_attrs.reindex(key_fusion).to_numpy()
        # si en ALL hay valor, aplicarlo; si no, mantener lo existente
        fusion_df["Atributos"] = pd.Series(attrs_from_all, index=fusion_df.index).combine_first(fusion_df["Atributos"])

    # --- NOVEDADES: filas de ALL cuyo SKU no está en ODOO (comparación normalizada)
    novedades_mask = ~all_idx.isin(set(odoo_idx.dropna()))
    novedades_df = df_all.loc[novedades_mask].copy()

    # Ordenar columnas de FUSION como ODOO + extras al final
    odoo_cols = list(df_odoo.columns)
    other_cols = [c for c in fusion_df.columns if c not in odoo_cols]
    fusion_df = fusion_df[odoo_cols + other_cols] if other_cols else fusion_df[odoo_cols]

    # Escribir hojas
    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
            fusion_df.to_excel(w, sheet_name=hoja_fusion, index=False)
            novedades_df.to_excel(w, sheet_name=hoja_novedades, index=False)
    except FileNotFoundError:
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as w:
            fusion_df.to_excel(w, sheet_name=hoja_fusion, index=False)
            novedades_df.to_excel(w, sheet_name=hoja_novedades, index=False)

    return {"FUSION_rows": len(fusion_df), "NOVEDADES_rows": len(novedades_df)}
#fusionar_all_odoo(ruta)

def descargar_multimedia_excel(excel_path, destino_base="D:/BANCOS/OPTIMA/src"):
    import pandas as pd
    import os
    import ast

    # Leer Excel
    df = pd.read_excel(excel_path, sheet_name="Sheet1")

    # Normalizar nombres de columnas
    df.columns = df.columns.str.strip()

    required_cols = ["SKU", "Imagen principal", "Galería"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"El Excel debe tener columna '{col}' en Sheet1")

    # Crear carpeta base
    os.makedirs(destino_base, exist_ok=True)

    for idx, row in df.iterrows():
        sku = str(row["SKU"]).strip()
        if not sku or sku.lower() in ("nan", "none"):
            print(f"⚠️ Fila {idx+2}: sin SKU, se omite")
            continue

        img_principal = str(row["Imagen principal"]).strip()
        galeria_raw = str(row["Galería"]).strip()

        # Carpeta destino por SKU
        ruta_producto = os.path.join(destino_base, sku)
        os.makedirs(ruta_producto, exist_ok=True)

        # Descargar imagen principal (0.jpg)
        if img_principal and img_principal.startswith("http"):
            path_principal = os.path.join(ruta_producto, "0.jpg")
            if os.path.exists(path_principal):
                print(f"⏩ Saltada principal de {sku} (ya existe)")
            else:
                try:
                    raw_b64 = Utils.image_url_to_base64(img_principal)
                    raw = base64.b64decode(raw_b64)
                    with open(path_principal, "wb") as f:
                        f.write(raw)
                    print(f"✅ Principal descargada para {sku}")
                except Exception as e:
                    print(f"❌ Error descargando principal {img_principal}: {e}")

        # Descargar galería
        galeria = []
        try:
            if galeria_raw and galeria_raw not in ("nan", "None", "[]"):
                galeria = ast.literal_eval(galeria_raw)
        except Exception:
            galeria = [x.strip() for x in galeria_raw.split(",") if x.strip().startswith("http")]

        for i, url in enumerate(galeria, start=1):
            path_img = os.path.join(ruta_producto, f"{i}.jpg")
            if os.path.exists(path_img):
                print(f"⏩ Saltada galería {i} de {sku} (ya existe)")
                continue

            try:
                raw_b64 = Utils.image_url_to_base64(url)
                raw = base64.b64decode(raw_b64)
                with open(path_img, "wb") as f:
                    f.write(raw)
                print(f"  ✅ Galería {i} descargada para {sku}")
            except Exception as e:
                print(f"  ❌ Error descargando {url}: {e}")

    print("🎯 Descarga terminada.")

def copiar_carpetas_excel(excel_path, src_base="D:/BANCOS/OPTIMA/src", dst_base="D:/BANCOS/OPTIMA/src_nuevo"):
    import pandas as pd
    import os
    import shutil

    # Leer Excel
    df = pd.read_excel(excel_path, sheet_name="Sheet1")

    # Normalizar columnas
    df.columns = df.columns.str.strip()
    if "SKU" not in df.columns:
        raise ValueError("El Excel debe tener una columna 'SKU' en Sheet1")

    # Crear carpeta destino si no existe
    os.makedirs(dst_base, exist_ok=True)

    # Iterar SKUs
    skus = df["SKU"].dropna().astype(str).str.strip().unique()

    copiados = 0
    faltantes = []

    for sku in skus:
        src_folder = os.path.join(src_base, sku)
        dst_folder = os.path.join(dst_base, sku)

        if os.path.exists(src_folder):
            if os.path.exists(dst_folder):
                print(f"⏩ Saltada {sku}, ya existe en destino")
            else:
                try:
                    shutil.copytree(src_folder, dst_folder)
                    copiados += 1
                    print(f"✅ Copiada carpeta {sku}")
                except Exception as e:
                    print(f"❌ Error copiando {sku}: {e}")
        else:
            faltantes.append(sku)

    print(f"🎯 Proceso terminado. Carpetas copiadas: {copiados}. Faltantes: {len(faltantes)}")
    if faltantes:
        print("⚠️ No encontradas en src:", faltantes[:20], "..." if len(faltantes) > 20 else "")


def actualizar_precios_y_descuentos(path_excel):
    """
    Lee la hoja 'Products' de un Excel y actualiza:
      1. Si Precio == 0 → se reemplaza por Precio final.
      2. Calcula el descuento (%) entre Precio y Precio final.
    Guarda los cambios en el mismo archivo.
    """

    print(f"📂 Abriendo archivo: {path_excel}")

    # 1️⃣ Leer hoja 'Products'
    try:
        df = pd.read_excel(path_excel, sheet_name="Products")
    except Exception as e:
        print(f"❌ Error leyendo el Excel: {e}")
        return

    # 2️⃣ Normalizar nombres de columnas
    df.columns = df.columns.str.strip()

    # Verificar columnas necesarias
    required_cols = ["Precio final", "Descuento", "Precio"]
    if not all(col in df.columns for col in required_cols):
        print(f"⚠️ Faltan columnas. Se esperaban: {required_cols}")
        return

    # 3️⃣ Reemplazar Precio = 0 → Precio final
    print("🔧 Reemplazando precios iguales a 0...")
    df.loc[df["Precio"] == 0, "Precio"] = df["Precio final"]

    # 4️⃣ Calcular descuento
    print("🧮 Calculando descuentos...")
    def calcular_descuento(row):
        precio = row["Precio"]
        precio_final = row["Precio final"]

        # Evitar divisiones por cero o valores inválidos
        if precio and precio != 0:
            descuento = (1 - (precio_final / precio)) * 100
            return round(descuento, 2)
        return 0.0

    df["Descuento"] = df.apply(calcular_descuento, axis=1)

    # 5️⃣ Guardar cambios
    try:
        with pd.ExcelWriter(path_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name="Products", index=False)
        print("✅ Archivo actualizado correctamente.")
    except Exception as e:
        print(f"❌ Error guardando el Excel: {e}")

#actualizar_precios_y_descuentos(ruta)