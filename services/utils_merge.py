# services/utils_merge.py
import pandas as pd


def detectar_skus_unicos_(archivo_excel):
    # Cargar todas las hojas
    hojas = pd.read_excel(archivo_excel, sheet_name=['ES', 'UK', 'ITA', 'OPT'])
    df_es = hojas['ES']
    df_uk = hojas['UK']
    df_ita = hojas['ITA']
    df_opt = hojas['OPT']

    # Extraer SKUs como sets
    skus_es = set(df_es['SKU'].astype(str))
    skus_uk = set(df_uk['SKU'].astype(str))
    skus_ita = set(df_ita['SKU'].astype(str))
    skus_opt = set(df_opt['SKU'].astype(str))

    # Todos los SKUs de comparación
    todos_comparar = {
        'ES': (df_es, skus_es - skus_uk - skus_ita - skus_opt),
        'UK': (df_uk, skus_uk - skus_es - skus_ita - skus_opt),
        'ITA': (df_ita, skus_ita - skus_es - skus_uk - skus_opt)
    }

    # Filtrar filas únicas por hoja
    df_nuevos = pd.DataFrame()
    for nombre_hoja, (df, skus_unicos) in todos_comparar.items():
        filas = df[df['SKU'].astype(str).isin(skus_unicos)]
        df_nuevos = pd.concat([df_nuevos, filas], ignore_index=True)
    return df_nuevos