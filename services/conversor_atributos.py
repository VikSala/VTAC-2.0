import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import ast
import os

def extract_unique_keys(df, col_name='Specifications'):
    """
    Extrae claves únicas de la columna Specifications.
    """
    def safe_eval(x):
        try:
            return ast.literal_eval(x)
        except Exception:
            return {}

    keys = set()
    for val in df[col_name].dropna():
        d = safe_eval(val)
        if isinstance(d, dict):
            keys.update(d.keys())
    return sorted(keys)

def main():
    root = tk.Tk()
    root.withdraw()

    # Selección del archivo Excel
    file_path = filedialog.askopenfilename(
        title='Selecciona el archivo Excel',
        filetypes=[('Archivos Excel', '*.xlsx *.xls')]
    )
    if not file_path:
        messagebox.showwarning('Aviso', 'No se ha seleccionado ningún archivo.')
        return

    try:
        sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
    except Exception as e:
        messagebox.showerror('Error al leer Excel', str(e))
        return

    result_dict = {}
    max_keys = 0

    for name, df in sheets.items():
        if 'Specifications' in df.columns:
            keys = extract_unique_keys(df, 'Specifications')
            result_dict[name] = keys
            max_keys = max(max_keys, len(keys))

    if not result_dict:
        messagebox.showerror('Error', "No se encontró la columna 'Specifications' en ninguna hoja.")
        return

    # Crear DataFrame con columnas por hoja y valores como claves únicas
    result_df = pd.DataFrame({name: pd.Series(keys) for name, keys in result_dict.items()})

    # Guardar en un nuevo archivo Excel
    base, ext = os.path.splitext(os.path.basename(file_path))
    save_path = filedialog.asksaveasfilename(
        title='Guardar archivo con claves únicas',
        defaultextension='.xlsx',
        initialfile=f'{base}_claves_unicas{ext}',
        filetypes=[('Archivos Excel', '*.xlsx')]
    )
    if not save_path:
        messagebox.showwarning('Aviso', 'No se ha especificado ruta de guardado.')
        return

    try:
        result_df.to_excel(save_path, index=False, sheet_name='Claves Specifications')
        messagebox.showinfo('Éxito', f'Archivo guardado en: {save_path}')
    except Exception as e:
        messagebox.showerror('Error al guardar Excel', str(e))

if __name__ == '__main__':
    main()
