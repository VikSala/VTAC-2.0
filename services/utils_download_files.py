import os
import requests
import pandas as pd
import ast
from yt_dlp import YoutubeDL

def _descargar_archivo(url, destino, urls_descargadas):
    if url in urls_descargadas:
        print(f"Ya descargado anteriormente (saltado): {url}")
        return
    try:
        os.makedirs(os.path.dirname(destino), exist_ok=True)
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        with open(destino, 'wb') as f:
            f.write(response.content)
        print(f"Descargado: {destino}")
        urls_descargadas.add(url)
    except Exception as e:
        print(f"Error al descargar {url}: {e}")


def _descargar_video_youtube(url, destino, urls_descargadas):
    def limpiar_nombre_archivo(nombre):
        # Elimina caracteres no válidos para nombres de archivo en Windows
        caracteres_invalidos = r'\/:*?"<>|'
        return ''.join(c for c in nombre if c not in caracteres_invalidos).strip()

    if url in urls_descargadas:
        print(f"Video ya descargado anteriormente (saltado): {url}")
        return
    try:
        os.makedirs(destino, exist_ok=True)

        # Obtener información del video sin descargarlo
        info_opts = {'quiet': True, 'skip_download': True, 'format_sort': ['ext:mp4:m4a', 'vcodec:h264', 'acodec:aac']}
        with YoutubeDL(info_opts) as ydl:
            info = ydl.extract_info(url, download=False)
            titulo = info.get('title', 'video')
            titulo_limpio = limpiar_nombre_archivo(titulo)

        ydl_opts = {
            'outtmpl': os.path.join(destino, f"{titulo_limpio}.%(ext)s"),
            'format_sort': ['ext:mp4:m4a', 'vcodec:h264', 'acodec:aac'],
            'quiet': True,
            'merge_output_format': 'mp4',
        }

        with YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])

        print(f"Video de YouTube descargado: {url}.mp4")
        urls_descargadas.add(url)

    except Exception as e:
        print(f"Error al descargar video de YouTube {url}: {e}")


def procesar_excel(ruta_excel, carpeta_base):
    df = pd.read_excel(ruta_excel)
    urls_descargadas = set()
    videos_por_url = {}

    for _, fila in df.iterrows():
        key = str(fila.get('default_code', '')).strip()
        if not key:
            continue

        image_url = str(fila.get('Image', '')).strip()
        image_urls = fila.get('Image_Urls', '[]')
        pdf_urls = fila.get('Pdf_Urls', '{}')
        video_url = str(fila.get('Video_Urls', '')).strip()

        # Parsear campos complejos
        try:
            image_urls = ast.literal_eval(image_urls) if isinstance(image_urls, str) else []
            pdf_urls = ast.literal_eval(pdf_urls) if isinstance(pdf_urls, str) else {}
        except Exception as e:
            print(f"Error al interpretar URLs en línea {key}: {e}")
            image_urls, pdf_urls = [], {}

        # Descargar imagen principal
        if image_url and image_url.startswith('http'):
            ext = os.path.splitext(image_url)[1]
            nombre = f"{key}_0{ext}"
            destino = os.path.join(carpeta_base, 'images', nombre)
            _descargar_archivo(image_url, destino, urls_descargadas)

        # Imágenes adicionales
        for i, url in enumerate(image_urls):
            if url and isinstance(url, str) and url.startswith('http'):
                ext = os.path.splitext(url)[1]
                nombre = f"{key}_{i+1}{ext}"
                destino = os.path.join(carpeta_base, 'images', nombre)
                _descargar_archivo(url, destino, urls_descargadas)

        # PDFs
        for nombre_pdf, url in pdf_urls.items():
            if url and isinstance(url, str) and url.startswith('http'):
                ext = os.path.splitext(url)[1]
                nombre_archivo = f"{key}_{nombre_pdf}{ext}"
                destino = os.path.join(carpeta_base, 'pdfs', nombre_archivo)
                _descargar_archivo(url, destino, urls_descargadas)

        # Videos: agrupar por URL
        if video_url and video_url.startswith('http'):
            videos_por_url.setdefault(video_url, []).append(key)

    # Descargar videos agrupados por URL
    for url, skus in videos_por_url.items():
        carpeta_nombre = "_".join(sorted(set(skus)))
        carpeta_video = os.path.join(carpeta_base, 'videos', carpeta_nombre)
        if "youtube.com" in url or "youtu.be" in url:
            _descargar_video_youtube(url, carpeta_video, urls_descargadas)
        else:
            # Usa primer SKU como base para nombre de archivo
            nombre = f"{skus[0]}_0{os.path.splitext(url)[1]}"
            destino = os.path.join(carpeta_video, nombre)
            _descargar_archivo(url, destino, urls_descargadas)

# --- USO ---
ruta_excel = "C:/Users/mario/Documents/SMI Files/Multimedia.xlsx"
carpeta_salida = "D:/BANCOS/V-TAC-ES"

#procesar_excel(ruta_excel, carpeta_salida)
