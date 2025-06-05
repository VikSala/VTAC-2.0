# Scrap/__init__.py
"""
Paquete de scrapers para diferentes regiones
"""

# Importaciones opcionales para evitar errores si falta algún módulo
try:
    from . import vtac_it
except ImportError:
    vtac_it = None

try:
    from . import vtac_uk
except ImportError:
    vtac_uk = None

try:
    from . import vtac_buyleds
except ImportError:
    vtac_buyleds = None

__all__ = ['vtac_it', 'vtac_uk', 'vtac_buyleds']
