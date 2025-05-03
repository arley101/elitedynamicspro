"""
Paquete Shared

Este paquete contiene módulos compartidos que incluyen constantes,
utilidades y funciones comunes utilizadas en todo el proyecto.
"""

# Importar módulos clave del paquete para facilitar el acceso
from .constants import BASE_URL, GRAPH_API_TIMEOUT

# Definir __all__ para controlar qué se exporta al usar `from Shared import *`
__all__ = ["BASE_URL", "GRAPH_API_TIMEOUT"]

# Inicializaciones opcionales
# Por ejemplo, configuración de un logger o carga de configuraciones globales
# import logging
# logger = logging.getLogger(__name__)
# logger.setLevel(logging.INFO)
