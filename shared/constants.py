# shared/constants.py

"""
Archivo de constantes compartidas para la aplicación.

Este archivo centraliza las constantes globales que se usan en múltiples
partes del proyecto para mejorar la mantenibilidad del código.

NOTA: No incluir información sensible como claves de API o secretos.
"""

from typing import Literal

# --- Configuración de la API ---
# URL base para las solicitudes a la API de Microsoft Graph
BASE_URL: str = "https://graph.microsoft.com/v1.0"

# Versión de la API a utilizar
API_VERSION: Literal["v1.0", "beta"] = "v1.0"

# --- Configuración de Tiempo ---
# Tiempo máximo de espera para las solicitudes (en segundos)
GRAPH_API_TIMEOUT: int = 45

# --- Configuración Adicional ---
# Ejemplo de constante adicional: Nombre del entorno
ENVIRONMENT: Literal["development", "staging", "production"] = "development"

# Nota: Mantén este archivo actualizado con los valores esperados.
