# shared/constants.py

"""
Archivo de constantes compartidas para la aplicación.

Este archivo centraliza las constantes globales que se usan en múltiples
partes del proyecto para mejorar la mantenibilidad del código.
"""

from typing import Literal

# URL base para las solicitudes a la API de Microsoft Graph
BASE_URL: str = "https://graph.microsoft.com/v1.0"

# Tiempo máximo de espera para las solicitudes (en segundos)
GRAPH_API_TIMEOUT: int = 45

# Ejemplo de nueva constante: un tipo específico
API_VERSION: Literal["v1.0", "beta"] = "v1.0"

# Nota: Asegúrate de mantener las constantes actualizadas con los valores esperados.
