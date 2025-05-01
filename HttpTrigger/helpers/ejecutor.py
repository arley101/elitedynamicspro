"""
helpers/ejecutor.py

Módulo para ejecución de acciones mapeadas.
"""

from typing import Callable, Dict, Any

def ejecutar_accion(funcion: Callable[..., Any], parametros: Dict[str, Any], headers: Dict[str, str]) -> Any:
    """
    Ejecuta la acción solicitada con los parámetros y cabeceras proporcionados.
    """
    try:
        # Validar si la función requiere "headers"
        if 'headers' in funcion.__annotations__:
            return funcion(headers=headers, **parametros)
        else:
            return funcion(**parametros)
    except Exception as e:
        raise RuntimeError(f"Error al ejecutar la función '{funcion.__name__}': {e}")
