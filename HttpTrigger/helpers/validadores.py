"""
helpers/validadores.py

Módulo para validaciones y conversiones de parámetros.
"""

import json
from datetime import datetime
from typing import Any, Dict

def validar_parametros(parametros: Dict[str, Any], type_hints: Dict[str, Any]) -> Dict[str, Any]:
    """
    Valida y convierte los parámetros según las anotaciones de tipo.
    """
    params_procesados = parametros.copy()
    for param_name, param_type in type_hints.items():
        if param_name not in params_procesados or params_procesados[param_name] is None:
            continue

        original_value = params_procesados[param_name]
        try:
            if param_type is int:
                params_procesados[param_name] = int(original_value)
            elif param_type is bool:
                params_procesados[param_name] = str(original_value).lower() in ['true', '1', 'yes']
            elif param_type is float:
                params_procesados[param_name] = float(original_value)
            elif param_type is datetime:
                params_procesados[param_name] = datetime.fromisoformat(original_value.replace('Z', '+00:00'))
            elif param_type is list and isinstance(original_value, str):
                params_procesados[param_name] = json.loads(original_value)
            elif param_type is dict and isinstance(original_value, str):
                params_procesados[param_name] = json.loads(original_value)
            elif param_type is str:
                params_procesados[param_name] = str(original_value)
            else:
                raise ValueError(f"Tipo no soportado para '{param_name}': {param_type}")
        except (ValueError, TypeError, json.JSONDecodeError):
            raise ValueError(f"Error al convertir '{param_name}' (valor: {original_value}) a {param_type}")

    return params_procesados
