"""
helpers/ejecutor.py

Módulo para ejecución dinámica de acciones mapeadas.

Correcciones:
- Se utiliza `inspect.signature` para verificar de forma robusta si la
  función destino acepta el parámetro 'headers'.
"""

import logging
import inspect # Importar el módulo inspect
from typing import Callable, Dict, Any

logger = logging.getLogger("azure.functions") # Usar el logger estándar

def ejecutar_accion(funcion: Callable[..., Any], parametros: Dict[str, Any], headers: Dict[str, Any]) -> Any:
    """
    Ejecuta la acción solicitada con los parámetros y cabeceras proporcionados,
    verificando la firma de la función destino.

    Args:
        funcion: La función a ejecutar (obtenida del mapeo).
        parametros: Diccionario con los parámetros validados para la función.
        headers: Diccionario con las cabeceras HTTP de la solicitud original.

    Returns:
        El resultado de la ejecución de la función.

    Raises:
        RuntimeError: Si ocurre un error durante la ejecución de la función.
        TypeError: Si los parámetros proporcionados no coinciden con la firma de la función.
    """
    try:
        # Obtener la firma de la función para analizar sus parámetros
        sig = inspect.signature(funcion)
        func_params = sig.parameters

        # Crear el diccionario de argumentos a pasar
        kwargs_to_pass = {}

        # Pasar headers solo si la función lo define explícitamente en su firma
        if 'headers' in func_params:
            # Verificar si 'headers' espera un tipo específico (opcional pero buena práctica)
            # header_param_type = func_params['headers'].annotation
            # if header_param_type == inspect.Parameter.empty or isinstance(headers, header_param_type):
            kwargs_to_pass['headers'] = headers
            # else:
            #    logger.warning(f"El parámetro 'headers' en {funcion.__name__} espera {header_param_type}, pero se recibió {type(headers)}. Se intentará pasar de todas formas.")
            #    kwargs_to_pass['headers'] = headers # O lanzar error si se prefiere estricto

        # Pasar los parámetros validados
        # Se puede hacer una validación adicional aquí para asegurar que todos los
        # parámetros requeridos por la función (sin default) están presentes en 'parametros'.
        for param_name in func_params:
             if param_name != 'headers': # Ya manejamos headers
                 if param_name in parametros:
                     kwargs_to_pass[param_name] = parametros[param_name]
                 # else:
                     # Si el parámetro no tiene valor por defecto y no está en 'parametros',
                     # inspect.signature().bind(**kwargs_to_pass) fallará más adelante.
                     # O puedes añadir una verificación explícita aquí:
                     # if func_params[param_name].default == inspect.Parameter.empty:
                     #     raise TypeError(f"Falta el parámetro requerido '{param_name}' para la función {funcion.__name__}")

        # Llamar a la función con los argumentos desempaquetados
        # Usar sig.bind para validar que los kwargs coinciden con la firma antes de llamar
        try:
            bound_args = sig.bind(**kwargs_to_pass)
            bound_args.apply_defaults() # Aplicar valores por defecto si los hay
            logger.info(f"Llamando a {funcion.__name__} con argumentos: {bound_args.arguments}")
            return funcion(*bound_args.args, **bound_args.kwargs)
        except TypeError as te:
            logger.error(f"Error de tipo al intentar llamar a {funcion.__name__} con parámetros {kwargs_to_pass}: {te}")
            # Re-lanzar el error para que sea manejado por el __init__.py
            raise TypeError(f"Parámetros incorrectos para la acción '{funcion.__name__}': {te}")


    except Exception as e:
        # Captura cualquier otra excepción durante la ejecución de la función
        logger.exception(f"Error inesperado al ejecutar la función '{funcion.__name__}': {e}")
        # Re-lanzar como RuntimeError o un tipo de excepción personalizado
        raise RuntimeError(f"Error interno al ejecutar la acción '{funcion.__name__}': {e}")

