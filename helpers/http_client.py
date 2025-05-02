"""
shared/helpers/http_client.py

Cliente HTTP estándar y centralizado para realizar llamadas a Microsoft Graph API
u otros servicios externos. Incluye manejo robusto de errores, logging detallado
y timeouts configurables.
"""

import logging
import requests
import json
from typing import Dict, Any, Optional, Union

# Asumiendo que constants.py está en el directorio 'shared' padre
# Ajusta la ruta si tu estructura es diferente (ej. from ..constants import ...)
try:
    # Intenta importar desde la estructura relativa esperada
    from ..constants import GRAPH_API_TIMEOUT, BASE_URL
except ImportError:
    # Fallback si la importación falla (ej. al ejecutar localmente fuera de la estructura)
    # Es mejor asegurar que la estructura y PYTHONPATH estén correctos.
    GRAPH_API_TIMEOUT = 45 # Valor por defecto razonable
    BASE_URL = "https://graph.microsoft.com/v1.0" # Asegurar un valor por defecto
    logging.warning(
        f"No se pudo importar GRAPH_API_TIMEOUT/BASE_URL desde constants. "
        f"Usando defaults: Timeout={GRAPH_API_TIMEOUT}s, BaseURL={BASE_URL}. "
        f"Verifica la estructura del proyecto y los imports relativos."
    )

# Usar el logger estándar de Azure Functions para integración automática
logger = logging.getLogger("azure.functions")

def hacer_llamada_api(
    metodo: str,
    url: str,
    headers: Dict[str, str],
    params: Optional[Dict[str, Any]] = None,
    json_data: Optional[Dict[str, Any]] = None,
    data: Optional[Union[bytes, str]] = None, # Permitir bytes o string para data
    timeout: int = GRAPH_API_TIMEOUT,
    expect_json: bool = True
) -> Any:
    """
    Realiza una llamada HTTP genérica usando la librería requests, con logging
    y manejo de errores mejorados.

    Args:
        metodo (str): Método HTTP (GET, POST, PUT, PATCH, DELETE).
        url (str): URL completa del endpoint. Debe ser la URL final (ej., incluyendo BASE_URL si aplica).
        headers (Dict[str, str]): Cabeceras HTTP, DEBE incluir el token 'Authorization: Bearer ...'.
        params (Optional[Dict[str, Any]], optional): Parámetros de query string. Defaults to None.
        json_data (Optional[Dict[str, Any]], optional): Payload para enviar como JSON. Ignorado si 'data' se proporciona. Defaults to None.
        data (Optional[Union[bytes, str]], optional): Payload para enviar como raw data (bytes o string). Defaults to None.
        timeout (int, optional): Timeout en segundos para la solicitud. Defaults to GRAPH_API_TIMEOUT.
        expect_json (bool, optional): Indica si se espera una respuesta JSON.
                                      Si es False, devuelve el objeto Response completo. Defaults to True.

    Returns:
        Any: El cuerpo de la respuesta JSON decodificado si expect_json es True y la respuesta no está vacía (2xx).
             None si la respuesta es 204 No Content.
             El objeto requests.Response completo si expect_json es False.

    Raises:
        requests.exceptions.Timeout: Si la solicitud excede el tiempo de espera.
        requests.exceptions.RequestException: Si ocurre otro error durante la solicitud HTTP (conexión, estado HTTP 4xx/5xx).
        json.JSONDecodeError: Si expect_json es True pero la respuesta no es JSON válido.
        ValueError: Si falta la cabecera 'Authorization'.
    """
    # --- Validación de Entrada ---
    # Es CRUCIAL que el token venga en los headers desde la función principal (__init__.py)
    if not headers.get("Authorization"):
        # Lanzar un error claro si falta el token, ya que Graph API siempre lo requiere.
        error_msg = f"Llamada a {metodo} {url} SIN cabecera 'Authorization'. El token es obligatorio."
        logger.error(error_msg)
        raise ValueError(error_msg)

    # Asegurar que el método sea en mayúsculas
    metodo = metodo.upper()

    # --- Logging de la Solicitud ---
    # Log detallado para depuración (nivel DEBUG)
    logger.debug(f"Iniciando llamada API: {metodo} {url}")
    # No loguear headers completos por seguridad (puede contener tokens), solo indicar su presencia.
    logger.debug(f"Headers presentes: {list(headers.keys())}")
    if params:
        logger.debug(f"Query Params: {params}")
    # Loguear payload con cuidado (puede contener info sensible)
    if json_data and data is None:
        # Loguear solo las claves o una versión truncada/sanitizada si es necesario
        logger.debug(f"JSON Payload (claves): {list(json_data.keys())}")
    elif data:
        data_type = type(data).__name__
        data_preview = str(data[:100]) + '...' if isinstance(data, (str, bytes)) and len(data) > 100 else str(data)
        logger.debug(f"Raw Data Payload (tipo: {data_type}, preview: {data_preview})")
    logger.debug(f"Timeout: {timeout}s, Expect JSON: {expect_json}")

    # --- Ejecución de la Solicitud ---
    try:
        response = requests.request(
            method=metodo,
            url=url,
            headers=headers,
            params=params,
            # Enviar 'json' solo si 'json_data' tiene valor y 'data' no.
            json=json_data if data is None and json_data is not None else None,
            data=data,
            timeout=timeout
        )

        # Loguear status code y razón para todas las respuestas
        logger.debug(f"Respuesta recibida: Status={response.status_code}, Reason='{response.reason}'")

        # Lanzar excepción para respuestas 4xx (errores del cliente) y 5xx (errores del servidor)
        # Esto detendrá la ejecución aquí si hay un error HTTP.
        response.raise_for_status()

        # --- Procesamiento de Respuesta Exitosa (2xx) ---

        # Manejar respuesta 204 No Content (común en DELETE o PUT/PATCH sin retorno)
        if response.status_code == 204:
            logger.info(f"Llamada {metodo} {url} exitosa (204 No Content).")
            return None # Retornar None explícitamente

        # Procesar la respuesta según 'expect_json'
        if expect_json:
            try:
                # Intentar decodificar JSON. Si response.text está vacío, .json() puede fallar.
                if not response.text:
                     logger.warning(f"Respuesta 2xx de {url} recibida sin cuerpo para decodificar JSON.")
                     return None # O un diccionario vacío {} si es más apropiado

                json_response = response.json()
                # Loguear solo una parte o claves del JSON por si es muy grande o sensible
                # logger.debug(f"Respuesta JSON decodificada: {str(json_response)[:200]}...")
                logger.info(f"Llamada {metodo} {url} exitosa (Status: {response.status_code}). Respuesta JSON obtenida.")
                return json_response
            except json.JSONDecodeError as json_err:
                logger.error(f"Error al decodificar JSON de {url} (Status: {response.status_code}). Respuesta: {response.text[:500]}...")
                # Re-lanzar el error específico para que sea manejado arriba
                raise json_err
        else:
            # Devolver el objeto Response completo si no se espera JSON
            logger.info(f"Llamada {metodo} {url} exitosa (Status: {response.status_code}). Devolviendo objeto Response completo.")
            return response

    # --- Manejo de Excepciones Específicas ---
    except requests.exceptions.Timeout:
        logger.error(f"Timeout excedido ({timeout}s) en la llamada API: {metodo} {url}")
        # Re-lanzar Timeout para que la función llamante pueda manejarlo si es necesario
        raise
    except requests.exceptions.RequestException as e:
        # Capturar otros errores de requests (conexión, HTTPError ya capturado por raise_for_status, etc.)
        error_message = f"Error en la llamada API {metodo} {url}: {e}"
        # Intentar obtener más detalles del cuerpo de la respuesta de error si existe
        error_response_text = ""
        if e.response is not None:
             error_response_text = e.response.text[:500] # Limitar longitud del texto
             error_message += f" | Respuesta Error: Status={e.response.status_code}, Reason='{e.response.reason}', Body='{error_response_text}...'"
        logger.error(error_message)
        # Re-lanzar la excepción original de requests para que sea manejada por el __init__.py principal
        raise

