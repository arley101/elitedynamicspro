# helpers/http_client.py (Revisado y Final)
"""
Módulo auxiliar para realizar llamadas HTTP estandarizadas a APIs,
con logging incorporado y manejo básico de errores.
"""

import logging
import requests
import json
from typing import Dict, Any, Optional

# Usar el logger configurado en la función principal (Azure Functions)
logger = logging.getLogger("azure.functions")

# Importar constantes desde el módulo compartido
try:
    from shared.constants import GRAPH_API_TIMEOUT # Usar el timeout global por defecto
except ImportError:
    GRAPH_API_TIMEOUT = 45 # Fallback
    logger.warning("No se pudo importar GRAPH_API_TIMEOUT desde shared, usando default.")


def hacer_llamada_api(
    metodo: str,
    url: str,
    headers: Dict[str, str],
    params: Optional[Dict[str, Any]] = None,
    json_data: Optional[Dict[str, Any]] = None, # Para JSON body
    data: Optional[bytes] = None, # Para raw body (bytes)
    timeout: int = GRAPH_API_TIMEOUT,
    expect_json: bool = True # Flag para indicar si esperamos JSON o no
) -> Any:
    """
    Realiza una llamada HTTP genérica usando la librería requests.

    Args:
        metodo (str): Método HTTP ('GET', 'POST', 'PUT', 'PATCH', 'DELETE').
        url (str): URL completa del endpoint.
        headers (Dict[str, str]): Cabeceras (incluyendo Authorization y Content-Type si aplica).
        params (Optional[Dict[str, Any]]): Parámetros de consulta URL.
        json_data (Optional[Dict[str, Any]]): Cuerpo de la solicitud como JSON.
        data (Optional[bytes]): Cuerpo de la solicitud como bytes crudos (prioridad sobre json_data).
        timeout (int): Timeout de la solicitud en segundos.
        expect_json (bool): Si es True (default), intenta devolver response.json(). Si es False, devuelve el objeto Response.

    Returns:
        Any: La respuesta JSON decodificada por defecto, o el objeto Response si expect_json=False,
             o None si la respuesta es 204 No Content.
             Lanza una excepción si la llamada HTTP falla.
    """
    response: Optional[requests.Response] = None
    try:
        # Preparar headers (asegurar Content-Type si se envía JSON)
        request_headers = headers.copy()
        if json_data is not None and data is None and 'Content-Type' not in request_headers:
            request_headers['Content-Type'] = 'application/json'
        elif data is not None and 'Content-Type' not in request_headers:
             # Para data binaria, el Content-Type usualmente se pone antes de llamar
             logger.debug("Llamando helper con 'data' pero sin 'Content-Type' explícito en headers.")

        logger.info(f"API Call Helper: {metodo.upper()} {url}")
        if params: logger.debug(f"Params: {params}")
        if json_data and data is None: logger.debug(f"JSON Body: {json_data}")
        if data: logger.debug(f"Data Body: <{len(data)} bytes>")

        response = requests.request(
            method=metodo.upper(),
            url=url,
            headers=request_headers,
            params=params,
            json=json_data if data is None else None, # Enviar JSON solo si no hay data binaria
            data=data, # Enviar data binaria si existe
            timeout=timeout
        )
        response.raise_for_status() # Lanza HTTPError para 4xx/5xx

        # Manejar respuestas sin contenido
        if response.status_code == 204:
            logger.info(f"Respuesta OK ({response.status_code} No Content) desde {url}")
            return None

        # Devolver objeto Response completo si no se espera JSON (ej: para descargas)
        if not expect_json:
            logger.info(f"Respuesta OK ({response.status_code}) desde {url}. Devolviendo objeto Response.")
            return response

        # Intentar decodificar JSON si se espera y hay contenido
        if response.content:
             try:
                 return response.json()
             except json.JSONDecodeError:
                 logger.warning(f"Respuesta OK ({response.status_code}) pero no es JSON válido desde {url}. Devolviendo texto.")
                 return response.text
        else:
             logger.info(f"Respuesta OK ({response.status_code}) sin cuerpo JSON desde {url}")
             # Si se esperaba JSON pero no vino, devolver None o un dict vacío? Devolvemos None.
             return None

    except requests.exceptions.Timeout as timeout_err:
        logger.error(f"Timeout en llamada HTTP: {metodo.upper()} {url} - {timeout_err}", exc_info=True)
        raise Exception(f"Timeout ({timeout}s) durante la solicitud a {url}")
    except requests.exceptions.HTTPError as http_err:
        error_details = ""
        if http_err.response is not None:
            try: err_json = http_err.response.json(); error_details = f" - API Details: {err_json.get('error', {}).get('message', http_err.response.text)}"
            except json.JSONDecodeError: error_details = f" - Raw Response: {http_err.response.text}"
        logger.error(f"Error HTTP {http_err.response.status_code if http_err.response is not None else ''} en llamada: {metodo.upper()} {url}{error_details}", exc_info=True)
        raise Exception(f"Error HTTP ({http_err.response.status_code if http_err.response is not None else ''}) en la solicitud a {url}{error_details}")
    except requests.exceptions.RequestException as req_err:
        logger.error(f"Error de red/conexión en llamada HTTP: {metodo.upper()} {url} - {req_err}", exc_info=True)
        raise Exception(f"Error de red/conexión durante la solicitud a {url}: {req_err}")
    except Exception as e:
        logger.error(f"Error inesperado en http_client: {e}", exc_info=True)
        raise
