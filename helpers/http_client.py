# helpers/http_client.py
"""
Módulo auxiliar para realizar llamadas HTTP estandarizadas a APIs (ej: Microsoft Graph),
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
    json_data: Optional[Dict[str, Any]] = None, # Para POST/PUT/PATCH con JSON body
    data: Optional[bytes] = None, # Para PUT/POST con body binario (ej: subida de archivos)
    timeout: int = GRAPH_API_TIMEOUT
) -> Any:
    """
    Realiza una llamada HTTP genérica usando la librería requests.

    Args:
        metodo (str): Método HTTP ('GET', 'POST', 'PUT', 'PATCH', 'DELETE').
        url (str): URL completa del endpoint.
        headers (Dict[str, str]): Cabeceras de la solicitud (incluyendo Authorization).
        params (Optional[Dict[str, Any]]): Parámetros de consulta URL.
        json_data (Optional[Dict[str, Any]]): Cuerpo de la solicitud en formato JSON.
        data (Optional[bytes]): Cuerpo de la solicitud como bytes crudos.
        timeout (int): Timeout de la solicitud en segundos.

    Returns:
        Any: La respuesta JSON decodificada si la llamada fue exitosa y devolvió JSON.
             Puede devolver None o texto si la respuesta no es JSON (ej: 204 No Content, descarga de archivo).
             Lanza una excepción si la llamada HTTP falla.
    """
    response: Optional[requests.Response] = None
    try:
        # Asegurar Content-Type si se envía JSON y no está en headers
        request_headers = headers.copy()
        if json_data is not None and 'Content-Type' not in request_headers:
            request_headers['Content-Type'] = 'application/json'
        # Si se envían bytes (data), Content-Type debe establecerse explícitamente antes si es necesario (ej: octet-stream)

        logger.info(f"API Call: {metodo.upper()} {url}")
        if params: logger.debug(f"Params: {params}")
        if json_data: logger.debug(f"JSON Body: {json_data}")
        if data: logger.debug(f"Data Body: <{len(data)} bytes>")

        response = requests.request(
            method=metodo.upper(),
            url=url,
            headers=request_headers,
            params=params,
            json=json_data,
            data=data,
            timeout=timeout
        )
        response.raise_for_status() # Lanza HTTPError para 4xx/5xx

        # Manejar respuestas sin contenido (ej: 204 No Content en DELETE/PATCH)
        if response.status_code == 204:
            logger.info(f"Respuesta exitosa (Sin Contenido {response.status_code}) desde {url}")
            return None # O un dict de status si se prefiere: {"status": "ok", "code": 204}

        # Intentar decodificar JSON si hay contenido
        if response.content:
             try:
                 return response.json()
             except json.JSONDecodeError:
                 logger.warning(f"Respuesta exitosa ({response.status_code}) pero no es JSON válido desde {url}. Devolviendo texto.")
                 return response.text # Devolver como texto si no es JSON
        else:
             logger.info(f"Respuesta exitosa ({response.status_code}) sin cuerpo desde {url}")
             return None # O dict de status

    except requests.exceptions.Timeout as timeout_err:
        logger.error(f"Timeout en llamada HTTP: {metodo.upper()} {url} - {timeout_err}", exc_info=True)
        raise Exception(f"Timeout ({timeout}s) durante la solicitud a {url}")
    except requests.exceptions.HTTPError as http_err:
        # Intentar obtener detalles del error de la respuesta si es posible
        error_details = ""
        if http_err.response is not None:
            try:
                err_json = http_err.response.json()
                error_details = f" - Detalles API: {err_json.get('error', {}).get('message', http_err.response.text)}"
            except json.JSONDecodeError:
                error_details = f" - Respuesta no JSON: {http_err.response.text}"
        logger.error(f"Error HTTP {http_err.response.status_code if http_err.response is not None else ''} en llamada: {metodo.upper()} {url}{error_details}", exc_info=True)
        raise Exception(f"Error HTTP ({http_err.response.status_code if http_err.response is not None else ''}) en la solicitud a {url}{error_details}")
    except requests.exceptions.RequestException as req_err:
        logger.error(f"Error de red/conexión en llamada HTTP: {metodo.upper()} {url} - {req_err}", exc_info=True)
        raise Exception(f"Error de red/conexión durante la solicitud a {url}: {req_err}")
    except Exception as e:
        logger.error(f"Error inesperado en http_client: {e}", exc_info=True)
        raise # Re-lanzar error inesperado
