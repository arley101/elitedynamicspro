import logging
import requests
import json
from typing import Dict, Any, Optional

logger = logging.getLogger("azure.functions")

try:
    from shared.constants import GRAPH_API_TIMEOUT
except ImportError:
    GRAPH_API_TIMEOUT = 45
    logger.warning("Usando default para timeout.")

def hacer_llamada_api(
    metodo: str,
    url: str,
    headers: Dict[str, str],
    params: Optional[Dict[str, Any]] = None,
    json_data: Optional[Dict[str, Any]] = None,
    data: Optional[bytes] = None,
    timeout: int = GRAPH_API_TIMEOUT,
    expect_json: bool = True
) -> Any:
    """
    Realiza una llamada HTTP genérica usando requests.
    """
    try:
        response = requests.request(
            method=metodo.upper(),
            url=url,
            headers=headers,
            params=params,
            json=json_data if data is None else None,
            data=data,
            timeout=timeout
        )
        response.raise_for_status()

        if response.status_code == 204:
            return None

        if expect_json:
            return response.json()
        return response
    except requests.exceptions.RequestException as e:
        logger.error(f"Error en {metodo} {url}: {e}")
        raise
