"""
actions/sharepoint.py

Módulo mejorado para la interacción avanzada con SharePoint utilizando Microsoft Graph API.
Incluye funcionalidades para listas, documentos, manejo de sesiones y memoria persistente.
"""

import logging
import requests
import os
from typing import Dict, List, Optional, Any

# Configuración del logger
logger = logging.getLogger("azure.functions")

# Constantes globales
BASE_URL = os.getenv("BASE_URL", "https://graph.microsoft.com/v1.0")
GRAPH_API_TIMEOUT = int(os.getenv("GRAPH_API_TIMEOUT", 45))
SHAREPOINT_DEFAULT_SITE_ID = os.getenv("SHAREPOINT_DEFAULT_SITE_ID")
SHAREPOINT_DEFAULT_DRIVE_ID = os.getenv("SHAREPOINT_DEFAULT_DRIVE_ID", "Documents")

# --- Helper para manejar solicitudes HTTP ---
def hacer_llamada_http(
    metodo: str,
    url: str,
    headers: Dict[str, str],
    params: Optional[Dict[str, Any]] = None,
    data: Optional[Dict[str, Any]] = None,
    timeout: int = GRAPH_API_TIMEOUT
) -> Any:
    """
    Realiza una llamada HTTP con manejo de errores y logging.
    """
    try:
        logger.info(f"API Call: {metodo.upper()} {url} Params: {params}")
        response = requests.request(
            method=metodo,
            url=url,
            headers=headers,
            params=params,
            json=data,
            timeout=timeout
        )
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        logger.error(f"Error en la llamada HTTP: {e}", exc_info=True)
        raise Exception(f"Error en la solicitud HTTP: {e}")

# --- Helper para obtener Site ID ---
def obtener_site_id(headers: Dict[str, str], site_id: Optional[str] = None) -> str:
    """
    Obtiene el ID de un sitio de SharePoint. Usa un ID predeterminado si no se proporciona.
    """
    if site_id:
        return site_id
    if SHAREPOINT_DEFAULT_SITE_ID:
        return SHAREPOINT_DEFAULT_SITE_ID

    # Obtener el sitio raíz como fallback
    url = f"{BASE_URL}/sites/root?$select=id"
    data = hacer_llamada_http("GET", url, headers)
    site_id = data.get("id")
    if not site_id:
        raise ValueError("No se pudo obtener el ID del sitio raíz.")
    return site_id

# --- Funciones de Memoria Persistente ---
def crear_lista_memoria(headers: Dict[str, str], site_id: Optional[str] = None) -> dict:
    """
    Crea una lista en SharePoint para almacenar datos persistentes.
    """
    target_site_id = obtener_site_id(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists"
    body = {
        "displayName": "MemoriaPersistente",
        "columns": [
            {"name": "SessionID", "text": {}},
            {"name": "Clave", "text": {}},
            {"name": "Valor", "text": {}}
        ],
        "list": {"template": "genericList"}
    }
    return hacer_llamada_http("POST", url, headers, data=body)

def guardar_dato_memoria(headers: Dict[str, str], session_id: str, clave: str, valor: str, site_id: Optional[str] = None) -> dict:
    """
    Guarda un dato en la lista de memoria persistente asociada a una sesión.
    """
    target_site_id = obtener_site_id(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/MemoriaPersistente/items"
    body = {"fields": {"SessionID": session_id, "Clave": clave, "Valor": valor}}
    return hacer_llamada_http("POST", url, headers, data=body)

def recuperar_datos_sesion(headers: Dict[str, str], session_id: str, site_id: Optional[str] = None) -> List[dict]:
    """
    Recupera todos los datos asociados a una sesión específica.
    """
    target_site_id = obtener_site_id(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/MemoriaPersistente/items"
    params = {"$filter": f"fields/SessionID eq '{session_id}'", "$expand": "fields"}
    data = hacer_llamada_http("GET", url, headers, params=params)
    return [item["fields"] for item in data.get("value", [])]

def eliminar_memoria_sesion(headers: Dict[str, str], session_id: str, site_id: Optional[str] = None) -> dict:
    """
    Elimina todos los datos asociados a una sesión específica.
    """
    target_site_id = obtener_site_id(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/MemoriaPersistente/items"
    params = {"$filter": f"fields/SessionID eq '{session_id}'"}
    data = hacer_llamada_http("GET", url, headers, params=params)

    # Elimina cada elemento encontrado
    for item in data.get("value", []):
        item_id = item.get("id")
        delete_url = f"{url}/{item_id}"
        hacer_llamada_http("DELETE", delete_url, headers)
    return {"status": "Eliminado", "session_id": session_id}

# --- Funciones Avanzadas ---
def exportar_datos_lista(headers: Dict[str, str], lista_id: str, formato: str = "json", site_id: Optional[str] = None) -> str:
    """
    Exporta los datos de una lista en formato JSON o CSV.
    """
    target_site_id = obtener_site_id(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id}/items"
    params = {"$expand": "fields"}
    data = hacer_llamada_http("GET", url, headers, params=params)

    if formato.lower() == "json":
        return data
    elif formato.lower() == "csv":
        import csv
        from io import StringIO

        output = StringIO()
        writer = csv.writer(output)
        writer.writerow(data["value"][0]["fields"].keys())  # Escribir encabezados
        for item in data["value"]:
            writer.writerow(item["fields"].values())
        return output.getvalue()
    else:
        raise ValueError("Formato no soportado. Use 'json' o 'csv'.")
