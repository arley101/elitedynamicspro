import logging
import os
import requests
import json # Para manejo de errores
from typing import Dict, List, Optional, Union

# Importar Credential de Azure Identity
from azure.identity import ClientSecretCredential, CredentialUnavailableError

# Usar el logger de la función principal
logger = logging.getLogger("azure.functions")

# --- Constantes y Variables de Entorno Específicas ---
try:
    AZURE_SUBSCRIPTION_ID = os.environ['AZURE_SUBSCRIPTION_ID']
    AZURE_RESOURCE_GROUP = os.environ['AZURE_RESOURCE_GROUP']
    AZURE_LOCATION = os.environ.get('AZURE_LOCATION')
    CLIENT_ID = os.environ['CLIENT_ID']
    TENANT_ID = os.environ['TENANT_ID']
    CLIENT_SECRET = os.environ['CLIENT_SECRET']
except KeyError as e:
    logger.critical(f"Error Crítico: Falta variable de entorno esencial para Power Automate: {e}")
    raise ValueError(f"Configuración incompleta para Power Automate: falta {e}")

AZURE_MGMT_BASE_URL = "https://management.azure.com"
AZURE_MGMT_SCOPE = "https://management.azure.com/.default"
LOGIC_API_VERSION = "2019-05-01"
AZURE_MGMT_TIMEOUT = 60

# --- Helper de Autenticación (Específico para este módulo) ---
_credential_pa = None # Renombrado para evitar colisión potencial
_cached_mgmt_token_pa = None

def _get_azure_mgmt_token() -> str:
    """Obtiene un token para Azure Management API usando Client Credentials."""
    global _credential_pa, _cached_mgmt_token_pa
    if _cached_mgmt_token_pa: return _cached_mgmt_token_pa
    if not _credential_pa:
        logger.info("Creando credencial ClientSecretCredential para Azure Management (PA).")
        _credential_pa = ClientSecretCredential(tenant_id=TENANT_ID, client_id=CLIENT_ID, client_secret=CLIENT_SECRET)
    try:
        logger.info(f"Solicitando token para Azure Management con scope: {AZURE_MGMT_SCOPE}")
        token_info = _credential_pa.get_token(AZURE_MGMT_SCOPE)
        _cached_mgmt_token_pa = token_info.token
        logger.info("Token para Azure Management (PA) obtenido exitosamente.")
        return _cached_mgmt_token_pa
    except Exception as e:
        logger.error(f"Error obteniendo token de Azure Management (PA): {e}", exc_info=True)
        raise Exception(f"Error obteniendo token Azure (PA): {e}")

def _get_auth_headers_for_mgmt() -> Dict[str, str]:
    token = _get_azure_mgmt_token()
    return {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

# ---- POWER AUTOMATE (Flows) ----
# Funciones con parámetros reordenados (obligatorios primero)

def listar_flows(suscripcion_id: Optional[str] = None, grupo_recurso: Optional[str] = None, headers: Optional[Dict[str, str]] = None) -> dict:
    """Lista los flujos en un grupo de recursos."""
    auth_headers = _get_auth_headers_for_mgmt() # Usa auth interna
    sid = suscripcion_id or AZURE_SUBSCRIPTION_ID
    rg = grupo_recurso or AZURE_RESOURCE_GROUP
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows?api-version={LOGIC_API_VERSION}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Listando flows en '{rg}')")
        response = requests.get(url, headers=auth_headers, timeout=AZURE_MGMT_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados flows en el grupo de recursos '{rg}'.")
        return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en listar_flows: {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en listar_flows: {e}", exc_info=True); raise

def obtener_flow(nombre_flow: str, suscripcion_id: Optional[str] = None, grupo_recurso: Optional[str] = None, headers: Optional[Dict[str, str]] = None) -> dict:
    """Obtiene un flujo específico."""
    auth_headers = _get_auth_headers_for_mgmt()
    sid = suscripcion_id or AZURE_SUBSCRIPTION_ID
    rg = grupo_recurso or AZURE_RESOURCE_GROUP
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo flow '{nombre_flow}')")
        response = requests.get(url, headers=auth_headers, timeout=AZURE_MGMT_TIMEOUT)
        response.raise_for_status()
        data = response.json(); logger.info(f"Obtenido flow '{nombre_flow}'."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en obtener_flow: {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en obtener_flow: {e}", exc_info=True); raise

def crear_flow(nombre_flow: str, definicion_flow: dict, ubicacion: Optional[str] = None, suscripcion_id: Optional[str] = None, grupo_recurso: Optional[str] = None, headers: Optional[Dict[str, str]] = None) -> dict:
    """Crea un nuevo flujo."""
    auth_headers = _get_auth_headers_for_mgmt()
    sid = suscripcion_id or AZURE_SUBSCRIPTION_ID
    rg = grupo_recurso or AZURE_RESOURCE_GROUP
    loc = ubicacion or AZURE_LOCATION
    if not loc: raise ValueError("Se requiere 'ubicacion' o 'AZURE_LOCATION' para crear flow.")
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    body = {"location": loc, "properties": {"definition": definicion_flow}}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PUT {url} (Creando flow '{nombre_flow}')")
        response = requests.put(url, headers=auth_headers, json=body, timeout=AZURE_MGMT_TIMEOUT * 2)
        response.raise_for_status(); data = response.json(); logger.info(f"Flujo '{nombre_flow}' creado en '{rg}'."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en crear_flow: {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en crear_flow: {e}", exc_info=True); raise

def actualizar_flow(nombre_flow: str, definicion_flow: dict, suscripcion_id: Optional[str] = None, grupo_recurso: Optional[str] = None, headers: Optional[Dict[str, str]] = None) -> dict:
    """Actualiza un flujo existente."""
    auth_headers = _get_auth_headers_for_mgmt() # Usar auth interna
    sid = suscripcion_id or AZURE_SUBSCRIPTION_ID
    rg = grupo_recurso or AZURE_RESOURCE_GROUP
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    try:
         # Obtener flow actual para preservar otras propiedades
         current_flow = obtener_flow(nombre_flow=nombre_flow, suscripcion_id=sid, grupo_recurso=rg) # Llamada interna
         body = current_flow
         body["properties"]["definition"] = definicion_flow
    except Exception as get_err: raise Exception(f"No se pudo obtener flow actual para actualizar: {get_err}")
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PUT {url} (Actualizando flow '{nombre_flow}')")
        response = requests.put(url, headers=auth_headers, json=body, timeout=AZURE_MGMT_TIMEOUT * 2)
        response.raise_for_status(); data = response.json(); logger.info(f"Flujo '{nombre_flow}' actualizado en '{rg}'."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en actualizar_flow: {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en actualizar_flow: {e}", exc_info=True); raise

def eliminar_flow(nombre_flow: str, suscripcion_id: Optional[str] = None, grupo_recurso: Optional[str] = None, headers: Optional[Dict[str, str]] = None) -> dict:
    """Elimina un flujo."""
    auth_headers = _get_auth_headers_for_mgmt()
    sid = suscripcion_id or AZURE_SUBSCRIPTION_ID
    rg = grupo_recurso or AZURE_RESOURCE_GROUP
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando flow '{nombre_flow}')")
        response = requests.delete(url, headers=auth_headers, timeout=AZURE_MGMT_TIMEOUT)
        response.raise_for_status(); logger.info(f"Flujo '{nombre_flow}' eliminado de '{rg}'."); return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en eliminar_flow: {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en eliminar_flow: {e}", exc_info=True); raise

def ejecutar_flow(flow_url: str, parametros: Optional[dict] = None, headers: Optional[Dict[str, str]] = None) -> dict:
    """Ejecuta un flujo a través de su URL de desencadenador HTTP."""
    # Ignora headers de Graph/Mgmt; depende de la auth del trigger del flow
    request_headers = {'Content-Type': 'application/json'}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {flow_url} (Ejecutando flow trigger)")
        response = requests.post(flow_url, headers=request_headers, json=parametros if parametros else {}, timeout=AZURE_MGMT_TIMEOUT)
        response.raise_for_status()
        logger.info(f"Flujo en URL '{flow_url}' ejecutado (Triggered). Status: {response.status_code}")
        try: resp_data = response.json()
        except json.JSONDecodeError: resp_data = response.text
        return {"status": "Ejecutado", "code": response.status_code, "response_body": resp_data}
    except requests.exceptions.RequestException as e:
        # CORREGIDA INDENTACIÓN DEL RAISE
        logger.error(f"Error Request en ejecutar_flow: {e}", exc_info=True)
        raise # <-- Esta línea debe estar indentada bajo el except
    except Exception as e:
        logger.error(f"Error inesperado en ejecutar_flow: {e}", exc_info=True)
        raise

def obtener_estado_ejecucion_flow(run_id: str, nombre_flow: str, suscripcion_id: Optional[str] = None, grupo_recurso: Optional[str] = None, headers: Optional[Dict[str, str]] = None) -> dict:
    """Obtiene el estado de una ejecución específica de un flujo."""
    auth_headers = _get_auth_headers_for_mgmt()
    sid = suscripcion_id or AZURE_SUBSCRIPTION_ID
    rg = grupo_recurso or AZURE_RESOURCE_GROUP
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}/runs/{run_id}?api-version={LOGIC_API_VERSION}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo estado ejecución '{run_id}' de flow '{nombre_flow}')")
        response = requests.get(url, headers=auth_headers, timeout=AZURE_MGMT_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Obtenido estado ejecución '{run_id}'. Status: {data.get('properties', {}).get('status')}"); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en obtener_estado_ejecucion_flow: {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en obtener_estado_ejecucion_flow: {e}", exc_info=True); raise
