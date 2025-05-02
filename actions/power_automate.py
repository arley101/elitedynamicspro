# actions/power_automate.py (Refactorizado v3 - Corrección Final)

import logging
import os
import requests # Para ejecutar_flow y tipos de excepción
import json
from typing import Dict, List, Optional, Union, Any

# Importar Credential de Azure Identity para autenticación con Azure Management API
# CORRECCIÓN: Eliminar try...except aquí. Si no se puede importar, debe fallar.
from azure.identity import ClientSecretCredential, CredentialUnavailableError

# Importar helper HTTP y constantes
try:
    from ..shared.helpers.http_client import hacer_llamada_api
    from ..shared.constants import GRAPH_API_TIMEOUT # Timeout base
except ImportError as e:
    logging.critical(f"Error CRÍTICO importando helpers/constantes en Power Automate: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    GRAPH_API_TIMEOUT = 45 # Default si falla import
    raise ImportError("No se pudo importar 'hacer_llamada_api' desde helpers.") from e

# Usar el logger estándar de Azure Functions
logger = logging.getLogger("azure.functions")

# --- Constantes y Variables de Entorno Específicas para Azure Management ---
try:
    AZURE_SUBSCRIPTION_ID = os.environ['AZURE_SUBSCRIPTION_ID']
    AZURE_RESOURCE_GROUP = os.environ['AZURE_RESOURCE_GROUP']
    AZURE_LOCATION = os.environ.get('AZURE_LOCATION')
    AZURE_CLIENT_ID = os.environ['AZURE_CLIENT_ID_MGMT']
    AZURE_TENANT_ID = os.environ['AZURE_TENANT_ID']
    AZURE_CLIENT_SECRET = os.environ['AZURE_CLIENT_SECRET_MGMT']
except KeyError as e:
    logger.critical(f"Error Crítico: Falta variable de entorno esencial para Power Automate Management: {e}")
    raise ValueError(f"Configuración incompleta para Power Automate Management: falta {e}")

AZURE_MGMT_BASE_URL = "https://management.azure.com"
AZURE_MGMT_SCOPE = "https://management.azure.com/.default"
LOGIC_API_VERSION = "2019-05-01"
AZURE_MGMT_TIMEOUT = max(GRAPH_API_TIMEOUT, 60)

# --- Helper de Autenticación (Específico para este módulo) ---
_credential_pa: Optional[ClientSecretCredential] = None
_cached_mgmt_token_pa: Optional[str] = None

def _get_azure_mgmt_token() -> str:
    """Obtiene un token de acceso para Azure Management API."""
    global _credential_pa, _cached_mgmt_token_pa

    # Verificar si azure-identity se importó correctamente (ya no es necesario el mock check)
    # if ClientSecretCredential is None:
    #     raise ImportError("Módulo azure.identity no disponible.")

    if _cached_mgmt_token_pa: return _cached_mgmt_token_pa

    if not _credential_pa:
        logger.info("Creando credencial ClientSecretCredential para Azure Management (PA).")
        try:
            _credential_pa = ClientSecretCredential(tenant_id=AZURE_TENANT_ID, client_id=AZURE_CLIENT_ID, client_secret=AZURE_CLIENT_SECRET)
        except Exception as cred_err:
             logger.critical(f"Error al crear ClientSecretCredential (PA): {cred_err}", exc_info=True)
             raise Exception(f"Error configurando credencial Azure (PA): {cred_err}") from cred_err

    try:
        logger.info(f"Solicitando token para Azure Management con scope: {AZURE_MGMT_SCOPE}")
        if _credential_pa is None: raise Exception("Credencial PA no inicializada.")
        token_info = _credential_pa.get_token(AZURE_MGMT_SCOPE)
        _cached_mgmt_token_pa = token_info.token
        logger.info("Token para Azure Management (PA) obtenido.")
        return _cached_mgmt_token_pa
    except CredentialUnavailableError as cred_err:
         logger.critical(f"Credencial no disponible para obtener token ARM: {cred_err}", exc_info=True)
         raise Exception(f"Credencial Azure (PA) no disponible: {cred_err}") from cred_err
    except Exception as e:
        logger.error(f"Error inesperado obteniendo token ARM (PA): {e}", exc_info=True)
        raise Exception(f"Error obteniendo token Azure (PA): {e}") from e

def _get_auth_headers_for_mgmt() -> Dict[str, str]:
    """Construye las cabeceras de autenticación para ARM API."""
    try:
        token = _get_azure_mgmt_token()
        return {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    except Exception as e:
        raise Exception(f"No se pudieron obtener cabeceras auth para Management API: {e}") from e

# ========================================================
# ==== FUNCIONES DE ACCIÓN PARA POWER AUTOMATE (FLOWS) ====
# ========================================================
# (Funciones listar_flows, obtener_flow, crear_flow, actualizar_flow,
#  eliminar_flow, ejecutar_flow sin cambios funcionales respecto a v2)

def listar_flows(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    auth_headers = _get_auth_headers_for_mgmt(); sid = parametros.get('suscripcion_id', AZURE_SUBSCRIPTION_ID); rg = parametros.get('grupo_recurso', AZURE_RESOURCE_GROUP)
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows?api-version={LOGIC_API_VERSION}"
    logger.info(f"Listando flows en Sub '{sid}', RG '{rg}'"); return hacer_llamada_api("GET", url, auth_headers, timeout=AZURE_MGMT_TIMEOUT)

def obtener_flow(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    nombre_flow: Optional[str] = parametros.get("nombre_flow");
    if not nombre_flow: raise ValueError("'nombre_flow' requerido.")
    auth_headers = _get_auth_headers_for_mgmt(); sid = parametros.get('suscripcion_id', AZURE_SUBSCRIPTION_ID); rg = parametros.get('grupo_recurso', AZURE_RESOURCE_GROUP)
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    logger.info(f"Obteniendo flow '{nombre_flow}' en RG '{rg}'"); return hacer_llamada_api("GET", url, auth_headers, timeout=AZURE_MGMT_TIMEOUT)

def crear_flow(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    nombre_flow: Optional[str] = parametros.get("nombre_flow"); definicion_flow: Optional[Dict[str, Any]] = parametros.get("definicion_flow"); ubicacion: Optional[str] = parametros.get("ubicacion", AZURE_LOCATION)
    if not nombre_flow: raise ValueError("'nombre_flow' requerido.");
    if not definicion_flow or not isinstance(definicion_flow, dict): raise ValueError("'definicion_flow' (dict) requerido.")
    if not ubicacion: raise ValueError("Se requiere 'ubicacion' o AZURE_LOCATION.")
    auth_headers = _get_auth_headers_for_mgmt(); sid = parametros.get('suscripcion_id', AZURE_SUBSCRIPTION_ID); rg = parametros.get('grupo_recurso', AZURE_RESOURCE_GROUP)
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    body: Dict[str, Any] = {"location": ubicacion, "properties": {"definition": definicion_flow}}
    logger.info(f"Creando flow '{nombre_flow}' en RG '{rg}', Loc '{ubicacion}'"); return hacer_llamada_api("PUT", url, auth_headers, json_data=body, timeout=AZURE_MGMT_TIMEOUT * 2)

def actualizar_flow(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    nombre_flow: Optional[str] = parametros.get("nombre_flow"); definicion_flow: Optional[Dict[str, Any]] = parametros.get("definicion_flow")
    if not nombre_flow: raise ValueError("'nombre_flow' requerido.");
    if not definicion_flow or not isinstance(definicion_flow, dict): raise ValueError("'definicion_flow' (dict) requerido.")
    auth_headers = _get_auth_headers_for_mgmt(); sid = parametros.get('suscripcion_id', AZURE_SUBSCRIPTION_ID); rg = parametros.get('grupo_recurso', AZURE_RESOURCE_GROUP)
    try:
        params_get = {"nombre_flow": nombre_flow, "suscripcion_id": sid, "grupo_recurso": rg}; current_flow = obtener_flow(params_get, {})
        current_location = current_flow.get("location");
        if not current_location: raise ValueError("No se pudo obtener ubicación del flow existente.")
    except Exception as get_err: raise Exception(f"No se pudo obtener flow actual '{nombre_flow}' para actualizar: {get_err}") from get_err
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    body: Dict[str, Any] = {"location": current_location, "properties": {"definition": definicion_flow}}
    logger.info(f"Actualizando flow '{nombre_flow}' en RG '{rg}'"); return hacer_llamada_api("PUT", url, auth_headers, json_data=body, timeout=AZURE_MGMT_TIMEOUT * 2)

def eliminar_flow(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    nombre_flow: Optional[str] = parametros.get("nombre_flow");
    if not nombre_flow: raise ValueError("'nombre_flow' requerido.")
    auth_headers = _get_auth_headers_for_mgmt(); sid = parametros.get('suscripcion_id', AZURE_SUBSCRIPTION_ID); rg = parametros.get('grupo_recurso', AZURE_RESOURCE_GROUP)
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    logger.info(f"Eliminando flow '{nombre_flow}' de RG '{rg}'"); hacer_llamada_api("DELETE", url, auth_headers, timeout=AZURE_MGMT_TIMEOUT); return {"status": "Eliminado", "flow": nombre_flow}

def ejecutar_flow(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    flow_url: Optional[str] = parametros.get("flow_url"); payload: Optional[Dict[str, Any]] = parametros.get("payload")
    if not flow_url: raise ValueError("'flow_url' requerido.")
    request_headers = headers.copy();
    if payload: request_headers['Content-Type'] = 'application/json'
    logger.info(f"Ejecutando trigger de flow: POST {flow_url}")
    try:
        response = requests.post(flow_url, headers=request_headers, json=payload if payload else None, timeout=AZURE_MGMT_TIMEOUT)
        response.raise_for_status(); logger.info(f"Trigger flow '{flow_url}' ejecutado. Status: {response.status_code}")
        try: resp_data = response.json()
        except json.JSONDecodeError: resp_data = response.text
        return {"status": "Ejecutado" if response.ok else "Fallido", "status_code": response.status_code, "response_body": resp_data}
    except requests.exceptions.RequestException as e: error_body = e.response.text[:200] if e.response else "N/A"; logger.error(f"Error Request ejecutando trigger flow '{flow_url}': {e}. Respuesta: {error_body}", exc_info=True); raise Exception(f"Error API ejecutando trigger flow: {e}") from e
    except Exception as e: logger.error(f"Error inesperado ejecutando trigger flow '{flow_url}': {e}", exc_info=True); raise

def obtener_estado_ejecucion_flow(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Obtiene el estado de una ejecución (run) específica de un flujo."""
    nombre_flow: Optional[str] = parametros.get("nombre_flow")
    run_id: Optional[str] = parametros.get("run_id")

    # Corrección Flake8 E999: Separar los if/raise en líneas distintas
    if not nombre_flow:
        raise ValueError("'nombre_flow' es requerido.")
    if not run_id:
        raise ValueError("'run_id' es requerido.")

    auth_headers = _get_auth_headers_for_mgmt()
    sid = parametros.get('suscripcion_id', AZURE_SUBSCRIPTION_ID)
    rg = parametros.get('grupo_recurso', AZURE_RESOURCE_GROUP)

    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}/runs/{run_id}?api-version={LOGIC_API_VERSION}"
    logger.info(f"Obteniendo estado de ejecución '{run_id}' flow '{nombre_flow}'")
    return hacer_llamada_api("GET", url, auth_headers, timeout=AZURE_MGMT_TIMEOUT)

# --- FIN DEL MÓDULO actions/power_automate.py ---
