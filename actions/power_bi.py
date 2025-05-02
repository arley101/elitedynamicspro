# actions/power_bi.py (Refactorizado v3 - Corrección MyPy)

import logging
import os
import requests # Solo para tipos de excepción
import json
from typing import Dict, List, Optional, Union, Any

# Importar Credential de Azure Identity para autenticación con Power BI API
try:
    from azure.identity import ClientSecretCredential, CredentialUnavailableError
except ImportError:
    # Log crítico y error si falta dependencia esencial
    logging.critical("Error CRÍTICO: Falta 'azure-identity'. Instala con 'pip install azure-identity'. Power BI actions no funcionarán.")
    # CORRECCIÓN MyPy: NO definir mocks aquí para evitar errores [no-redef]
    ClientSecretCredential = None # type: ignore
    CredentialUnavailableError = None # type: ignore

# Importar helper HTTP y constantes
try:
    from ..shared.helpers.http_client import hacer_llamada_api
    from ..shared.constants import GRAPH_API_TIMEOUT # Timeout base
except ImportError as e:
    logging.critical(f"Error CRÍTICO importando helpers/constantes en Power BI: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    GRAPH_API_TIMEOUT = 45 # Default si falla import
    raise ImportError("No se pudo importar 'hacer_llamada_api' desde helpers.") from e

# Usar el logger estándar de Azure Functions
logger = logging.getLogger("azure.functions")

# --- Constantes y Variables de Entorno Específicas para Power BI API ---
try:
    PBI_CLIENT_ID = os.environ['AZURE_CLIENT_ID_PBI']
    PBI_TENANT_ID = os.environ['AZURE_TENANT_ID']
    PBI_CLIENT_SECRET = os.environ['AZURE_CLIENT_SECRET_PBI']
except KeyError as e:
    logger.critical(f"Error Crítico: Falta variable de entorno esencial para autenticación Power BI: {e}")
    raise ValueError(f"Configuración incompleta para Power BI API: falta {e}")

PBI_BASE_URL = "https://api.powerbi.com/v1.0/myorg"
PBI_SCOPE = "https://analysis.windows.net/powerbi/api/.default"
PBI_TIMEOUT = max(GRAPH_API_TIMEOUT, 60)

# --- Helper de Autenticación (Específico para este módulo) ---
_credential_pbi: Optional[ClientSecretCredential] = None
_cached_pbi_token: Optional[str] = None

def _get_pbi_token() -> str:
    """Obtiene un token de acceso para Power BI API usando Client Credentials."""
    global _credential_pbi, _cached_pbi_token

    # Verificar si azure-identity se importó correctamente
    if ClientSecretCredential is None:
        raise ImportError("Módulo azure.identity no disponible. No se puede autenticar con Power BI.")

    if _cached_pbi_token: return _cached_pbi_token

    if not _credential_pbi:
        logger.info("Creando credencial ClientSecretCredential para Power BI API.")
        try:
            _credential_pbi = ClientSecretCredential(tenant_id=PBI_TENANT_ID, client_id=PBI_CLIENT_ID, client_secret=PBI_CLIENT_SECRET)
        except Exception as cred_err:
             logger.critical(f"Error al crear ClientSecretCredential para PBI: {cred_err}", exc_info=True)
             raise Exception(f"Error configurando credencial Power BI: {cred_err}") from cred_err

    try:
        logger.info(f"Solicitando token para Power BI con scope: {PBI_SCOPE}")
        if _credential_pbi is None: raise Exception("Credencial PBI no inicializada.")
        token_info = _credential_pbi.get_token(PBI_SCOPE)
        _cached_pbi_token = token_info.token
        logger.info("Token para Power BI API obtenido exitosamente.")
        return _cached_pbi_token
    except CredentialUnavailableError as cred_err:
         # Este error es específico de azure-identity, verificar si se importó
         if CredentialUnavailableError is None: # Si la importación falló
              logger.critical(f"Error obteniendo token PBI: CredentialUnavailableError no disponible (azure.identity falló).")
              raise Exception("Error obteniendo token PBI: azure.identity no disponible.") from cred_err
         else: # La importación funcionó pero la credencial falló
              logger.critical(f"Credencial no disponible para obtener token PBI: {cred_err}", exc_info=True)
              raise Exception(f"Credencial Power BI no disponible: {cred_err}") from cred_err
    except Exception as e:
        logger.error(f"Error inesperado obteniendo token de Power BI: {e}", exc_info=True)
        raise Exception(f"Error obteniendo token Power BI: {e}") from e

def _get_auth_headers_for_pbi() -> Dict[str, str]:
    """Construye las cabeceras de autenticación para llamadas a Power BI API."""
    try:
        token = _get_pbi_token()
        return {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    except Exception as e:
        raise Exception(f"No se pudieron obtener las cabeceras de autenticación para Power BI API: {e}") from e

# ==========================================
# ==== FUNCIONES DE ACCIÓN PARA POWER BI ====
# ==========================================
# (Resto de las funciones listar_workspaces, obtener_workspace, etc., sin cambios respecto a v2)
def listar_workspaces(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    auth_headers = _get_auth_headers_for_pbi(); expand: Optional[List[str]] = parametros.get("expand")
    url = f"{PBI_BASE_URL}/groups"; params_query: Dict[str, Any] = {};
    if expand and isinstance(expand, list): params_query['$expand'] = ','.join(expand)
    logger.info(f"Listando workspaces PBI (Expand: {expand})"); return hacer_llamada_api("GET", url, auth_headers, params=params_query or None, timeout=PBI_TIMEOUT)

def obtener_workspace(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    auth_headers = _get_auth_headers_for_pbi(); workspace_id: Optional[str] = parametros.get("workspace_id")
    if not workspace_id: raise ValueError("'workspace_id' requerido."); url = f"{PBI_BASE_URL}/groups/{workspace_id}"
    logger.info(f"Obteniendo workspace PBI: {workspace_id}"); return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def listar_dashboards(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    auth_headers = _get_auth_headers_for_pbi(); workspace_id: Optional[str] = parametros.get("workspace_id")
    if not workspace_id: raise ValueError("'workspace_id' requerido."); url = f"{PBI_BASE_URL}/groups/{workspace_id}/dashboards"
    logger.info(f"Listando dashboards PBI '{workspace_id}'."); return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def obtener_dashboard(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    auth_headers = _get_auth_headers_for_pbi(); workspace_id: Optional[str] = parametros.get("workspace_id"); dashboard_id: Optional[str] = parametros.get("dashboard_id")
    if not workspace_id: raise ValueError("'workspace_id' requerido."); if not dashboard_id: raise ValueError("'dashboard_id' requerido.")
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/dashboards/{dashboard_id}"; logger.info(f"Obteniendo dashboard PBI: {dashboard_id} en ws {workspace_id}")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def listar_reports(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    auth_headers = _get_auth_headers_for_pbi(); workspace_id: Optional[str] = parametros.get("workspace_id")
    if not workspace_id: raise ValueError("'workspace_id' requerido."); url = f"{PBI_BASE_URL}/groups/{workspace_id}/reports"
    logger.info(f"Listando informes PBI '{workspace_id}'."); return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def obtener_reporte(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    auth_headers = _get_auth_headers_for_pbi(); workspace_id: Optional[str] = parametros.get("workspace_id"); report_id: Optional[str] = parametros.get("report_id")
    if not workspace_id: raise ValueError("'workspace_id' requerido."); if not report_id: raise ValueError("'report_id' requerido.")
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/reports/{report_id}"; logger.info(f"Obteniendo informe PBI: {report_id} en ws {workspace_id}")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def listar_datasets(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    auth_headers = _get_auth_headers_for_pbi(); workspace_id: Optional[str] = parametros.get("workspace_id")
    if not workspace_id: raise ValueError("'workspace_id' requerido."); url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets"
    logger.info(f"Listando datasets PBI '{workspace_id}'."); return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def obtener_dataset(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    auth_headers = _get_auth_headers_for_pbi(); workspace_id: Optional[str] = parametros.get("workspace_id"); dataset_id: Optional[str] = parametros.get("dataset_id")
    if not workspace_id: raise ValueError("'workspace_id' requerido."); if not dataset_id: raise ValueError("'dataset_id' requerido.")
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}"; logger.info(f"Obteniendo dataset PBI: {dataset_id} en ws {workspace_id}")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def refrescar_dataset(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    auth_headers = _get_auth_headers_for_pbi(); workspace_id: Optional[str] = parametros.get("workspace_id"); dataset_id: Optional[str] = parametros.get("dataset_id"); notify_option: Optional[str] = parametros.get("notifyOption")
    if not workspace_id: raise ValueError("'workspace_id' requerido."); if not dataset_id: raise ValueError("'dataset_id' requerido.")
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}/refreshes"; body: Dict[str, Any] = {};
    if notify_option: body["notifyOption"] = notify_option
    logger.info(f"Iniciando refresco dataset PBI '{dataset_id}' en ws '{workspace_id}'")
    response = hacer_llamada_api("POST", url, auth_headers, json_data=body or None, timeout=PBI_TIMEOUT, expect_json=False)
    if isinstance(response, requests.Response):
        request_id = response.headers.get('RequestId')
        if response.status_code == 202: logger.info(f"Refresco PBI '{dataset_id}' iniciado. RequestId: {request_id}"); return {"status": "Refresco iniciado", "status_code": 202, "requestId": request_id}
        else: logger.error(f"Respuesta inesperada refresco PBI '{dataset_id}'. Status: {response.status_code}. RequestId: {request_id}. Body: {response.text[:200]}"); try: error_body = response.json()
        except json.JSONDecodeError: error_body = response.text; return {"status": "Error", "status_code": response.status_code, "requestId": request_id, "error": error_body}
    else: logger.error(f"Respuesta inesperada helper refresco PBI: {type(response)}"); raise Exception("Error interno refresco PBI.")

def obtener_estado_refresco_dataset(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    auth_headers = _get_auth_headers_for_pbi(); workspace_id: Optional[str] = parametros.get("workspace_id"); dataset_id: Optional[str] = parametros.get("dataset_id"); top: int = int(parametros.get("top", 1))
    if not workspace_id: raise ValueError("'workspace_id' requerido."); if not dataset_id: raise ValueError("'dataset_id' requerido.")
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}/refreshes"; params_query = {'$top': top};
    logger.info(f"Obteniendo estado {top} refrescos PBI dataset '{dataset_id}'"); return hacer_llamada_api("GET", url, auth_headers, params=params_query, timeout=PBI_TIMEOUT)

def obtener_embed_url(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    auth_headers = _get_auth_headers_for_pbi(); workspace_id: Optional[str] = parametros.get("workspace_id"); report_id: Optional[str] = parametros.get("report_id")
    if not workspace_id: raise ValueError("'workspace_id' requerido."); if not report_id: raise ValueError("'report_id' requerido.")
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/reports/{report_id}"; logger.info(f"Obteniendo info informe PBI '{report_id}' para embed URL")
    try:
        report_data = hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT); embed_url = report_data.get("embedUrl")
        if embed_url: logger.info(f"Obtenida URL base PBI '{report_id}'"); return {"embedUrl": embed_url, "reportId": report_data.get("id"), "reportName": report_data.get("name"), "datasetId": report_data.get("datasetId"), "warning": "Se requiere Embed Token."}
        else: logger.warning(f"No se encontró 'embedUrl' PBI '{report_id}'."); return {"error": f"No se encontró embedUrl para informe {report_id}."}
    except requests.exceptions.RequestException as e: logger.error(f"Error API obteniendo informe PBI '{report_id}': {e}", exc_info=True); error_detail = e.response.json() if e.response and e.response.content else str(e); return {"error": f"Error API obteniendo informe {report_id}", "detail": error_detail, "status_code": e.response.status_code if e.response else None}
    except Exception as e: logger.error(f"Error inesperado obteniendo informe PBI '{report_id}': {e}", exc_info=True); raise

# --- FIN DEL MÓDULO actions/power_bi.py ---
