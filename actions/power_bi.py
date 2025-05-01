# actions/power_bi.py (Refactorizado y Corregido - Final)

import logging
import os
import requests
import json
# Corregido: Añadir Any
from typing import Dict, List, Optional, Union, Any

# Importar Credential de Azure Identity
from azure.identity import ClientSecretCredential, CredentialUnavailableError

# Importar helper HTTP
try:
    from helpers.http_client import hacer_llamada_api
except ImportError:
    logger = logging.getLogger("azure.functions")
    logger.error("Error importando http_client en Power BI.")
    def hacer_llamada_api(*args, **kwargs): raise NotImplementedError("Helper no importado")

# Usar el logger de la función principal
logger = logging.getLogger("azure.functions")

# --- Constantes y Variables de Entorno Específicas ---
try:
    CLIENT_ID = os.environ['CLIENT_ID']
    TENANT_ID = os.environ['TENANT_ID']
    CLIENT_SECRET = os.environ['CLIENT_SECRET']
except KeyError as e:
    logger.critical(f"Error Crítico: Falta variable de entorno esencial para Power BI: {e}")
    raise ValueError(f"Configuración incompleta para Power BI: falta {e}")

PBI_BASE_URL = "https://api.powerbi.com/v1.0/myorg"
PBI_SCOPE = "https://analysis.windows.net/powerbi/api/.default"
PBI_TIMEOUT = 60

# --- Helper de Autenticación (Específico para este módulo) ---
_credential_pbi: Optional[ClientSecretCredential] = None
_cached_pbi_token: Optional[str] = None

def _get_pbi_token() -> str:
    """Obtiene un token para Power BI API usando Client Credentials."""
    global _credential_pbi, _cached_pbi_token
    if _cached_pbi_token: return _cached_pbi_token
    if not _credential_pbi:
        logger.info("Creando credencial ClientSecretCredential para Power BI.")
        _credential_pbi = ClientSecretCredential(tenant_id=TENANT_ID, client_id=CLIENT_ID, client_secret=CLIENT_SECRET)
    try:
        logger.info(f"Solicitando token para Power BI con scope: {PBI_SCOPE}")
        assert _credential_pbi is not None
        token_info = _credential_pbi.get_token(PBI_SCOPE)
        _cached_pbi_token = token_info.token
        logger.info("Token para Power BI obtenido exitosamente.")
        return _cached_pbi_token
    except Exception as e:
        logger.error(f"Error obteniendo token de Power BI: {e}", exc_info=True)
        raise Exception(f"Error obteniendo token Power BI: {e}")

def _get_auth_headers_for_pbi() -> Dict[str, str]:
    token = _get_pbi_token()
    return {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

# ---- POWER BI ----
# Funciones con parámetros reordenados y usando auth interna + helper HTTP

def listar_workspaces(expand: Optional[List[str]] = None, headers: Optional[Dict[str, str]] = None) -> dict:
    """Lista los workspaces (grupos) de Power BI."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups"; params: Dict[str, Any] = {};
    if expand: params['$expand'] = ','.join(expand)
    logger.info("Listando workspaces de Power BI.")
    return hacer_llamada_api("GET", url, auth_headers, params=params or None, timeout=PBI_TIMEOUT)

def obtener_workspace(workspace_id: str, headers: Optional[Dict[str, str]] = None) -> dict:
    """Obtiene un workspace de Power BI específico."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}"
    logger.info(f"Obteniendo workspace PBI: {workspace_id}")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def listar_dashboards(workspace_id: str, headers: Optional[Dict[str, str]] = None) -> dict:
    """Lista los dashboards en un workspace de Power BI."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/dashboards"
    logger.info(f"Listando dashboards del workspace PBI '{workspace_id}'.")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def obtener_dashboard(workspace_id: str, dashboard_id: str, headers: Optional[Dict[str, str]] = None) -> dict:
    """Obtiene un dashboard de Power BI específico."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/dashboards/{dashboard_id}"
    logger.info(f"Obteniendo dashboard PBI: {dashboard_id}")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def listar_reports(workspace_id: str, headers: Optional[Dict[str, str]] = None) -> dict:
    """Lista los informes en un workspace de Power BI."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/reports"
    logger.info(f"Listando informes del workspace PBI '{workspace_id}'.")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def obtener_reporte(workspace_id: str, report_id: str, headers: Optional[Dict[str, str]] = None) -> dict:
    """Obtiene un reporte de Power BI específico."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/reports/{report_id}"
    logger.info(f"Obteniendo reporte PBI: {report_id}")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def listar_datasets(workspace_id: str, headers: Optional[Dict[str, str]] = None) -> dict:
    """Lista los datasets en un workspace de Power BI."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets"
    logger.info(f"Listando datasets del workspace PBI '{workspace_id}'.")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def obtener_dataset(workspace_id: str, dataset_id: str, headers: Optional[Dict[str, str]] = None) -> dict:
    """Obtiene un dataset de Power BI específico."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}"
    logger.info(f"Obteniendo dataset PBI: {dataset_id}")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)

def refrescar_dataset(workspace_id: str, dataset_id: str, headers: Optional[Dict[str, str]] = None) -> dict:
    """Inicia un refresco de un dataset de Power BI."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}/refreshes"
    body: Dict[str, Any] = {} # Anotación corregida
    logger.info(f"Iniciando refresco dataset PBI '{dataset_id}'")
    # Hacer llamada API devuelve None si 202, necesitamos el status code real
    try:
        response = requests.post(url, headers=auth_headers, json=body, timeout=PBI_TIMEOUT)
        if response.status_code == 202:
             logger.info(f"Refresco del dataset PBI '{dataset_id}' iniciado (encolado).")
             return {"status": "Refresh iniciado", "code": response.status_code}
        else:
            response.raise_for_status()
            logger.warning(f"Respuesta inesperada refrescar PBI '{dataset_id}'. Status: {response.status_code}")
            # Si no es 202 ni error, ¿qué devolvemos? Un status genérico.
            return {"status": f"Respuesta inesperada {response.status_code}", "code": response.status_code}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en refrescar_dataset (PBI) {dataset_id}: {e}", exc_info=True); raise Exception(f"Error API refrescando dataset PBI: {e}")
    except Exception as e: logger.error(f"Error inesperado en refrescar_dataset (PBI) {dataset_id}: {e}", exc_info=True); raise

def obtener_estado_refresco_dataset(workspace_id: str, dataset_id: str, top: int = 1, headers: Optional[Dict[str, str]] = None) -> dict:
    """Obtiene el historial de refrescos (por defecto el último) de un dataset."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}/refreshes"
    params = {'$top': top}
    logger.info(f"Obteniendo estado refresco PBI dataset '{dataset_id}'")
    return hacer_llamada_api("GET", url, auth_headers, params=params, timeout=PBI_TIMEOUT)

def obtener_embed_url(workspace_id: str, report_id: str, headers: Optional[Dict[str, str]] = None) -> dict:
    """Obtiene la URL base de un informe (NO incluye Embed Token)."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/reports/{report_id}"
    logger.info(f"Obteniendo info reporte PBI '{report_id}' para embed URL")
    data = hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)
    embed_url = data.get("embedUrl")
    if embed_url:
        logger.info(f"Obtenida URL (base) para informe PBI '{report_id}': {embed_url}")
        return {"embedUrl": embed_url, "reportId": data.get("id"), "datasetId": data.get("datasetId"), "reportName": data.get("name"), "warning": "Requires Embed Token generation for actual embedding."}
    else:
        logger.warning(f"No se encontró 'embedUrl' para informe PBI '{report_id}'."); return {"error": "No se encontró embedUrl"}
