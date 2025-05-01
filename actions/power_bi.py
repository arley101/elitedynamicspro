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
# Necesitamos credenciales para obtener token para Power BI API
try:
    CLIENT_ID = os.environ['CLIENT_ID']
    TENANT_ID = os.environ['TENANT_ID']
    CLIENT_SECRET = os.environ['CLIENT_SECRET']
except KeyError as e:
    logger.critical(f"Error Crítico: Falta variable de entorno esencial para Power BI: {e}")
    raise ValueError(f"Configuración incompleta para Power BI: falta {e}")

PBI_BASE_URL = "https://api.powerbi.com/v1.0/myorg" # URL base para Power BI en la organización
# El scope para obtener un token para Power BI API
PBI_SCOPE = "https://analysis.windows.net/powerbi/api/.default"
PBI_TIMEOUT = 60 # Timeout para llamadas a Power BI API

# --- Helper de Autenticación (Específico para este módulo) ---
_credential_pbi = None
_cached_pbi_token = None # Cache simple en memoria

def _get_pbi_token() -> str:
    """Obtiene un token para Power BI API usando Client Credentials."""
    global _credential_pbi, _cached_pbi_token
    # TODO: Añadir lógica de expiración si se usa caché

    if _cached_pbi_token:
        logger.info("Usando token de Power BI cacheado (solo válido en esta invocación).")
        return _cached_pbi_token

    if not _credential_pbi:
        logger.info("Creando credencial ClientSecretCredential para Power BI.")
        # Usar las mismas credenciales de app registration
        _credential_pbi = ClientSecretCredential(
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET
        )
    try:
        logger.info(f"Solicitando token para Power BI con scope: {PBI_SCOPE}")
        token_info = _credential_pbi.get_token(PBI_SCOPE)
        _cached_pbi_token = token_info.token # Cachear token obtenido
        logger.info("Token para Power BI obtenido exitosamente.")
        return _cached_pbi_token
    except CredentialUnavailableError as cred_err:
        logger.error(f"Error de credencial al obtener token de Power BI: {cred_err}", exc_info=True)
        raise Exception(f"Error de credencial Azure para Power BI: {cred_err}")
    except Exception as e:
        logger.error(f"Error inesperado al obtener token de Power BI: {e}", exc_info=True)
        raise Exception(f"Error inesperado obteniendo token Power BI: {e}")

def _get_auth_headers_for_pbi() -> Dict[str, str]:
    """Obtiene las cabeceras de autorización para Power BI API."""
    token = _get_pbi_token()
    return {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

# ---- POWER BI ----
# NOTA: Estas funciones ignoran el parámetro 'headers' (que trae token de Graph)
#       y usan su propia autenticación (_get_auth_headers_for_pbi).

def listar_workspaces(headers: Optional[Dict[str, str]] = None, expand: Optional[List[str]] = None) -> dict:
    """Lista los workspaces (grupos) de Power BI."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups" # Endpoint correcto es /groups
    params = {}
    if expand:
        params['$expand'] = ','.join(expand)
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call (PBI): GET {url}")
        response = requests.get(url, headers=auth_headers, params=params or None, timeout=PBI_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados workspaces de Power BI.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_workspaces (PBI): {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_workspaces (PBI): {e}", exc_info=True)
        raise

def obtener_workspace(workspace_id: str, headers: Optional[Dict[str, str]] = None) -> dict:
    """Obtiene un workspace de Power BI específico."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call (PBI): GET {url}")
        response = requests.get(url, headers=auth_headers, timeout=PBI_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Obtenido workspace PBI: {workspace_id}")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en obtener_workspace (PBI) {workspace_id}: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_workspace (PBI) {workspace_id}: {e}", exc_info=True)
        raise

def listar_dashboards(headers: Optional[Dict[str, str]] = None, workspace_id: str) -> dict:
    """Lista los dashboards en un workspace de Power BI."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/dashboards"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call (PBI): GET {url} (Listando dashboards)")
        response = requests.get(url, headers=auth_headers, timeout=PBI_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados dashboards del workspace PBI '{workspace_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_dashboards (PBI) {workspace_id}: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_dashboards (PBI) {workspace_id}: {e}", exc_info=True)
        raise

def obtener_dashboard(headers: Optional[Dict[str, str]] = None, workspace_id: str, dashboard_id: str) -> dict:
    """Obtiene un dashboard de Power BI específico."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/dashboards/{dashboard_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call (PBI): GET {url} (Obteniendo dashboard '{dashboard_id}')")
        response = requests.get(url, headers=auth_headers, timeout=PBI_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Obtenido dashboard PBI: {dashboard_id}")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en obtener_dashboard (PBI) {dashboard_id}: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_dashboard (PBI) {dashboard_id}: {e}", exc_info=True)
        raise

def listar_reports(headers: Optional[Dict[str, str]] = None, workspace_id: str) -> dict:
    """Lista los informes en un workspace de Power BI."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/reports"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call (PBI): GET {url} (Listando reports)")
        response = requests.get(url, headers=auth_headers, timeout=PBI_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados informes del workspace PBI '{workspace_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_reports (PBI) {workspace_id}: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_reports (PBI) {workspace_id}: {e}", exc_info=True)
        raise

def obtener_reporte(headers: Optional[Dict[str, str]] = None, workspace_id: str, report_id: str) -> dict:
    """Obtiene un reporte de Power BI específico."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/reports/{report_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call (PBI): GET {url} (Obteniendo reporte '{report_id}')")
        response = requests.get(url, headers=auth_headers, timeout=PBI_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Obtenido reporte PBI: {report_id}")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en obtener_reporte (PBI) {report_id}: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_reporte (PBI) {report_id}: {e}", exc_info=True)
        raise

def listar_datasets(headers: Optional[Dict[str, str]] = None, workspace_id: str) -> dict:
    """Lista los datasets en un workspace de Power BI."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call (PBI): GET {url} (Listando datasets)")
        response = requests.get(url, headers=auth_headers, timeout=PBI_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados datasets del workspace PBI '{workspace_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_datasets (PBI) {workspace_id}: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_datasets (PBI) {workspace_id}: {e}", exc_info=True)
        raise

def obtener_dataset(headers: Optional[Dict[str, str]] = None, workspace_id: str, dataset_id: str) -> dict:
    """Obtiene un dataset de Power BI específico."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call (PBI): GET {url} (Obteniendo dataset '{dataset_id}')")
        response = requests.get(url, headers=auth_headers, timeout=PBI_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Obtenido dataset PBI: {dataset_id}")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en obtener_dataset (PBI) {dataset_id}: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_dataset (PBI) {dataset_id}: {e}", exc_info=True)
        raise

def refrescar_dataset(headers: Optional[Dict[str, str]] = None, workspace_id: str, dataset_id: str) -> dict:
    """Inicia un refresco de un dataset de Power BI."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}/refreshes"
    # Body opcional para especificar tipo de refresco, etc.
    # body = {"notifyOption": "MailOnCompletion"}
    body = {} # Sin cuerpo por ahora
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call (PBI): POST {url} (Iniciando refresco dataset '{dataset_id}')")
        current_headers = auth_headers.copy()
        current_headers.setdefault('Content-Type', 'application/json') # Aunque el body esté vacío
        response = requests.post(url, headers=current_headers, json=body, timeout=PBI_TIMEOUT)
        # Power BI devuelve 202 Accepted si el refresco se encola
        if response.status_code == 202:
             logger.info(f"Refresco del dataset PBI '{dataset_id}' iniciado (encolado).")
             # La respuesta 202 no suele tener cuerpo JSON útil
             return {"status": "Refresh iniciado", "code": response.status_code, "response_headers": dict(response.headers)}
        else:
            # Si devuelve otro status, intentar parsear error
            response.raise_for_status()
            # Si no lanza error pero no es 202, loguear y devolver status
            logger.warning(f"Respuesta inesperada al refrescar dataset PBI '{dataset_id}'. Status: {response.status_code}")
            return {"status": f"Respuesta inesperada {response.status_code}", "code": response.status_code}

    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en refrescar_dataset (PBI) {dataset_id}: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en refrescar_dataset (PBI) {dataset_id}: {e}", exc_info=True)
        raise

def obtener_estado_refresco_dataset(headers: Optional[Dict[str, str]] = None, workspace_id: str, dataset_id: str, top: int = 1) -> dict:
    """Obtiene el historial de refrescos (por defecto el último) de un dataset."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}/refreshes?$top={top}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call (PBI): GET {url} (Obteniendo estado refresco dataset '{dataset_id}')")
        response = requests.get(url, headers=auth_headers, timeout=PBI_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Obtenido historial de refresco PBI dataset '{dataset_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en obtener_estado_refresco_dataset (PBI) {dataset_id}: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_estado_refresco_dataset (PBI) {dataset_id}: {e}", exc_info=True)
        raise

# Nota: La función obtener_embed_url es más compleja, ya que generalmente requiere
# generar un Embed Token usando una API diferente o lógica adicional, especialmente
# para escenarios de "Embed for your customers". La llamada directa al report
# solo da la URL base, no el token necesario para embeberlo de forma segura.
# Dejamos la versión simple del archivo original que solo devuelve la URL base.
def obtener_embed_url(headers: Optional[Dict[str, str]] = None, workspace_id: str, report_id: str) -> dict:
    """Obtiene la URL base de un informe (NO incluye Embed Token)."""
    auth_headers = _get_auth_headers_for_pbi()
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/reports/{report_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call (PBI): GET {url} (Obteniendo info reporte '{report_id}' para embed URL)")
        response = requests.get(url, headers=auth_headers, timeout=PBI_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        embed_url = data.get("embedUrl")
        if embed_url:
            logger.info(f"Obtenida URL (base) para informe PBI '{report_id}': {embed_url}")
            # Advertencia: Esto normalmente no es suficiente para embeber.
            return {"embedUrl": embed_url, "reportId": data.get("id"), "datasetId": data.get("datasetId"), "reportName": data.get("name"), "warning": "Requires Embed Token generation for actual embedding."}
        else:
            logger.warning(f"No se encontró 'embedUrl' para informe PBI '{report_id}'.")
            return {"error": "No se encontró embedUrl"}
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en obtener_embed_url (PBI) {report_id}: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_embed_url (PBI) {report_id}: {e}", exc_info=True)
        raise
