# actions/power_bi.py (Refactorizado)

import logging
import os
import requests # Solo para tipos de excepción
import json
from typing import Dict, List, Optional, Union, Any

# Importar Credential de Azure Identity para autenticación con Power BI API
try:
    from azure.identity import ClientSecretCredential, CredentialUnavailableError
except ImportError:
    logging.critical("Error CRÍTICO: Falta 'azure-identity'. Instala con 'pip install azure-identity'.")
    class ClientSecretCredential: pass # Mock simple
    class CredentialUnavailableError(Exception): pass # Mock simple

# Importar helper HTTP y constantes (aunque BASE_URL no se use directamente)
try:
    # Asume que shared está un nivel arriba de actions
    from ..shared.helpers.http_client import hacer_llamada_api
    from ..shared.constants import BASE_URL, GRAPH_API_TIMEOUT # Timeout base
except ImportError as e:
    logging.critical(f"Error CRÍTICO importando helpers/constantes en Power BI: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs):
        raise NotImplementedError("Dependencia 'hacer_llamada_api' no importada correctamente.")

# Usar el logger estándar de Azure Functions
logger = logging.getLogger("azure.functions")

# --- Constantes y Variables de Entorno Específicas para Power BI API ---
# Necesarias para la autenticación con ClientSecretCredential
try:
    # Usar nombres específicos si las credenciales son diferentes de las de Graph/Management
    PBI_CLIENT_ID = os.environ['AZURE_CLIENT_ID_PBI'] # Ejemplo: AZURE_CLIENT_ID_PBI
    PBI_TENANT_ID = os.environ['AZURE_TENANT_ID'] # Usualmente el mismo tenant
    PBI_CLIENT_SECRET = os.environ['AZURE_CLIENT_SECRET_PBI'] # Ejemplo: AZURE_CLIENT_SECRET_PBI
except KeyError as e:
    logger.critical(f"Error Crítico: Falta variable de entorno esencial para autenticación Power BI: {e}")
    raise ValueError(f"Configuración incompleta para Power BI API: falta {e}")

# Endpoints y configuración para Power BI REST API
PBI_BASE_URL = "https://api.powerbi.com/v1.0/myorg" # URL base para API Power BI
PBI_SCOPE = "https://analysis.windows.net/powerbi/api/.default" # Scope específico para Power BI
PBI_TIMEOUT = max(GRAPH_API_TIMEOUT, 60) # Timeout específico (ej: 60s)

# --- Helper de Autenticación (Específico para este módulo) ---
# Cache simple para la credencial y el token PBI
_credential_pbi: Optional[ClientSecretCredential] = None
_cached_pbi_token: Optional[str] = None # TODO: Añadir manejo de expiración

def _get_pbi_token() -> str:
    """Obtiene un token de acceso para Power BI API usando Client Credentials."""
    global _credential_pbi, _cached_pbi_token

    # TODO: Implementar chequeo de expiración del token cacheado
    if _cached_pbi_token:
        # logger.debug("Usando token Power BI cacheado.")
        return _cached_pbi_token

    if not _credential_pbi:
        logger.info("Creando credencial ClientSecretCredential para Power BI API.")
        try:
            _credential_pbi = ClientSecretCredential(
                tenant_id=PBI_TENANT_ID,
                client_id=PBI_CLIENT_ID,
                client_secret=PBI_CLIENT_SECRET
            )
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
         logger.critical(f"Credencial no disponible para obtener token PBI: {cred_err}", exc_info=True)
         raise Exception(f"Credencial Power BI no disponible: {cred_err}") from cred_err
    except Exception as e:
        logger.error(f"Error inesperado obteniendo token de Power BI: {e}", exc_info=True)
        raise Exception(f"Error obteniendo token Power BI: {e}") from e

def _get_auth_headers_for_pbi() -> Dict[str, str]:
    """Construye las cabeceras de autenticación para llamadas a Power BI API."""
    try:
        token = _get_pbi_token()
        # Power BI API usa 'Authorization: Bearer <token>'
        return {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    except Exception as e:
        raise Exception(f"No se pudieron obtener las cabeceras de autenticación para Power BI API: {e}") from e

# ==========================================
# ==== FUNCIONES DE ACCIÓN PARA POWER BI ====
# ==========================================
# Usan la firma (parametros: Dict[str, Any], headers: Dict[str, str])
# PERO usan la autenticación interna (_get_auth_headers_for_pbi).
# Los 'headers' de entrada (Graph API) se ignoran aquí.

def listar_workspaces(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista los workspaces (grupos) de Power BI a los que tiene acceso la credencial.

    Args:
        parametros (Dict[str, Any]): Opcional: 'expand' (List[str], ej. ['reports', 'datasets']).
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: Respuesta de Power BI API, usualmente {'value': [...]}.
    """
    auth_headers = _get_auth_headers_for_pbi() # Usar auth PBI
    expand: Optional[List[str]] = parametros.get("expand")

    url = f"{PBI_BASE_URL}/groups" # Endpoint para listar grupos/workspaces
    params_query: Dict[str, Any] = {}
    if expand and isinstance(expand, list):
        params_query['$expand'] = ','.join(expand)

    logger.info(f"Listando workspaces de Power BI (Expand: {expand})")
    # Usar helper con auth PBI
    return hacer_llamada_api("GET", url, auth_headers, params=params_query or None, timeout=PBI_TIMEOUT)


def obtener_workspace(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene los detalles de un workspace de Power BI específico por su ID.

    Args:
        parametros (Dict[str, Any]): Debe contener 'workspace_id'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: El objeto del workspace de Power BI API.
    """
    auth_headers = _get_auth_headers_for_pbi()
    workspace_id: Optional[str] = parametros.get("workspace_id")
    if not workspace_id: raise ValueError("Parámetro 'workspace_id' es requerido.")

    url = f"{PBI_BASE_URL}/groups/{workspace_id}"
    logger.info(f"Obteniendo workspace Power BI: {workspace_id}")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)


def listar_dashboards(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista los dashboards dentro de un workspace de Power BI.

    Args:
        parametros (Dict[str, Any]): Debe contener 'workspace_id'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: Respuesta de Power BI API, usualmente {'value': [...]}.
    """
    auth_headers = _get_auth_headers_for_pbi()
    workspace_id: Optional[str] = parametros.get("workspace_id")
    if not workspace_id: raise ValueError("Parámetro 'workspace_id' es requerido.")

    url = f"{PBI_BASE_URL}/groups/{workspace_id}/dashboards"
    logger.info(f"Listando dashboards del workspace Power BI '{workspace_id}'.")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)


def obtener_dashboard(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene los detalles de un dashboard específico.

    Args:
        parametros (Dict[str, Any]): Debe contener 'workspace_id', 'dashboard_id'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: El objeto del dashboard.
    """
    auth_headers = _get_auth_headers_for_pbi()
    workspace_id: Optional[str] = parametros.get("workspace_id")
    dashboard_id: Optional[str] = parametros.get("dashboard_id")
    if not workspace_id: raise ValueError("Parámetro 'workspace_id' es requerido.")
    if not dashboard_id: raise ValueError("Parámetro 'dashboard_id' es requerido.")

    url = f"{PBI_BASE_URL}/groups/{workspace_id}/dashboards/{dashboard_id}"
    logger.info(f"Obteniendo dashboard Power BI: {dashboard_id} en workspace {workspace_id}")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)


def listar_reports(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista los informes (reports) dentro de un workspace de Power BI.

    Args:
        parametros (Dict[str, Any]): Debe contener 'workspace_id'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: Respuesta de Power BI API, usualmente {'value': [...]}.
    """
    auth_headers = _get_auth_headers_for_pbi()
    workspace_id: Optional[str] = parametros.get("workspace_id")
    if not workspace_id: raise ValueError("Parámetro 'workspace_id' es requerido.")

    url = f"{PBI_BASE_URL}/groups/{workspace_id}/reports"
    logger.info(f"Listando informes del workspace Power BI '{workspace_id}'.")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)


def obtener_reporte(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene los detalles de un informe (report) específico.

    Args:
        parametros (Dict[str, Any]): Debe contener 'workspace_id', 'report_id'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: El objeto del informe.
    """
    auth_headers = _get_auth_headers_for_pbi()
    workspace_id: Optional[str] = parametros.get("workspace_id")
    report_id: Optional[str] = parametros.get("report_id")
    if not workspace_id: raise ValueError("Parámetro 'workspace_id' es requerido.")
    if not report_id: raise ValueError("Parámetro 'report_id' es requerido.")

    url = f"{PBI_BASE_URL}/groups/{workspace_id}/reports/{report_id}"
    logger.info(f"Obteniendo informe Power BI: {report_id} en workspace {workspace_id}")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)


def listar_datasets(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista los conjuntos de datos (datasets) dentro de un workspace de Power BI.

    Args:
        parametros (Dict[str, Any]): Debe contener 'workspace_id'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: Respuesta de Power BI API, usualmente {'value': [...]}.
    """
    auth_headers = _get_auth_headers_for_pbi()
    workspace_id: Optional[str] = parametros.get("workspace_id")
    if not workspace_id: raise ValueError("Parámetro 'workspace_id' es requerido.")

    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets"
    logger.info(f"Listando datasets del workspace Power BI '{workspace_id}'.")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)


def obtener_dataset(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene los detalles de un conjunto de datos (dataset) específico.

    Args:
        parametros (Dict[str, Any]): Debe contener 'workspace_id', 'dataset_id'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: El objeto del dataset.
    """
    auth_headers = _get_auth_headers_for_pbi()
    workspace_id: Optional[str] = parametros.get("workspace_id")
    dataset_id: Optional[str] = parametros.get("dataset_id")
    if not workspace_id: raise ValueError("Parámetro 'workspace_id' es requerido.")
    if not dataset_id: raise ValueError("Parámetro 'dataset_id' es requerido.")

    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}"
    logger.info(f"Obteniendo dataset Power BI: {dataset_id} en workspace {workspace_id}")
    return hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)


def refrescar_dataset(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Inicia un refresco (actualización) de un conjunto de datos. Operación asíncrona.

    Args:
        parametros (Dict[str, Any]): Debe contener 'workspace_id', 'dataset_id'.
                                     Opcional: 'notifyOption' (ej. 'MailOnCompletion', 'NoNotification').
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: Confirmación de inicio de refresco (usualmente 202 Accepted).
    """
    auth_headers = _get_auth_headers_for_pbi()
    workspace_id: Optional[str] = parametros.get("workspace_id")
    dataset_id: Optional[str] = parametros.get("dataset_id")
    notify_option: Optional[str] = parametros.get("notifyOption") # Ej: MailOnCompletion

    if not workspace_id: raise ValueError("Parámetro 'workspace_id' es requerido.")
    if not dataset_id: raise ValueError("Parámetro 'dataset_id' es requerido.")

    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}/refreshes"
    # El body puede incluir notifyOption
    body: Dict[str, Any] = {}
    if notify_option:
        body["notifyOption"] = notify_option

    logger.info(f"Iniciando refresco dataset Power BI '{dataset_id}' en workspace '{workspace_id}'")

    # POST a /refreshes devuelve 202 Accepted. Usar helper con expect_json=False.
    response = hacer_llamada_api("POST", url, auth_headers, json_data=body or None, timeout=PBI_TIMEOUT, expect_json=False)

    if isinstance(response, requests.Response) and response.status_code == 202:
        # El header 'RequestId' puede ser útil para seguimiento
        request_id = response.headers.get('RequestId')
        logger.info(f"Refresco del dataset PBI '{dataset_id}' iniciado (encolado). RequestId: {request_id}")
        return {"status": "Refresco iniciado", "status_code": response.status_code, "requestId": request_id}
    elif isinstance(response, requests.Response):
         # Si la API devuelve otro status (ej. 429 Too Many Requests)
         logger.error(f"Respuesta inesperada al iniciar refresco PBI '{dataset_id}'. Status: {response.status_code}. Body: {response.text[:200]}")
         # Intentar devolver el cuerpo del error si es JSON
         try: error_body = response.json()
         except json.JSONDecodeError: error_body = response.text
         return {"status": "Error", "status_code": response.status_code, "error": error_body}
    else:
         logger.error(f"Respuesta inesperada del helper al iniciar refresco PBI: {type(response)}")
         raise Exception("Error interno al procesar la solicitud de refresco PBI.")


def obtener_estado_refresco_dataset(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene el historial de refrescos de un dataset (por defecto el último).

    Args:
        parametros (Dict[str, Any]): Debe contener 'workspace_id', 'dataset_id'.
                                     Opcional: 'top' (int, default 1, para obtener los N últimos).
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: Respuesta de Power BI API con el historial, usualmente {'value': [...]}.
    """
    auth_headers = _get_auth_headers_for_pbi()
    workspace_id: Optional[str] = parametros.get("workspace_id")
    dataset_id: Optional[str] = parametros.get("dataset_id")
    top: int = int(parametros.get("top", 1)) # Default a 1 para obtener solo el último

    if not workspace_id: raise ValueError("Parámetro 'workspace_id' es requerido.")
    if not dataset_id: raise ValueError("Parámetro 'dataset_id' es requerido.")

    url = f"{PBI_BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}/refreshes"
    params_query = {'$top': top}

    logger.info(f"Obteniendo estado de los últimos {top} refrescos PBI dataset '{dataset_id}'")
    return hacer_llamada_api("GET", url, auth_headers, params=params_query, timeout=PBI_TIMEOUT)


def obtener_embed_url(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene la URL base para embeber un informe.
    NOTA: Esto NO genera el Embed Token necesario para la visualización final.

    Args:
        parametros (Dict[str, Any]): Debe contener 'workspace_id', 'report_id'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: Información del informe incluyendo 'embedUrl'.
                       O un dict con 'error' si no se encuentra.
    """
    auth_headers = _get_auth_headers_for_pbi()
    workspace_id: Optional[str] = parametros.get("workspace_id")
    report_id: Optional[str] = parametros.get("report_id")

    if not workspace_id: raise ValueError("Parámetro 'workspace_id' es requerido.")
    if not report_id: raise ValueError("Parámetro 'report_id' es requerido.")

    # Obtener detalles del informe para extraer embedUrl
    url = f"{PBI_BASE_URL}/groups/{workspace_id}/reports/{report_id}"
    logger.info(f"Obteniendo información del informe Power BI '{report_id}' para obtener embed URL")

    try:
        report_data = hacer_llamada_api("GET", url, auth_headers, timeout=PBI_TIMEOUT)
        embed_url = report_data.get("embedUrl")

        if embed_url:
            logger.info(f"Obtenida URL base para embeber informe PBI '{report_id}': {embed_url}")
            # Devolver información útil
            return {
                "embedUrl": embed_url,
                "reportId": report_data.get("id"),
                "reportName": report_data.get("name"),
                "datasetId": report_data.get("datasetId"),
                "warning": "Esta es solo la URL base. Se requiere generar un Embed Token por separado para la visualización."
            }
        else:
            logger.warning(f"No se encontró 'embedUrl' en la respuesta para el informe PBI '{report_id}'.")
            return {"error": f"No se encontró 'embedUrl' para el informe {report_id}."}
    except requests.exceptions.RequestException as e:
        # Manejar error si el informe no se encuentra o hay problema de permisos
        logger.error(f"Error API obteniendo informe PBI '{report_id}': {e}", exc_info=True)
        raise Exception(f"Error API obteniendo informe PBI '{report_id}': {e}") from e
    except Exception as e:
         logger.error(f"Error inesperado obteniendo informe PBI '{report_id}': {e}", exc_info=True)
         raise

# --- FIN DEL MÓDULO actions/power_bi.py ---
