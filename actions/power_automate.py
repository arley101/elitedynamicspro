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
# Necesitamos estas variables para la API de Azure Management y la autenticación
try:
    AZURE_SUBSCRIPTION_ID = os.environ['AZURE_SUBSCRIPTION_ID']
    AZURE_RESOURCE_GROUP = os.environ['AZURE_RESOURCE_GROUP']
    # AZURE_LOCATION es necesaria solo para crear_flow
    AZURE_LOCATION = os.environ.get('AZURE_LOCATION') # Puede ser opcional si no se usa crear_flow
    # Credenciales para obtener token para Azure Management API
    CLIENT_ID = os.environ['CLIENT_ID']
    TENANT_ID = os.environ['TENANT_ID']
    CLIENT_SECRET = os.environ['CLIENT_SECRET']
except KeyError as e:
    logger.critical(f"Error Crítico: Falta variable de entorno esencial para Power Automate: {e}")
    # Si falta, las funciones de este módulo fallarán
    raise ValueError(f"Configuración incompleta para Power Automate: falta {e}")

AZURE_MGMT_BASE_URL = "https://management.azure.com"
# El scope para obtener un token para Azure Resource Management API
AZURE_MGMT_SCOPE = "https://management.azure.com/.default"
# API Version (puede necesitar actualización en el futuro)
LOGIC_API_VERSION = "2019-05-01" # Usar una versión más reciente si es posible/necesario
AZURE_MGMT_TIMEOUT = 60 # Timeout un poco más largo para llamadas de gestión

# --- Helper de Autenticación (Específico para este módulo) ---
_credential = None
_cached_mgmt_token = None # Cache simple en memoria (válido solo para una invocación)

def _get_azure_mgmt_token() -> str:
    """Obtiene un token para Azure Management API usando Client Credentials."""
    global _credential, _cached_mgmt_token
    # TODO: Añadir lógica de expiración si se usa caché entre invocaciones (poco probable)

    if _cached_mgmt_token:
        logger.info("Usando token de Azure Management cacheado (solo válido en esta invocación).")
        return _cached_mgmt_token

    if not _credential:
        logger.info("Creando credencial ClientSecretCredential para Azure Management.")
        _credential = ClientSecretCredential(
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET
        )
    try:
        logger.info(f"Solicitando token para Azure Management con scope: {AZURE_MGMT_SCOPE}")
        token_info = _credential.get_token(AZURE_MGMT_SCOPE)
        _cached_mgmt_token = token_info.token # Cachear token obtenido
        logger.info("Token para Azure Management obtenido exitosamente.")
        return _cached_mgmt_token
    except CredentialUnavailableError as cred_err:
        logger.error(f"Error de credencial al obtener token de Azure Management: {cred_err}", exc_info=True)
        raise Exception(f"Error de credencial Azure: {cred_err}")
    except Exception as e:
        logger.error(f"Error inesperado al obtener token de Azure Management: {e}", exc_info=True)
        raise Exception(f"Error inesperado obteniendo token Azure: {e}")

def _get_auth_headers_for_mgmt() -> Dict[str, str]:
    """Obtiene las cabeceras de autorización para Azure Management API."""
    token = _get_azure_mgmt_token()
    return {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

# ---- POWER AUTOMATE (Flows) ----
# NOTA: Estas funciones ahora ignoran el parámetro 'headers' si se pasa desde main,
#       y usan su propia autenticación (_get_auth_headers_for_mgmt).

def listar_flows(headers: Optional[Dict[str, str]] = None, suscripcion_id: Optional[str] = None, grupo_recurso: Optional[str] = None) -> dict:
    """Lista los flujos (Logic Apps Standard/Consumption o Power Automate) en un grupo de recursos."""
    # Ignora 'headers' pasados, usa auth interna para Azure Mgmt API
    auth_headers = _get_auth_headers_for_mgmt()
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
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_flows: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_flows: {e}", exc_info=True)
        raise

def obtener_flow(headers: Optional[Dict[str, str]] = None, nombre_flow: str, suscripcion_id: Optional[str] = None, grupo_recurso: Optional[str] = None) -> dict:
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
        data = response.json()
        logger.info(f"Obtenido flow '{nombre_flow}'.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en obtener_flow: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_flow: {e}", exc_info=True)
        raise

def crear_flow(nombre_flow: str, definicion_flow: dict, headers: Optional[Dict[str, str]] = None, ubicacion: Optional[str] = None, suscripcion_id: Optional[str] = None, grupo_recurso: Optional[str] = None) -> dict:
    """Crea un nuevo flujo."""
    auth_headers = _get_auth_headers_for_mgmt()
    sid = suscripcion_id or AZURE_SUBSCRIPTION_ID
    rg = grupo_recurso or AZURE_RESOURCE_GROUP
    loc = ubicacion or AZURE_LOCATION
    if not loc:
        raise ValueError("Se requiere 'ubicacion' o la variable de entorno 'AZURE_LOCATION' para crear un flow.")
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    body = {
        "location": loc,
        "properties": {
            "definition": definicion_flow # La definición del flujo (JSON)
            # Faltan otros parámetros posiblemente necesarios como state, parameters, etc.
        }
    }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PUT {url} (Creando flow '{nombre_flow}')")
        response = requests.put(url, headers=auth_headers, json=body, timeout=AZURE_MGMT_TIMEOUT * 2) # Mayor timeout para PUT
        response.raise_for_status() # 201 Created o 200 OK
        data = response.json()
        logger.info(f"Flujo '{nombre_flow}' creado en '{rg}'.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en crear_flow: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_flow: {e}", exc_info=True)
        raise

def actualizar_flow(nombre_flow: str, definicion_flow: dict, headers: Optional[Dict[str, str]] = None, suscripcion_id: Optional[str] = None, grupo_recurso: Optional[str] = None) -> dict:
    """Actualiza un flujo existente (normalmente solo la definición)."""
    auth_headers = _get_auth_headers_for_mgmt()
    sid = suscripcion_id or AZURE_SUBSCRIPTION_ID
    rg = grupo_recurso or AZURE_RESOURCE_GROUP
    # Para actualizar, usualmente se hace PUT al mismo endpoint de creación
    # OJO: PUT reemplaza todo el recurso, PATCH modifica. La API de Logic App usa PUT/PATCH según el caso.
    # Revisar documentación específica, pero PUT es común para la definición.
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    # Body podría necesitar la ubicación y otras propiedades, no solo la definición
    # body = { "properties": { "definition": definicion_flow } } # Simplificado
    # Obteniendo el flow actual para mantener otras propiedades
    try:
         current_flow = obtener_flow(headers=None, nombre_flow=nombre_flow, suscripcion_id=sid, grupo_recurso=rg) # Llamada interna sin headers
         body = current_flow # Usar el objeto actual
         body["properties"]["definition"] = definicion_flow # Sobreescribir solo la definición
     except Exception as get_err:
         logger.error(f"Error obteniendo flow actual para actualizar '{nombre_flow}': {get_err}", exc_info=True)
         raise Exception(f"No se pudo obtener el flow actual para actualizar: {get_err}")

    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PUT {url} (Actualizando flow '{nombre_flow}')")
        response = requests.put(url, headers=auth_headers, json=body, timeout=AZURE_MGMT_TIMEOUT * 2)
        response.raise_for_status() # 200 OK
        data = response.json()
        logger.info(f"Flujo '{nombre_flow}' actualizado en '{rg}'.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en actualizar_flow: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en actualizar_flow: {e}", exc_info=True)
        raise

def eliminar_flow(headers: Optional[Dict[str, str]] = None, nombre_flow: str, suscripcion_id: Optional[str] = None, grupo_recurso: Optional[str] = None) -> dict:
    """Elimina un flujo."""
    auth_headers = _get_auth_headers_for_mgmt()
    sid = suscripcion_id or AZURE_SUBSCRIPTION_ID
    rg = grupo_recurso or AZURE_RESOURCE_GROUP
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando flow '{nombre_flow}')")
        response = requests.delete(url, headers=auth_headers, timeout=AZURE_MGMT_TIMEOUT)
        response.raise_for_status() # 200 OK o 204 No Content
        logger.info(f"Flujo '{nombre_flow}' eliminado de '{rg}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en eliminar_flow: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_flow: {e}", exc_info=True)
        raise

def ejecutar_flow(headers: Optional[Dict[str, str]] = None, flow_url: str, parametros: Optional[dict] = None) -> dict:
    """Ejecuta un flujo a través de su URL de desencadenador HTTP."""
    # --- IMPORTANTE ---
    # Esta función llama a la URL del *desencadenador* del flujo (ej: 'Cuando se recibe una solicitud HTTP').
    # La autenticación aquí DEPENDE de cómo esté configurado ESE desencadenador en Power Automate/Logic App.
    # Puede que no necesite token, o necesite un API Key, o un token específico.
    # Los 'headers' pasados (con token de Graph) probablemente NO sirvan aquí.
    # Asumiremos por ahora que el trigger no requiere autenticación compleja o que la URL ya la incluye.
    request_headers = {'Content-Type': 'application/json'} # Header básico
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {flow_url} (Ejecutando flow trigger)")
        # No usamos los headers de Graph/Mgmt por defecto. Si el trigger requiere auth específica, debe manejarse.
        response = requests.post(flow_url, headers=request_headers, json=parametros if parametros else {}, timeout=AZURE_MGMT_TIMEOUT)
        response.raise_for_status() # Espera 202 Accepted normalmente
        logger.info(f"Flujo en URL '{flow_url}' ejecutado (Triggered). Status: {response.status_code}")
        # La respuesta del trigger puede variar, a veces está vacía, a veces contiene IDs de ejecución
        try:
            resp_data = response.json()
        except json.JSONDecodeError:
            resp_data = response.text # Si no es JSON
        return {"status": "Ejecutado", "code": response.status_code, "response_body": resp_data}
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en ejecutar_flow: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en ejecutar_flow: {e}", exc_info=True)
        raise

def obtener_estado_ejecucion_flow(headers: Optional[Dict[str, str]] = None, run_id: str, nombre_flow: str, suscripcion_id: Optional[str] = None, grupo_recurso: Optional[str] = None) -> dict:
    """Obtiene el estado de una ejecución específica de un flujo."""
    auth_headers = _get_auth_headers_for_mgmt() # Necesita token de Mgmt
    sid = suscripcion_id or AZURE_SUBSCRIPTION_ID
    rg = grupo_recurso or AZURE_RESOURCE_GROUP
    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}/runs/{run_id}?api-version={LOGIC_API_VERSION}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo estado ejecución '{run_id}' de flow '{nombre_flow}')")
        response = requests.get(url, headers=auth_headers, timeout=AZURE_MGMT_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Obtenido estado ejecución '{run_id}'. Status: {data.get('properties', {}).get('status')}")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en obtener_estado_ejecucion_flow: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_estado_ejecucion_flow: {e}", exc_info=True)
        raise
