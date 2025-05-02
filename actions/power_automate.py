# actions/power_automate.py (Refactorizado)

import logging
import os
import requests # Para ejecutar_flow y tipos de excepción
import json
from typing import Dict, List, Optional, Union, Any

# Importar Credential de Azure Identity para autenticación con Azure Management API
try:
    from azure.identity import ClientSecretCredential, CredentialUnavailableError
except ImportError:
    logging.critical("Error CRÍTICO: Falta 'azure-identity'. Instala con 'pip install azure-identity'.")
    # Definir un mock o lanzar error si azure-identity no está instalado
    class ClientSecretCredential: pass # Mock simple
    class CredentialUnavailableError(Exception): pass # Mock simple

# Importar helper HTTP y constantes (aunque BASE_URL no se use aquí)
try:
    # Asume que shared está un nivel arriba de actions
    from ..shared.helpers.http_client import hacer_llamada_api
    from ..shared.constants import BASE_URL, GRAPH_API_TIMEOUT # GRAPH_API_TIMEOUT se usa como base
except ImportError as e:
    logging.critical(f"Error CRÍTICO importando helpers/constantes en Power Automate: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs):
        raise NotImplementedError("Dependencia 'hacer_llamada_api' no importada correctamente.")

# Usar el logger estándar de Azure Functions
logger = logging.getLogger("azure.functions")

# --- Constantes y Variables de Entorno Específicas para Azure Management ---
# Estas son necesarias para la autenticación con ClientSecretCredential
# Deben estar configuradas en el entorno de la Azure Function App
try:
    AZURE_SUBSCRIPTION_ID = os.environ['AZURE_SUBSCRIPTION_ID']
    AZURE_RESOURCE_GROUP = os.environ['AZURE_RESOURCE_GROUP']
    # AZURE_LOCATION es opcional si se pasa en 'parametros' al crear
    AZURE_LOCATION = os.environ.get('AZURE_LOCATION')
    # Credenciales de la App Registration con permisos sobre los workflows/Logic Apps
    AZURE_CLIENT_ID = os.environ['AZURE_CLIENT_ID_MGMT'] # Usar nombre específico si es diferente
    AZURE_TENANT_ID = os.environ['AZURE_TENANT_ID'] # Usualmente el mismo tenant
    AZURE_CLIENT_SECRET = os.environ['AZURE_CLIENT_SECRET_MGMT'] # Usar nombre específico
except KeyError as e:
    logger.critical(f"Error Crítico: Falta variable de entorno esencial para Power Automate Management: {e}")
    # Lanzar error para detener la carga del módulo si falta configuración esencial
    raise ValueError(f"Configuración incompleta para Power Automate Management: falta {e}")

# Endpoints y configuración para Azure Resource Manager (ARM) API
AZURE_MGMT_BASE_URL = "https://management.azure.com"
AZURE_MGMT_SCOPE = "https://management.azure.com/.default" # Scope para obtener token ARM
LOGIC_API_VERSION = "2019-05-01" # API version para workflows (Logic Apps)
# Timeout específico para llamadas a Management API (pueden ser más lentas)
AZURE_MGMT_TIMEOUT = max(GRAPH_API_TIMEOUT, 60) # Ej: 60 segundos mínimo

# --- Helper de Autenticación (Específico para este módulo) ---
# Cache simple para la credencial y el token ARM
_credential_pa: Optional[ClientSecretCredential] = None
_cached_mgmt_token_pa: Optional[str] = None # TODO: Añadir manejo de expiración de token

def _get_azure_mgmt_token() -> str:
    """Obtiene un token de acceso para Azure Management API usando Client Credentials."""
    global _credential_pa, _cached_mgmt_token_pa

    # TODO: Implementar chequeo de expiración del token cacheado si es necesario
    if _cached_mgmt_token_pa:
        # logger.debug("Usando token ARM cacheado.")
        return _cached_mgmt_token_pa

    # Crear credencial si no existe
    if not _credential_pa:
        logger.info("Creando credencial ClientSecretCredential para Azure Management (PA).")
        try:
            _credential_pa = ClientSecretCredential(
                tenant_id=AZURE_TENANT_ID,
                client_id=AZURE_CLIENT_ID,
                client_secret=AZURE_CLIENT_SECRET
            )
        except Exception as cred_err:
             logger.critical(f"Error al crear ClientSecretCredential: {cred_err}", exc_info=True)
             raise Exception(f"Error configurando credencial Azure (PA): {cred_err}") from cred_err

    # Obtener token
    try:
        logger.info(f"Solicitando token para Azure Management con scope: {AZURE_MGMT_SCOPE}")
        # Asegurar que _credential_pa no sea None (para mypy)
        if _credential_pa is None: raise Exception("Credencial no inicializada.")

        token_info = _credential_pa.get_token(AZURE_MGMT_SCOPE)
        _cached_mgmt_token_pa = token_info.token
        # Podríamos cachear token_info.expires_on para invalidar el caché
        logger.info("Token para Azure Management (PA) obtenido exitosamente.")
        return _cached_mgmt_token_pa
    except CredentialUnavailableError as cred_err:
         logger.critical(f"Credencial no disponible para obtener token ARM: {cred_err}", exc_info=True)
         raise Exception(f"Credencial Azure (PA) no disponible: {cred_err}") from cred_err
    except Exception as e:
        logger.error(f"Error inesperado obteniendo token de Azure Management (PA): {e}", exc_info=True)
        raise Exception(f"Error obteniendo token Azure (PA): {e}") from e

def _get_auth_headers_for_mgmt() -> Dict[str, str]:
    """Construye el diccionario de cabeceras necesario para llamadas a ARM API."""
    try:
        token = _get_azure_mgmt_token()
        return {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    except Exception as e:
        # Propagar error si no se pudo obtener el token
        raise Exception(f"No se pudieron obtener las cabeceras de autenticación para Management API: {e}") from e

# ========================================================
# ==== FUNCIONES DE ACCIÓN PARA POWER AUTOMATE (FLOWS) ====
# ========================================================
# Usan la firma (parametros: Dict[str, Any], headers: Dict[str, str])
# PERO usan la autenticación interna (_get_auth_headers_for_mgmt) para las llamadas ARM.
# Los 'headers' de entrada (Graph API) no se usan directamente aquí.

def listar_flows(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista los flujos (workflows de Logic Apps) en una suscripción y grupo de recursos.

    Args:
        parametros (Dict[str, Any]): Opcional: 'suscripcion_id', 'grupo_recurso'.
                                     Si no se proporcionan, se usan los valores de entorno.
        headers (Dict[str, str]): Ignorados (se usa autenticación interna ARM).

    Returns:
        Dict[str, Any]: Respuesta de ARM API, usualmente {'value': [...]}.
    """
    # Obtener headers de autenticación específicos para ARM
    auth_headers = _get_auth_headers_for_mgmt()

    # Usar valores de entorno como default
    sid = parametros.get('suscripcion_id', AZURE_SUBSCRIPTION_ID)
    rg = parametros.get('grupo_recurso', AZURE_RESOURCE_GROUP)

    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows?api-version={LOGIC_API_VERSION}"
    logger.info(f"Listando flows (workflows) en Suscripción '{sid}', Grupo '{rg}'")

    # Usar el helper HTTP con los headers ARM
    return hacer_llamada_api("GET", url, auth_headers, timeout=AZURE_MGMT_TIMEOUT)


def obtener_flow(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene los detalles de un flujo (workflow) específico por su nombre.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_flow'.
                                     Opcional: 'suscripcion_id', 'grupo_recurso'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: El objeto del workflow de ARM API.
    """
    nombre_flow: Optional[str] = parametros.get("nombre_flow")
    if not nombre_flow: raise ValueError("Parámetro 'nombre_flow' es requerido.")

    auth_headers = _get_auth_headers_for_mgmt()
    sid = parametros.get('suscripcion_id', AZURE_SUBSCRIPTION_ID)
    rg = parametros.get('grupo_recurso', AZURE_RESOURCE_GROUP)

    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    logger.info(f"Obteniendo flow '{nombre_flow}' en Grupo '{rg}'")
    return hacer_llamada_api("GET", url, auth_headers, timeout=AZURE_MGMT_TIMEOUT)


def crear_flow(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea un nuevo flujo (workflow) a partir de su definición.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_flow', 'definicion_flow' (dict).
                                     Opcional: 'ubicacion', 'suscripcion_id', 'grupo_recurso'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: El objeto del workflow creado.
    """
    nombre_flow: Optional[str] = parametros.get("nombre_flow")
    definicion_flow: Optional[Dict[str, Any]] = parametros.get("definicion_flow")
    ubicacion: Optional[str] = parametros.get("ubicacion", AZURE_LOCATION) # Usar env var como fallback

    if not nombre_flow: raise ValueError("Parámetro 'nombre_flow' es requerido.")
    if not definicion_flow or not isinstance(definicion_flow, dict):
        raise ValueError("Parámetro 'definicion_flow' (diccionario) es requerido.")
    if not ubicacion:
        # Si no se proporcionó ni en params ni en env var
        raise ValueError("Se requiere 'ubicacion' (o variable de entorno AZURE_LOCATION) para crear el flow.")

    auth_headers = _get_auth_headers_for_mgmt()
    sid = parametros.get('suscripcion_id', AZURE_SUBSCRIPTION_ID)
    rg = parametros.get('grupo_recurso', AZURE_RESOURCE_GROUP)

    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    # Estructura del cuerpo para crear/actualizar workflow
    body: Dict[str, Any] = {
        "location": ubicacion,
        "properties": {
            "definition": definicion_flow
            # Se podrían añadir otros parámetros como 'state', 'parameters', etc.
        }
    }
    logger.info(f"Creando flow '{nombre_flow}' en Grupo '{rg}', Ubicación '{ubicacion}'")
    # Usar PUT para crear (idempotente)
    return hacer_llamada_api("PUT", url, auth_headers, json_data=body, timeout=AZURE_MGMT_TIMEOUT * 2)


def actualizar_flow(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Actualiza la definición de un flujo (workflow) existente.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_flow', 'definicion_flow' (dict).
                                     Opcional: 'suscripcion_id', 'grupo_recurso'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: El objeto del workflow actualizado.
    """
    nombre_flow: Optional[str] = parametros.get("nombre_flow")
    definicion_flow: Optional[Dict[str, Any]] = parametros.get("definicion_flow")

    if not nombre_flow: raise ValueError("Parámetro 'nombre_flow' es requerido.")
    if not definicion_flow or not isinstance(definicion_flow, dict):
        raise ValueError("Parámetro 'definicion_flow' (diccionario) es requerido.")

    auth_headers = _get_auth_headers_for_mgmt()
    sid = parametros.get('suscripcion_id', AZURE_SUBSCRIPTION_ID)
    rg = parametros.get('grupo_recurso', AZURE_RESOURCE_GROUP)

    # Para actualizar, usualmente se necesita obtener el objeto actual y luego modificarlo.
    # Sin embargo, PUT en ARM es idempotente, podemos enviar la nueva definición completa.
    # Necesitamos la ubicación del flujo existente.
    try:
        params_get = {"nombre_flow": nombre_flow, "suscripcion_id": sid, "grupo_recurso": rg}
        # Llamada interna para obtener el flujo actual (usa la misma autenticación)
        # No pasar 'headers' de entrada aquí, la función interna usa la auth ARM.
        current_flow = obtener_flow(params_get, {})
        current_location = current_flow.get("location")
        if not current_location: raise ValueError("No se pudo obtener la ubicación del flow existente.")
    except Exception as get_err:
        raise Exception(f"No se pudo obtener el flow actual '{nombre_flow}' para actualizar: {get_err}") from get_err

    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    # Cuerpo similar a crear, pero usando la ubicación existente
    body: Dict[str, Any] = {
        "location": current_location,
        "properties": {
            "definition": definicion_flow
        }
    }
    logger.info(f"Actualizando flow '{nombre_flow}' en Grupo '{rg}'")
    # Usar PUT para actualizar
    return hacer_llamada_api("PUT", url, auth_headers, json_data=body, timeout=AZURE_MGMT_TIMEOUT * 2)


def eliminar_flow(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina un flujo (workflow).

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_flow'.
                                     Opcional: 'suscripcion_id', 'grupo_recurso'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: Confirmación de eliminación.
    """
    nombre_flow: Optional[str] = parametros.get("nombre_flow")
    if not nombre_flow: raise ValueError("Parámetro 'nombre_flow' es requerido.")

    auth_headers = _get_auth_headers_for_mgmt()
    sid = parametros.get('suscripcion_id', AZURE_SUBSCRIPTION_ID)
    rg = parametros.get('grupo_recurso', AZURE_RESOURCE_GROUP)

    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={LOGIC_API_VERSION}"
    logger.info(f"Eliminando flow '{nombre_flow}' de Grupo '{rg}'")

    # DELETE devuelve 204 o 200 (None o {} del helper).
    hacer_llamada_api("DELETE", url, auth_headers, timeout=AZURE_MGMT_TIMEOUT)
    return {"status": "Eliminado", "flow": nombre_flow}


def ejecutar_flow(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Ejecuta un flujo llamando a su URL de trigger HTTP.
    La autenticación depende de cómo esté configurado el trigger del flujo (ej. SAS key en URL, AAD, etc.).
    Esta función NO usa la autenticación ARM.

    Args:
        parametros (Dict[str, Any]): Debe contener 'flow_url' (URL completa del trigger).
                                     Opcional: 'payload' (dict, cuerpo a enviar al flujo).
        headers (Dict[str, str]): Cabeceras originales (podrían contener auth si el trigger la requiere).

    Returns:
        Dict[str, Any]: Estado y respuesta del trigger del flujo.
    """
    flow_url: Optional[str] = parametros.get("flow_url")
    payload: Optional[Dict[str, Any]] = parametros.get("payload") # Cuerpo para el trigger

    if not flow_url: raise ValueError("Parámetro 'flow_url' (URL del trigger) es requerido.")

    # Usar las cabeceras originales PASADAS a esta función, ya que pueden
    # contener información de autenticación necesaria para el trigger del flujo.
    # Añadir Content-Type si enviamos payload.
    request_headers = headers.copy() # Usar headers de entrada
    if payload:
        request_headers['Content-Type'] = 'application/json'

    logger.info(f"Ejecutando trigger de flow: POST {flow_url}")

    # Usar requests directamente aquí porque:
    # 1. La URL es externa y no sigue necesariamente el patrón BASE_URL.
    # 2. La autenticación es variable (puede estar en la URL, headers, o ninguna).
    # 3. La respuesta puede no ser JSON o seguir el patrón del helper.
    try:
        # Timeout más corto para triggers? O usar AZURE_MGMT_TIMEOUT? Usamos MGMT por si el flujo es largo.
        response = requests.post(
            flow_url,
            headers=request_headers, # Usar headers de entrada
            json=payload if payload else None, # Enviar payload como JSON si existe
            timeout=AZURE_MGMT_TIMEOUT
        )
        # raise_for_status lanzará error si el trigger devuelve 4xx/5xx
        response.raise_for_status()
        logger.info(f"Trigger de flujo en '{flow_url}' ejecutado. Status: {response.status_code}")

        # Intentar obtener respuesta JSON, si no, texto.
        try:
            resp_data = response.json()
        except json.JSONDecodeError:
            resp_data = response.text # Devolver texto si no es JSON

        return {
            "status": "Ejecutado" if response.ok else "Fallido", # ok cubre 2xx
            "status_code": response.status_code,
            "response_body": resp_data
        }
    except requests.exceptions.RequestException as e:
        error_body = e.response.text[:200] if e.response else "N/A"
        logger.error(f"Error Request al ejecutar trigger de flow '{flow_url}': {e}. Respuesta: {error_body}", exc_info=True)
        raise Exception(f"Error API ejecutando trigger de flow: {e}") from e
    except Exception as e:
        logger.error(f"Error inesperado al ejecutar trigger de flow '{flow_url}': {e}", exc_info=True)
        raise


def obtener_estado_ejecucion_flow(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene el estado y detalles de una ejecución (run) específica de un flujo.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_flow', 'run_id'.
                                     Opcional: 'suscripcion_id', 'grupo_recurso'.
        headers (Dict[str, str]): Ignorados.

    Returns:
        Dict[str, Any]: El objeto 'run' de ARM API con el estado de la ejecución.
    """
    nombre_flow: Optional[str] = parametros.get("nombre_flow")
    run_id: Optional[str] = parametros.get("run_id") # ID de la ejecución

    if not nombre_flow: raise ValueError("Parámetro 'nombre_flow' es requerido.")
    if not run_id: raise ValueError("Parámetro 'run_id' es requerido.")

    auth_headers = _get_auth_headers_for_mgmt()
    sid = parametros.get('suscripcion_id', AZURE_SUBSCRIPTION_ID)
    rg = parametros.get('grupo_recurso', AZURE_RESOURCE_GROUP)

    url = f"{AZURE_MGMT_BASE_URL}/subscriptions/{sid}/resourceGroups/{rg}/providers/Microsoft.Logic/workflows/{nombre_flow}/runs/{run_id}?api-version={LOGIC_API_VERSION}"
    logger.info(f"Obteniendo estado de ejecución '{run_id}' del flow '{nombre_flow}'")
    return hacer_llamada_api("GET", url, auth_headers, timeout=AZURE_MGMT_TIMEOUT)

# --- FIN DEL MÓDULO actions/power_automate.py ---
