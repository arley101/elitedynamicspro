import json
import logging
import requests
import azure.functions as func
# Tipos necesarios
from typing import Dict, Any, Callable, List, Optional, Union, Mapping, Sequence
from datetime import datetime, timezone
import os

# Configuración de logging (Usando el logger de Azure Functions)
logger = logging.getLogger("azure.functions")
logger.setLevel(logging.INFO) # O logging.DEBUG si necesitas más detalle aún

# --- INICIO: Variables de Entorno y Configuración ---
def get_config_or_raise(key: str, default: Optional[str] = None) -> str:
    # Intenta obtener la variable de entorno. Lanza ValueError si falta y no hay default.
    value = os.environ.get(key, default)
    if value is None:
        logger.error(f"CONFIG ERROR: Falta la variable de entorno requerida: {key}")
        raise ValueError(f"Configuración esencial faltante: {key}")
    return value

try:
    # Cargar configuración esencial al inicio. Si falla, la función no puede operar.
    CLIENT_ID = get_config_or_raise('CLIENT_ID')
    TENANT_ID = get_config_or_raise('TENANT_ID')
    CLIENT_SECRET = get_config_or_raise('CLIENT_SECRET')
    MAILBOX = get_config_or_raise('MAILBOX', default='me') # Default 'me' si no se especifica
    GRAPH_SCOPE = os.environ.get('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')
    logger.info("Variables de entorno cargadas correctamente.")
except ValueError as e:
    # Log crítico y relanzar para que Azure sepa que la función no puede iniciar
    logger.critical(f"Error CRÍTICO de configuración inicial: {e}. La función no puede operar.")
    raise # Detiene la carga del módulo si falta configuración crítica

# --- FIN: Variables de Entorno y Configuración ---


# --- INICIO: Constantes y Autenticación ---
BASE_URL = "https://graph.microsoft.com/v1.0"
# Headers globales que se actualizarán con el token
HEADERS: Dict[str, Optional[str]] = {
    'Authorization': None,
    'Content-Type': 'application/json'
}
# Timeout general para llamadas a la API de Graph (en segundos)
GRAPH_API_TIMEOUT = 30

def obtener_token() -> str:
    """Obtiene un token de acceso de aplicación usando credenciales de cliente."""
    logger.info("Obteniendo token de acceso de aplicación...")
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {'client_id': CLIENT_ID, 'scope': GRAPH_SCOPE, 'client_secret': CLIENT_SECRET, 'grant_type': 'client_credentials'}
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    response = None
    try:
        response = requests.post(url, data=data, headers=headers, timeout=GRAPH_API_TIMEOUT) # Timeout añadido
        response.raise_for_status() # Lanza error para 4xx/5xx
        token_data = response.json()
        token = token_data.get('access_token')
        if not token:
            logger.error(f"No se encontró 'access_token' en la respuesta. Respuesta: {token_data}")
            raise Exception("No se pudo obtener el token de acceso de la respuesta.")
        # logger.info("Token obtenido correctamente.") # Opcional: menos verboso
        return token
    except requests.exceptions.Timeout:
        logger.error(f"Timeout al obtener token desde {url}")
        raise Exception("Timeout al contactar el servidor de autenticación.")
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"Error de red/HTTP al obtener token: {e}. Detalles: {error_details}")
        raise Exception(f"Error de red/HTTP al obtener token: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"Error al decodificar JSON del token: {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON del token: {e}")
    except Exception as e: # Captura otros errores inesperados
        logger.error(f"Error inesperado al obtener token: {e}")
        raise

def _actualizar_headers() -> None:
    """Obtiene un token fresco y actualiza los HEADERS globales."""
    try:
        token = obtener_token()
        HEADERS['Authorization'] = f'Bearer {token}'
        # logger.info("Cabecera de autorización actualizada.") # Opcional: menos verboso
    except Exception as e:
        # El error ya se loguea en obtener_token. Relanzamos para detener la operación.
        logger.error(f"Falló la actualización de la cabecera de autorización: {e}")
        raise Exception(f"Fallo al actualizar la cabecera: {e}")

# --- FIN: Constantes y Autenticación ---


# --- INICIO: Funciones Auxiliares de Graph API ---
# (Asegúrate de que TODAS tus funciones auxiliares: listar_correos, leer_correo, etc.
#  estén definidas aquí ANTES del diccionario 'acciones_disponibles')

# ---- CORREO ----
def listar_correos(top: int = 10, skip: int = 0, folder: str = 'Inbox', select: Optional[List[str]] = None, filter_query: Optional[str] = None, order_by: Optional[str] = None, mailbox: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/mailFolders/{folder}/messages"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
    response: Optional[requests.Response] = None
    try:
        clean_params = {k:v for k, v in params.items() if v is not None}; logger.info(f"API Call: GET {url} Params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data: Dict[str, Any] = response.json(); logger.info(f"Correos listados: {len(data.get('value',[]))}."); return data
    except requests.exceptions.Timeout: logger.error(f"Timeout listando correos: {url}"); raise Exception("Timeout en API Graph (listar correos).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (listar correos): {e}. Detalles: {error_details}"); raise Exception(f"Error API (listar correos): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response'); logger.error(f"Error JSON (listar correos): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar correos): {e}")

# ... (Pega aquí el resto de tus funciones auxiliares: leer_correo, enviar_correo, guardar_borrador,
#      enviar_borrador, responder_correo, reenviar_correo, eliminar_correo, listar_eventos,
#      crear_evento, actualizar_evento (la versión corregida), eliminar_evento, listar_chats,
#      listar_equipos, obtener_equipo, etc. ASEGÚRATE DE AÑADIR 'timeout=GRAPH_API_TIMEOUT'
#      a todas las llamadas requests.get/post/patch/delete/put) ...
# Ejemplo para leer_correo:
def leer_correo(message_id: str, select: Optional[List[str]] = None, mailbox: Optional[str] = None) -> dict:
     _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
     params = {}; response: Optional[requests.Response] = None
     if select and isinstance(select, list): params['$select'] = ','.join(select)
     try:
         logger.info(f"API Call: GET {url} Params: {params}"); response = requests.get(url, headers=HEADERS, params=params or None, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Correo '{message_id}' leído."); return data
     except requests.exceptions.Timeout: logger.error(f"Timeout leyendo correo: {url}"); raise Exception("Timeout en API Graph (leer correo).")
     except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (leer correo): {e}. URL: {url}. Detalles: {error_details}"); raise Exception(f"Error API (leer correo): {e}")
     except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response'); logger.error(f"Error JSON (leer correo): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (leer correo): {e}")

# !!!!! PEGA AQUÍ EL RESTO DE TUS FUNCIONES AUXILIARES (correo, calendario, teams, etc.), AÑADIENDO timeout=GRAPH_API_TIMEOUT a cada requests.xxx !!!!!

# --- FIN: Funciones Auxiliares de Graph API ---


# --- INICIO: Función Principal de Azure Functions (Entry Point) ---

# Mapeo de nombres de acción a las funciones DEFINIDAS ARRIBA
acciones_disponibles: Dict[str, Callable[..., Dict[str, Any]]] = {
    "listar_correos": listar_correos,
    "leer_correo": leer_correo,
    # "enviar_correo": enviar_correo, # Asegúrate que la función esté definida arriba
    # "guardar_borrador": guardar_borrador, # Asegúrate que la función esté definida arriba
    # ... (completa con TODAS las acciones/funciones que definiste arriba) ...
    # "listar_eventos": listar_eventos,
    # "crear_evento": crear_evento,
    # "actualizar_evento": actualizar_evento,
    # "eliminar_evento": eliminar_evento,
    # "listar_chats": listar_chats,
    # "listar_equipos": listar_equipos,
    # "obtener_equipo": obtener_equipo,
}
# Verificar que todas las funciones mapeadas existen
for accion, func_ref in acciones_disponibles.items():
    if not callable(func_ref):
        logger.error(f"Config Error: La función para la acción '{accion}' no es válida o no está definida.")
        # Podrías lanzar un error aquí para detener el inicio si hay un error de mapeo

def main(req: func.HttpRequest) -> func.HttpResponse:
    """Punto de entrada principal. Maneja la solicitud HTTP, llama a la acción apropiada y devuelve la respuesta."""
    logging.info(f'Python HTTP trigger function procesando solicitud. Method: {req.method}, URL: {req.url}')
    # Añadir ID de invocación para trazar logs
    invocation_id = req.headers.get('X-Azure-Functions-InvocationId', 'N/A')
    logging.info(f"Invocation ID: {invocation_id}")

    accion: Optional[str] = None
    parametros: Dict[str, Any] = {}
    funcion_a_ejecutar: Optional[Callable] = None

    # --- INICIO: Bloque Try-Except General ---
    # Sugerencia 1: Envuelve toda la lógica principal
    try:
        # --- Leer accion/parametros ---
        # Sugerencia 3: Manejo seguro de get_json ya estaba implementado
        req_body: Optional[Dict[str, Any]] = None
        if req.method in ('POST', 'PUT', 'PATCH'):
            try:
                req_body = req.get_json()
                if not isinstance(req_body, dict):
                     logger.warning(f'Invocation {invocation_id}: Cuerpo JSON no es un objeto.')
                     return func.HttpResponse("Cuerpo JSON debe ser un objeto.", status_code=400) # Sugerencia 2: Return garantizado
                accion = req_body.get('accion')
                params_input = req_body.get('parametros')
                if isinstance(params_input, dict): parametros = params_input
                elif params_input is not None: logger.warning(f"Invocation {invocation_id}: 'parametros' no es dict"); parametros = {}
                else: parametros = {}
            except ValueError:
                logger.warning(f'Invocation {invocation_id}: Cuerpo no es JSON válido.')
                return func.HttpResponse("Cuerpo JSON inválido.", status_code=400) # Sugerencia 2: Return garantizado
        else: # Para GET, etc. (simplificado)
            accion = req.params.get('accion')
            parametros = dict(req.params)

        # --- Validar acción ---
        if not accion or not isinstance(accion, str):
            logger.warning(f"Invocation {invocation_id}: Clave 'accion' faltante o no es string.")
            return func.HttpResponse("Falta 'accion' (string).", status_code=400) # Sugerencia 2: Return garantizado

        logger.info(f"Invocation {invocation_id}: Acción solicitada: '{accion}'. Parámetros iniciales: {parametros}")

        # --- Buscar y ejecutar la función ---
        if accion in acciones_disponibles:
            funcion_a_ejecutar = acciones_disponibles[accion]
            logger.info(f"Invocation {invocation_id}: Mapeado a función: {funcion_a_ejecutar.__name__}")

            # --- Validar/Convertir parámetros ---
            # (Esta sección sigue siendo importante y puede necesitar ajustes específicos por acción)
            params_procesados: Dict[str, Any] = {}
            try:
                params_procesados = parametros.copy()
                if accion in ["listar_correos", "listar_eventos", "listar_chats", "listar_equipos"]:
                    if 'top' in params_procesados: params_procesados['top'] = int(params_procesados['top'])
                    if 'skip' in params_procesados: params_procesados['skip'] = int(params_procesados['skip'])
                # ... (añadir más conversiones/validaciones aquí) ...

            except (ValueError, TypeError, KeyError) as conv_err:
                logger.error(f"Invocation {invocation_id}: Error en parámetros para '{accion}': {conv_err}. Recibido: {parametros}")
                return func.HttpResponse(f"Parámetros inválidos para '{accion}': {conv_err}", status_code=400) # Sugerencia 2: Return garantizado

            # --- Logging DEBUG antes de la llamada ---
            # (Mantenemos esto para diagnosticar el TypeError si persiste)
            logger.info(f"DEBUG Invocation {invocation_id}: Tipo de funcion_a_ejecutar: {type(funcion_a_ejecutar)}")
            logger.info(f"DEBUG Invocation {invocation_id}: Argumentos a pasar (params_procesados): {params_procesados}")
            logger.info(f"DEBUG Invocation {invocation_id}: Tipo de params_procesados: {type(params_procesados)}")

            # --- Llamar a la función auxiliar ---
            logger.info(f"Invocation {invocation_id}: Ejecutando {funcion_a_ejecutar.__name__}...")
            # El try-except interno captura errores específicos de la acción
            try:
                # Usamos la llamada genérica, ya que la específica para listar_correos no resolvió el TypeError
                resultado = funcion_a_ejecutar(**params_procesados)
                logger.info(f"Invocation {invocation_id}: Ejecución de '{accion}' completada.")
            except Exception as exec_err:
                # Loguea el error específico de la acción Y devuelve 500
                logger.exception(f"Invocation {invocation_id}: Error durante ejecución acción '{accion}': {exec_err}")
                return func.HttpResponse(f"Error interno al ejecutar '{accion}'.", status_code=500) # Sugerencia 2: Return garantizado

            # --- Devolver resultado exitoso ---
            try:
                 # Usar default=str para manejar datetimes, etc.
                 return func.HttpResponse(json.dumps(resultado, default=str), mimetype="application/json", status_code=200) # Sugerencia 2: Return garantizado
            except TypeError as serialize_err:
                 # Error si el resultado no se puede convertir a JSON
                 logger.error(f"Invocation {invocation_id}: Error al serializar resultado JSON para '{accion}': {serialize_err}.")
                 return func.HttpResponse(f"Error interno: Respuesta no serializable.", status_code=500) # Sugerencia 2: Return garantizado

        else: # Acción no encontrada
            logger.warning(f"Invocation {invocation_id}: Acción '{accion}' no reconocida."); acciones_validas = list(acciones_disponibles.keys());
            return func.HttpResponse(f"Acción '{accion}' no reconocida. Válidas: {acciones_validas}", status_code=400) # Sugerencia 2: Return garantizado

    # --- FIN: Bloque Try-Except General ---
    except Exception as e:
        # Captura CUALQUIER otro error inesperado en 'main'
        func_name = getattr(funcion_a_ejecutar, '__name__', 'N/A') if funcion_a_ejecutar else 'N/A'
        # Usar logger.exception para incluir stack trace en los logs de Azure
        logger.exception(f"Invocation {invocation_id}: Error GENERAL INESPERADO en main() procesando acción '{accion or 'desconocida'}' (Función: {func_name}): {e}")
        # Devolver siempre una respuesta HTTP, incluso en errores inesperados
        return func.HttpResponse("Error interno del servidor. Revise los logs.", status_code=500) # Sugerencia 1 y 2: Return garantizado

# --- FIN: Función Principal ---
