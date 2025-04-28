import json
import logging
import requests
import azure.functions as func # Asegúrate que azure.functions está importado
# CORRECCION: Añadir tipos necesarios
from typing import Dict, Any, Callable, List, Optional, Union, Mapping, Sequence
# CORRECCION: datetime ya estaba, timezone también
from datetime import datetime, timezone
import os

# Configuración de logging
logger = logging.getLogger("azure.functions")
logger.setLevel(logging.INFO)

# --- INICIO: Variables de Entorno y Configuración ---
def get_config_or_raise(key: str, default: Optional[str] = None) -> str:
    value = os.environ.get(key, default)
    if value is None:
        logger.error(f"Falta la variable de entorno requerida: {key}")
        raise ValueError(f"Falta la variable de entorno: {key}")
    return value

try:
    CLIENT_ID = get_config_or_raise('CLIENT_ID')
    TENANT_ID = get_config_or_raise('TENANT_ID')
    CLIENT_SECRET = get_config_or_raise('CLIENT_SECRET')
    MAILBOX = get_config_or_raise('MAILBOX', default='me')
    GRAPH_SCOPE = os.environ.get('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')
    logger.info("Variables de entorno cargadas correctamente.")
except ValueError as e:
    logger.critical(f"Error CRÍTICO de configuración inicial: {e}. La función no puede operar.")
    raise

# --- FIN: Variables de Entorno y Configuración ---


# --- INICIO: Constantes y Autenticación ---
BASE_URL = "https://graph.microsoft.com/v1.0"
# CORRECCION: Tipado más explícito para HEADERS globales
HEADERS: Dict[str, Optional[str]] = {
    'Authorization': None,
    'Content-Type': 'application/json'
}

def obtener_token() -> str:
    logger.info("Obteniendo token de acceso...")
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        'client_id': CLIENT_ID, 'scope': GRAPH_SCOPE,
        'client_secret': CLIENT_SECRET, 'grant_type': 'client_credentials'
    }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    response = None
    try:
        response = requests.post(url, data=data, headers=headers)
        response.raise_for_status()
        token_data = response.json()
        token = token_data.get('access_token')
        if not token:
            logger.error(f"❌ No se encontró 'access_token' en la respuesta. Respuesta: {token_data}")
            raise Exception("No se pudo obtener el token de acceso de la respuesta.")
        # logger.info(f"Token obtenido correctamente.") # Log un poco menos verboso
        return token
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error de red/HTTP al obtener el token: {e}. Detalles: {error_details}")
        raise Exception(f"Error de red/HTTP al obtener el token: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"❌ Error al decodificar la respuesta JSON del token: {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar la respuesta JSON del token: {e}")
    except Exception as e:
        logger.error(f"❌ Error inesperado al obtener el token: {e}")
        raise

def _actualizar_headers() -> None:
    try:
        token = obtener_token() # Llama a la función global definida en auth.py si auth.py está separado, o la local si está aquí.
        # Si auth.py está separado, asegúrate de importarla: from auth import obtener_token
        HEADERS['Authorization'] = f'Bearer {token}'
        logger.info("Cabecera de autorización actualizada.")
    except Exception as e:
        logger.error(f"❌ Falló la actualización de la cabecera de autorización: {e}")
        # No relanzar aquí necesariamente, depende de si la función que llama puede reintentar
        # Por ahora, relanzamos para mantener comportamiento anterior
        raise Exception(f"Fallo al actualizar la cabecera: {e}")

# --- FIN: Constantes y Autenticación ---


# --- INICIO: Funciones Auxiliares de Graph API ---

# ---- CORREO ----
def listar_correos(
    top: int = 10, skip: int = 0, folder: str = 'Inbox',
    select: Optional[List[str]] = None,
    filter_query: Optional[str] = None, # Espera str o None
    order_by: Optional[str] = None, # Espera str o None
    mailbox: Optional[str] = None
) -> Dict[str, Any]: # CORRECCION: Mejor tipo de retorno
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/mailFolders/{folder}/messages"
    # CORRECCION: Tipado explícito y asegurar int (aunque la conversión principal ocurre en main)
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by

    response: Optional[requests.Response] = None # Inicializar para except block
    try:
        clean_params = {k:v for k, v in params.items() if v is not None}
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}")
        response = requests.get(url, headers=HEADERS, params=clean_params)
        response.raise_for_status()
        data: Dict[str, Any] = response.json() # CORRECCION: Tipado explícito
        logger.info(f"Listados {len(data.get('value',[]))} correos.")
        return data
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error listar correos: {e}. URL: {url}. Detalles: {error_details}")
        raise Exception(f"Error al listar correos: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"❌ Error JSON (listar correos): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (listar correos): {e}")

# ... (Definiciones de leer_correo, enviar_correo, guardar_borrador, etc.) ...
# (Asegúrate de que estas funciones también estén aquí si este es tu único archivo)
# ... (Código de funciones de correo omitido por brevedad, asumiendo que son las mismas que antes)...

# ---- CALENDARIO ----
def listar_eventos(
    top: int = 10, start_date: Optional[datetime] = None, end_date: Optional[datetime] = None,
    filter_query: Optional[str] = None, order_by: Optional[str] = None,
    select: Optional[List[str]] = None, # CORRECCION: MyPy se quejaba de Collection[str], aseguramos List[str]
    use_calendar_view: bool = True, mailbox: Optional[str] = None
) -> Dict[str, Any]: # CORRECCION: Mejor tipo de retorno
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    base_endpoint = f"{BASE_URL}/users/{usuario}"
    params: Dict[str, Any] = {}
    endpoint_suffix = ""
    # Lógica para determinar endpoint y parámetros (CalendarView o Events)
    if use_calendar_view and start_date and end_date:
        endpoint_suffix = "/calendarView"
        if isinstance(start_date, datetime) and start_date.tzinfo is None: start_date = start_date.replace(tzinfo=timezone.utc)
        if isinstance(end_date, datetime) and end_date.tzinfo is None: end_date = end_date.replace(tzinfo=timezone.utc)
        # CORRECCION: Validar que sean datetime antes de llamar a isoformat
        if isinstance(start_date, datetime): params['startDateTime'] = start_date.isoformat()
        if isinstance(end_date, datetime): params['endDateTime'] = end_date.isoformat()
        params['$top'] = int(top)
        if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
        if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
        if select and isinstance(select, list): params['$select'] = ','.join(select) # Check type
    else:
        endpoint_suffix = "/events"
        params['$top'] = int(top)
        filters = []
        if start_date and isinstance(start_date, datetime): # Check type
             if start_date.tzinfo is None: start_date = start_date.replace(tzinfo=timezone.utc)
             filters.append(f"start/dateTime ge '{start_date.isoformat()}'")
        if end_date and isinstance(end_date, datetime): # Check type
             if end_date.tzinfo is None: end_date = end_date.replace(tzinfo=timezone.utc)
             filters.append(f"end/dateTime le '{end_date.isoformat()}'")
        if filter_query is not None and isinstance(filter_query, str): filters.append(f"({filter_query})")

        if filters: params['$filter'] = " and ".join(filters)
        if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
        if select and isinstance(select, list): params['$select'] = ','.join(select) # Check type

    url = f"{base_endpoint}{endpoint_suffix}"
    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}")
        response = requests.get(url, headers=HEADERS, params=clean_params)
        response.raise_for_status()
        data: Dict[str, Any] = response.json() # CORRECCION: Tipado explícito
        logger.info(f"Listados eventos.")
        return data
    # ... (resto de los except como estaban) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error listar eventos: {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}"); raise Exception(f"Error listar eventos: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (listar eventos): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar eventos): {e}")


def crear_evento(
    titulo: str, inicio: datetime, fin: datetime,
    asistentes: Optional[List[Dict[str, Any]]] = None, # MyPy espera List[Dict] aquí
    cuerpo: Optional[str] = None, es_reunion_online: bool = False,
    proveedor_reunion_online: str = "teamsForBusiness",
    recordatorio_minutos: Optional[int] = 15, ubicacion: Optional[str] = None,
    mostrar_como: str = "busy", mailbox: Optional[str] = None
) -> Dict[str, Any]: # CORRECCION: Mejor tipo de retorno
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/events"

    # CORRECCION: Asegurar que inicio y fin son datetime
    if not isinstance(inicio, datetime) or not isinstance(fin, datetime):
        # Esto debería ser manejado idealmente en la función 'main' al parsear params
        raise ValueError("'inicio' y 'fin' deben ser objetos datetime.")
    if inicio.tzinfo is None: inicio = inicio.replace(tzinfo=timezone.utc)
    if fin.tzinfo is None: fin = fin.replace(tzinfo=timezone.utc)

    # CORRECCION: Tipado explícito para 'body' para ayudar a MyPy
    body: Dict[str, Any] = {
        "subject": titulo,
        "start": {"dateTime": inicio.isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": fin.isoformat(), "timeZone": "UTC"},
        "showAs": mostrar_como
    }
    # CORRECCION: Validar tipo de 'asistentes' y estructura interna si es posible
    # El error [misc] en la línea 425 indicaba problema aquí. Asegurémonos que 'asistentes' es una lista de dicts.
    if asistentes is not None:
        if isinstance(asistentes, list):
             # Verificar que los elementos sean diccionarios (simplificado)
             if all(isinstance(a, dict) for a in asistentes):
                  body["attendees"] = [
                      {"emailAddress": {"address": a.get('emailAddress')}, # Asumir que viene como dict
                       "type": a.get('type', 'required')}
                      for a in asistentes if a and a.get('emailAddress') # Filtrar vacíos/inválidos
                  ]
             else:
                  logger.warning(f"Elementos en 'asistentes' no son todos diccionarios.")
                  # Decidir si fallar o continuar sin asistentes
        else:
             logger.warning(f"Tipo inesperado para 'asistentes': {type(asistentes)}")

    # CORRECCION: Validar tipos antes de asignar a body (Errores [assignment] en 428, 429, 430)
    if cuerpo is not None and isinstance(cuerpo, str): body["body"] = {"contentType": "HTML", "content": cuerpo}
    if ubicacion is not None and isinstance(ubicacion, str): body["location"] = {"displayName": ubicacion}
    # Estas asignaciones parecen correctas si body es Dict[str, Any], los errores de MyPy podrían ser por inferencia incorrecta
    if es_reunion_online: body["isOnlineMeeting"] = True
    if proveedor_reunion_online: body["onlineMeetingProvider"] = proveedor_reunion_online
    if recordatorio_minutos is not None and isinstance(recordatorio_minutos, int):
        body["isReminderOn"] = True
        body["reminderMinutesBeforeStart"] = recordatorio_minutos
    else:
        body["isReminderOn"] = False

    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: POST {url}")
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data: Dict[str, Any] = response.json()
        logger.info(f"Evento '{titulo}' creado con ID: {data.get('id')}.")
        return data
    # ... (resto de los except como estaban) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error crear evento: {e}. Detalles: {error_details}. URL: {url}"); raise Exception(f"Error crear evento: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (crear evento): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (crear evento): {e}")


# ... (actualizar_evento, eliminar_evento) ...
def actualizar_evento(evento_id: str, nuevos_valores: dict, mailbox: Optional[str] = None) -> Dict[str, Any]: # Mejor tipo retorno
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
    # Conversión de fechas (parece ok)
    if 'start' in nuevos_valores and isinstance(nuevos_valores['start'], datetime): start_dt = nuevos_valores['start']; if start_dt.tzinfo is None: start_dt = start_dt.replace(tzinfo=timezone.utc); nuevos_valores['start'] = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}
    if 'end' in nuevos_valores and isinstance(nuevos_valores['end'], datetime): end_dt = nuevos_valores['end']; if end_dt.tzinfo is None: end_dt = end_dt.replace(tzinfo=timezone.utc); nuevos_valores['end'] = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: PATCH {url}"); response = requests.patch(url, headers=HEADERS, json=nuevos_valores); response.raise_for_status(); data = response.json(); logger.info(f"Evento '{evento_id}' actualizado."); return data
    # ... (resto de los except como estaban) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error actualizar evento: {e}. Detalles: {error_details}. URL: {url}"); raise Exception(f"Error actualizar evento: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (actualizar evento): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (actualizar evento): {e}")

def eliminar_evento(evento_id: str, mailbox: Optional[str] = None) -> Dict[str, Any]: # Mejor tipo retorno
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: DELETE {url}"); response = requests.delete(url, headers=HEADERS); response.raise_for_status(); logger.info(f"Evento '{evento_id}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error eliminar evento: {e}. URL: {url}. Detalles: {error_details}"); raise Exception(f"Error eliminar evento: {e}")

# ---- TEAMS y OTROS ----
def listar_chats(top: int = 20, skip: int = 0, filter_query: Optional[str] = None, order_by: Optional[str] = None, expand: Optional[str] = None) -> Dict[str, Any]: # Mejor tipo retorno
    _actualizar_headers(); url = f"{BASE_URL}/me/chats"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
    if expand is not None and isinstance(expand, str): params['$expand'] = expand
    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params); response.raise_for_status(); data = response.json(); logger.info(f"Listados {len(data.get('value',[]))} chats."); return data
    # ... (resto de los except como estaban) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error listar chats: {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}"); raise Exception(f"Error listar chats: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (listar chats): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar chats): {e}")

# ... (listar_equipos, obtener_equipo) ...
def listar_equipos(top: int = 20, skip: int = 0, filter_query: Optional[str] = None) -> Dict[str, Any]: # Mejor tipo retorno
    _actualizar_headers(); url = f"{BASE_URL}/me/joinedTeams"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params); response.raise_for_status(); data = response.json(); logger.info(f"Listados {len(data.get('value',[]))} equipos."); return data
    # ... (resto de los except como estaban) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error listar equipos: {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}"); raise Exception(f"Error listar equipos: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (listar equipos): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar equipos): {e}")

def obtener_equipo(team_id: str, select: Optional[List[str]]=None) -> Dict[str, Any]: # Mejor tipo retorno
    _actualizar_headers(); url = f"{BASE_URL}/teams/{team_id}"
    params: Dict[str, Any] = {}
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params or None); response.raise_for_status(); data = response.json(); logger.info(f"Obtenido equipo ID: {team_id}."); return data
    # ... (resto de los except como estaban) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error obtener equipo: {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}"); raise Exception(f"Error obtener equipo: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (obtener equipo): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (obtener equipo): {e}")

# --- FIN: Funciones Auxiliares de Graph API ---


# --- INICIO: Función Principal de Azure Functions (Entry Point) ---

# Mapeo de nombres de acción a funciones
# CORRECCION: Tipado más específico
acciones_disponibles: Dict[str, Callable[..., Dict[str, Any]]] = {
    "listar_correos": listar_correos,
    "leer_correo": leer_correo,
    "enviar_correo": enviar_correo,
    "guardar_borrador": guardar_borrador,
    "enviar_borrador": enviar_borrador,
    "responder_correo": responder_correo,
    "reenviar_correo": reenviar_correo,
    "eliminar_correo": eliminar_correo,
    "listar_eventos": listar_eventos,
    "crear_evento": crear_evento,
    "actualizar_evento": actualizar_evento,
    "eliminar_evento": eliminar_evento,
    "listar_chats": listar_chats,
    "listar_equipos": listar_equipos,
    "obtener_equipo": obtener_equipo,
    # Añade aquí otras acciones que definas
}

def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info(f'Python HTTP trigger function procesando una solicitud. Method: {req.method}')

    accion: Optional[str] = None
    parametros: Dict[str, Any] = {}

    # --- Procesar solicitud ---
    try:
        # Intentar obtener JSON del cuerpo (solo para POST, PUT, PATCH, etc.)
        req_body: Optional[Dict[str, Any]] = None
        if req.method in ('POST', 'PUT', 'PATCH'):
            try:
                req_body = req.get_json()
                if not isinstance(req_body, dict):
                     logger.warning('Cuerpo JSON no es un diccionario.')
                     return func.HttpResponse("Cuerpo JSON debe ser un objeto.", status_code=400)
            except ValueError:
                logger.warning('No se pudo decodificar JSON del cuerpo.')
                return func.HttpResponse("Cuerpo JSON inválido.", status_code=400)
        else:
            # Para GET, etc., podríamos leer de query parameters si quisiéramos
            # Por ahora, asumimos que la acción/params vienen en el body para POST
            # o quizás en query params para GET (no implementado aquí)
             logger.info("Método no es POST/PUT/PATCH, no se buscará cuerpo JSON.")
             # Podrías querer leer accion/params de req.params aquí para GET

        # Obtener acción y parámetros (priorizar body si existe)
        if req_body:
            accion = req_body.get('accion')
            params_input = req_body.get('parametros')
        else:
            # Leer de query params como fallback (ejemplo)
            accion = req.params.get('accion')
            params_input = req.params # Query params son strings, necesitarían más conversión
            # Ajustar cómo se manejan los params si vienen de query
            if isinstance(params_input, Mapping):
                 # Convertir query params (que son strings) si es necesario
                 # Esto es simplista, necesitarías un parseo más robusto
                 parametros = dict(params_input)
                 logger.info(f"Acción/Params leídos de Query Parameters: {parametros}")
            else:
                 parametros = {}


        # --- CORRECCION: Asegurar que parametros sea un dict si viene del body ---
        if req_body and isinstance(params_input, dict):
             parametros = params_input
        elif req_body and params_input is not None:
             logger.warning(f"'parametros' en body no es un diccionario: {type(params_input)}")
             parametros = {} # Ignorar
        elif not req_body and not parametros: # Si no hubo body y no se leyeron de query
             parametros = {}

        # Validar acción
        if not accion or not isinstance(accion, str):
            logger.warning("Clave 'accion' faltante, no es string, o no encontrada.")
            return func.HttpResponse("Falta 'accion' (string) en cuerpo JSON o query params.", status_code=400)

        logger.info(f"Acción a ejecutar: '{accion}'. Parámetros iniciales: {parametros}")

        # --- Ejecutar Acción ---
        if accion in acciones_disponibles:
            funcion_a_ejecutar = acciones_disponibles[accion]
            logger.info(f"Preparando ejecución de: {funcion_a_ejecutar.__name__}")

            # --- CORRECCION: Validar/Convertir parámetros ANTES de llamar ---
            params_procesados: Dict[str, Any] = {}
            try:
                # Convertir 'top' y 'skip' a int si existen y son necesarios para la acción
                if accion in ["listar_correos", "listar_eventos", "listar_chats", "listar_equipos"]:
                    if 'top' in parametros: params_procesados['top'] = int(parametros['top'])
                    if 'skip' in parametros: params_procesados['skip'] = int(parametros['skip'])
                # Convertir fechas para crear_evento si existen y son strings
                # (¡Requiere manejo robusto de formatos y errores!)
                elif accion == "crear_evento":
                     if 'inicio' in parametros and isinstance(parametros['inicio'], str):
                         try: params_procesados['inicio'] = datetime.fromisoformat(parametros['inicio'].replace('Z', '+00:00'))
                         except ValueError: raise ValueError("Formato de fecha 'inicio' inválido.")
                     if 'fin' in parametros and isinstance(parametros['fin'], str):
                          try: params_procesados['fin'] = datetime.fromisoformat(parametros['fin'].replace('Z', '+00:00'))
                          except ValueError: raise ValueError("Formato de fecha 'fin' inválido.")
                     # Validar tipos requeridos después de posible conversión
                     if not isinstance(parametros.get('inicio'), datetime): raise ValueError("'inicio' debe ser datetime.")
                     if not isinstance(parametros.get('fin'), datetime): raise ValueError("'fin' debe ser datetime.")
                     if not isinstance(parametros.get('titulo'), str): raise ValueError("'titulo' debe ser string.")


                # Copiar el resto de parámetros que no necesitaron conversión específica
                for k, v in parametros.items():
                    if k not in params_procesados: # No sobreescribir los ya convertidos
                        params_procesados[k] = v

                # Validaciones de parámetros requeridos (ejemplos)
                if accion in ["leer_correo", "enviar_borrador", "responder_correo", "reenviar_correo", "eliminar_correo", "actualizar_evento", "eliminar_evento", "obtener_equipo"] and 'message_id' not in params_procesados and 'evento_id' not in params_procesados and 'team_id' not in params_procesados:
                    # Mejorar esta lógica para chequear el ID correcto según la acción
                    logger.warning(f"Falta ID requerido (message_id, evento_id, team_id) para acción '{accion}'.")
                    # raise ValueError(f"Falta ID requerido para acción '{accion}'.") # Opcional fallar aquí
                if accion == "enviar_correo" and not all(k in params_procesados for k in ["destinatario", "asunto", "mensaje"]):
                    raise ValueError("Faltan 'destinatario', 'asunto' o 'mensaje' para enviar_correo.")

            except (ValueError, TypeError) as conv_err:
                logger.error(f"Error en parámetros para '{accion}': {conv_err}. Recibido: {parametros}")
                return func.HttpResponse(f"Parámetros inválidos para '{accion}': {conv_err}", status_code=400)

            # Llamar a la función
            logger.info(f"Ejecutando {funcion_a_ejecutar.__name__} con params: {params_procesados}")
            # CORRECCION: Envolver llamada en try-except para capturar errores de la lógica interna
            try:
                resultado = funcion_a_ejecutar(**params_procesados)
                logger.info(f"Ejecución de '{accion}' completada.")
            except Exception as exec_err:
                # Capturar errores específicos de las funciones auxiliares
                logger.exception(f"Error durante la ejecución de la acción '{accion}': {exec_err}")
                return func.HttpResponse(f"Error interno al ejecutar la acción '{accion}'.", status_code=500)


            # Devolver resultado
            try:
                 # CORRECCION: Usar default=str para manejar objetos no serializables como datetime
                 return func.HttpResponse(json.dumps(resultado, default=str), mimetype="application/json")
            except TypeError as serialize_err:
                 logger.error(f"Error al serializar resultado JSON para '{accion}': {serialize_err}. Resultado: {resultado}")
                 return func.HttpResponse(f"Error interno: Respuesta no serializable para {accion}.", status_code=500)

        else:
            logger.warning(f"Acción '{accion}' no reconocida.")
            acciones_validas = list(acciones_disponibles.keys())
            return func.HttpResponse(f"Acción '{accion}' no reconocida. Válidas: {acciones_validas}", status_code=400)

    except Exception as e:
        # Captura general para errores inesperados en el flujo de 'main'
        logger.exception(f"Error GENERAL inesperado procesando solicitud: {e}")
        return func.HttpResponse("Error interno del servidor.", status_code=500)

# --- FIN: Función Principal ---
