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
            logger.error(f"❌ No se encontró 'access_token'. Respuesta: {token_data}")
            raise Exception("No se pudo obtener el token de acceso.")
        return token
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error red/HTTP (token): {e}. Detalles: {error_details}"); raise Exception(f"Error red/HTTP (token): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (token): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (token): {e}")
    except Exception as e: logger.error(f"❌ Error inesperado (token): {e}"); raise

def _actualizar_headers() -> None:
    try:
        # Asumiendo que obtener_token() está definido globalmente aquí o importado
        # Si está en auth.py, necesitarías: from auth import obtener_token
        token = obtener_token()
        HEADERS['Authorization'] = f'Bearer {token}'
        logger.info("Cabecera de autorización actualizada.")
    except Exception as e:
        logger.error(f"❌ Falló la actualización de la cabecera: {e}")
        raise Exception(f"Fallo al actualizar la cabecera: {e}")

# --- FIN: Constantes y Autenticación ---


# --- INICIO: Funciones Auxiliares de Graph API ---

# ---- CORREO ----
def listar_correos(
    top: int = 10, skip: int = 0, folder: str = 'Inbox',
    select: Optional[List[str]] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None,
    mailbox: Optional[str] = None
) -> Dict[str, Any]:
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/mailFolders/{folder}/messages"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by

    response: Optional[requests.Response] = None
    try:
        clean_params = {k:v for k, v in params.items() if v is not None}
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}")
        response = requests.get(url, headers=HEADERS, params=clean_params)
        response.raise_for_status()
        data: Dict[str, Any] = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} correos.")
        return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error listar correos: {e}. URL: {url}. Detalles: {error_details}"); raise Exception(f"Error al listar correos: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (listar correos): {e}. Respuesta: {response_text}"); raise Exception(f"Error al decodificar JSON (listar correos): {e}")

# ... (Definiciones de leer_correo, enviar_correo, guardar_borrador, etc.) ...
# ... (Código omitido por brevedad, asegúrate que estén aquí)...
def leer_correo(message_id: str, select: Optional[List[str]] = None, mailbox: Optional[str] = None) -> dict:
     _actualizar_headers()
     usuario = mailbox or MAILBOX
     url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
     params = {}
     if select: params['$select'] = ','.join(select)
     response: Optional[requests.Response] = None
     try:
         logger.info(f"Llamando a Graph API: GET {url} con params: {params}")
         response = requests.get(url, headers=HEADERS, params=params or None)
         response.raise_for_status()
         data = response.json(); logger.info(f"Correo '{message_id}' leído."); return data
     except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error leer correo: {e}. URL: {url}. Detalles: {error_details}"); raise Exception(f"Error leer correo: {e}")
     except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (leer correo): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (leer correo): {e}")

def enviar_correo(destinatario: Union[str, List[str]], asunto: str, mensaje: str, cc: Optional[Union[str, List[str]]] = None, bcc: Optional[Union[str, List[str]]] = None, attachments: Optional[List[dict]] = None, is_draft: bool = False, mailbox: Optional[str] = None) -> dict:
     _actualizar_headers()
     usuario = mailbox or MAILBOX
     endpoint = "messages" if is_draft else "sendMail"
     url = f"{BASE_URL}/users/{usuario}/{endpoint}"
     if isinstance(destinatario, str): destinatario = [destinatario]
     if isinstance(cc, str): cc = [cc]
     if isinstance(bcc, str): bcc = [bcc]
     to_recipients = [{"emailAddress": {"address": r}} for r in destinatario if r]; cc_recipients = [{"emailAddress": {"address": r}} for r in cc if r] if cc else []; bcc_recipients = [{"emailAddress": {"address": r}} for r in bcc if r] if bcc else []
     if not to_recipients: raise ValueError("Se requiere destinatario.")
     message_payload = {"subject": asunto, "body": {"contentType": "HTML", "content": mensaje},"toRecipients": to_recipients,}
     if cc_recipients: message_payload["ccRecipients"] = cc_recipients
     if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
     if attachments: message_payload["attachments"] = attachments
     payload = {"message": message_payload, "saveToSentItems": "true"} if not is_draft else message_payload
     response: Optional[requests.Response] = None
     try:
         logger.info(f"Llamando a Graph API: POST {url}"); response = requests.post(url, headers=HEADERS, json=payload); response.raise_for_status()
         if not is_draft: logger.info(f"Correo enviado por '{usuario}'."); return {"status": "Enviado", "code": response.status_code}
         else: data = response.json(); message_id = data.get('id'); logger.info(f"Borrador guardado por '{usuario}' ID: {message_id}."); return {"status": "Borrador Guardado", "code": response.status_code, "id": message_id, "data": data}
     except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error enviar/guardar correo: {e}. Detalles: {error_details}. URL: {url}"); raise Exception(f"Error enviar/guardar correo: {e}")
     except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (guardar borrador): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (guardar borrador): {e}")

# ... (resto de funciones de correo) ...

# ---- CALENDARIO ----
def listar_eventos(
    top: int = 10, start_date: Optional[datetime] = None, end_date: Optional[datetime] = None,
    filter_query: Optional[str] = None, order_by: Optional[str] = None,
    select: Optional[List[str]] = None, use_calendar_view: bool = True, mailbox: Optional[str] = None
) -> Dict[str, Any]:
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    base_endpoint = f"{BASE_URL}/users/{usuario}"
    params: Dict[str, Any] = {}
    endpoint_suffix = ""
    if use_calendar_view and start_date and end_date:
        endpoint_suffix = "/calendarView"
        if isinstance(start_date, datetime) and start_date.tzinfo is None: start_date = start_date.replace(tzinfo=timezone.utc)
        if isinstance(end_date, datetime) and end_date.tzinfo is None: end_date = end_date.replace(tzinfo=timezone.utc)
        if isinstance(start_date, datetime): params['startDateTime'] = start_date.isoformat()
        if isinstance(end_date, datetime): params['endDateTime'] = end_date.isoformat()
        params['$top'] = int(top)
        if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
        if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
        if select and isinstance(select, list): params['$select'] = ','.join(select)
    else:
        endpoint_suffix = "/events"
        params['$top'] = int(top)
        filters = []
        if start_date and isinstance(start_date, datetime):
             if start_date.tzinfo is None: start_date = start_date.replace(tzinfo=timezone.utc)
             filters.append(f"start/dateTime ge '{start_date.isoformat()}'")
        if end_date and isinstance(end_date, datetime):
             if end_date.tzinfo is None: end_date = end_date.replace(tzinfo=timezone.utc)
             filters.append(f"end/dateTime le '{end_date.isoformat()}'")
        if filter_query is not None and isinstance(filter_query, str): filters.append(f"({filter_query})")
        if filters: params['$filter'] = " and ".join(filters)
        if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
        if select and isinstance(select, list): params['$select'] = ','.join(select)

    url = f"{base_endpoint}{endpoint_suffix}"
    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}")
        response = requests.get(url, headers=HEADERS, params=clean_params)
        response.raise_for_status(); data = response.json(); logger.info(f"Listados eventos."); return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error listar eventos: {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}"); raise Exception(f"Error listar eventos: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (listar eventos): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar eventos): {e}")


def crear_evento(
    titulo: str, inicio: datetime, fin: datetime,
    asistentes: Optional[List[Dict[str, Any]]] = None,
    cuerpo: Optional[str] = None, es_reunion_online: bool = False,
    proveedor_reunion_online: str = "teamsForBusiness",
    recordatorio_minutos: Optional[int] = 15, ubicacion: Optional[str] = None,
    mostrar_como: str = "busy", mailbox: Optional[str] = None
) -> Dict[str, Any]:
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/events"
    if not isinstance(inicio, datetime) or not isinstance(fin, datetime): raise ValueError("'inicio' y 'fin' deben ser datetimes.")
    if inicio.tzinfo is None: inicio = inicio.replace(tzinfo=timezone.utc)
    if fin.tzinfo is None: fin = fin.replace(tzinfo=timezone.utc)
    body: Dict[str, Any] = {
        "subject": titulo,
        "start": {"dateTime": inicio.isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": fin.isoformat(), "timeZone": "UTC"},
        "showAs": mostrar_como
    }
    if asistentes is not None:
        if isinstance(asistentes, list) and all(isinstance(a, dict) for a in asistentes):
             body["attendees"] = [{"emailAddress": {"address": a.get('emailAddress')},"type": a.get('type', 'required')} for a in asistentes if a and a.get('emailAddress')]
        else: logger.warning(f"Tipo/Formato inesperado para 'asistentes': {type(asistentes)}")
    if cuerpo is not None and isinstance(cuerpo, str): body["body"] = {"contentType": "HTML", "content": cuerpo}
    if ubicacion is not None and isinstance(ubicacion, str): body["location"] = {"displayName": ubicacion}
    if es_reunion_online: body["isOnlineMeeting"] = True
    if proveedor_reunion_online: body["onlineMeetingProvider"] = proveedor_reunion_online # Asume str
    if recordatorio_minutos is not None and isinstance(recordatorio_minutos, int): body["isReminderOn"] = True; body["reminderMinutesBeforeStart"] = recordatorio_minutos
    else: body["isReminderOn"] = False
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: POST {url}"); response = requests.post(url, headers=HEADERS, json=body); response.raise_for_status(); data = response.json(); logger.info(f"Evento '{titulo}' creado."); return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error crear evento: {e}. Detalles: {error_details}. URL: {url}"); raise Exception(f"Error crear evento: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (crear evento): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (crear evento): {e}")

# --- INICIO: VERSIÓN CORREGIDA DE actualizar_evento ---
def actualizar_evento(evento_id: str, nuevos_valores: Dict[str, Any], mailbox: Optional[str] = None) -> Dict[str, Any]:
    """Actualiza un evento existente."""
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"

    payload = nuevos_valores.copy() # Trabajar con copia

    # Procesar fecha de inicio si está presente y es datetime
    if 'start' in payload and isinstance(payload.get('start'), datetime):
        start_dt = payload['start']
        if start_dt.tzinfo is None:
            start_dt = start_dt.replace(tzinfo=timezone.utc)
        payload['start'] = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}
        logging.info(f"Procesada fecha 'start' para actualización: {payload['start']}")

    # Procesar fecha de fin si está presente y es datetime
    if 'end' in payload and isinstance(payload.get('end'), datetime):
        end_dt = payload['end']
        if end_dt.tzinfo is None:
            end_dt = end_dt.replace(tzinfo=timezone.utc)
        payload['end'] = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}
        logging.info(f"Procesada fecha 'end' para actualización: {payload['end']}")

    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: PATCH {url}")
        response = requests.patch(url, headers=HEADERS, json=payload) # Usar payload modificado
        response.raise_for_status()
        data = response.json()
        logger.info(f"Evento '{evento_id}' actualizado para usuario '{usuario}'.")
        return data
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al actualizar evento '{evento_id}': {e}. Detalles: {error_details}. URL: {url}")
        raise Exception(f"Error al actualizar evento: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"❌ Error al decodificar JSON (actualizar evento): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (actualizar evento): {e}")
# --- FIN: VERSIÓN CORREGIDA DE actualizar_evento ---

def eliminar_evento(evento_id: str, mailbox: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: DELETE {url}"); response = requests.delete(url, headers=HEADERS); response.raise_for_status(); logger.info(f"Evento '{evento_id}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error eliminar evento: {e}. URL: {url}. Detalles: {error_details}"); raise Exception(f"Error eliminar evento: {e}")


# ---- TEAMS y OTROS ----
# ... (Definiciones de listar_chats, listar_equipos, obtener_equipo, etc.) ...
# ... (Código omitido por brevedad, asumiendo que son las mismas que antes)...
def listar_chats(top: int = 20, skip: int = 0, filter_query: Optional[str] = None, order_by: Optional[str] = None, expand: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/me/chats"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
    if expand is not None and isinstance(expand, str): params['$expand'] = expand
    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params); response.raise_for_status(); data = response.json(); logger.info(f"Listados {len(data.get('value',[]))} chats."); return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error listar chats: {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}"); raise Exception(f"Error listar chats: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (listar chats): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar chats): {e}")

def listar_equipos(top: int = 20, skip: int = 0, filter_query: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/me/joinedTeams"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params); response.raise_for_status(); data = response.json(); logger.info(f"Listados {len(data.get('value',[]))} equipos."); return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error listar equipos: {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}"); raise Exception(f"Error listar equipos: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (listar equipos): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar equipos): {e}")

def obtener_equipo(team_id: str, select: Optional[List[str]]=None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/teams/{team_id}"
    params: Dict[str, Any] = {}
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params or None); response.raise_for_status(); data = response.json(); logger.info(f"Obtenido equipo ID: {team_id}."); return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error obtener equipo: {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}"); raise Exception(f"Error obtener equipo: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (obtener equipo): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (obtener equipo): {e}")


# --- FIN: Funciones Auxiliares de Graph API ---


# --- INICIO: Función Principal de Azure Functions (Entry Point) ---

# Mapeo de nombres de acción a funciones
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

    # Procesar solicitud y ejecutar acción
    try:
        # Leer accion/parametros (priorizar body JSON para POST, etc.)
        req_body: Optional[Dict[str, Any]] = None
        if req.method in ('POST', 'PUT', 'PATCH'):
            try:
                req_body = req.get_json()
                if not isinstance(req_body, dict):
                     logger.warning('Cuerpo JSON no es un diccionario.')
                     return func.HttpResponse("Cuerpo JSON debe ser un objeto.", status_code=400)
                accion = req_body.get('accion')
                params_input = req_body.get('parametros')
                if isinstance(params_input, dict): parametros = params_input
                elif params_input is not None: logger.warning(f"'parametros' no es dict"); parametros = {}
                else: parametros = {}
            except ValueError:
                logger.warning('No se pudo decodificar JSON del cuerpo.')
                return func.HttpResponse("Cuerpo JSON inválido.", status_code=400)
        else: # Fallback a Query Params (simplificado)
            accion = req.params.get('accion')
            parametros = dict(req.params) # Copiar todos los query params

        # Validar acción
        if not accion or not isinstance(accion, str):
            logger.warning("Clave 'accion' faltante o no es string.")
            return func.HttpResponse("Falta 'accion' (string).", status_code=400)

        logger.info(f"Acción a ejecutar: '{accion}'. Parámetros iniciales: {parametros}")

        # Buscar y ejecutar la función
        if accion in acciones_disponibles:
            funcion_a_ejecutar = acciones_disponibles[accion]
            logger.info(f"Preparando ejecución de: {funcion_a_ejecutar.__name__}")

            # --- Validar/Convertir parámetros ANTES de llamar ---
            # (Esta sección necesita ser robusta y cubrir todos los casos)
            params_procesados: Dict[str, Any] = {}
            try:
                # Convertir 'top' y 'skip' a int
                if accion in ["listar_correos", "listar_eventos", "listar_chats", "listar_equipos"]:
                    # Usa .get para evitar KeyError, provee default, convierte a int
                    params_procesados['top'] = int(parametros.get('top', 10))
                    params_procesados['skip'] = int(parametros.get('skip', 0))

                # Convertir fechas para crear_evento / actualizar_evento (ejemplo básico)
                if accion in ["crear_evento", "actualizar_evento"]:
                    for date_key in ['inicio', 'fin']:
                         if date_key in parametros:
                             date_val = parametros[date_key]
                             if isinstance(date_val, str):
                                 try:
                                     # Quitar Z y añadir offset UTC para compatibilidad
                                     params_procesados[date_key] = datetime.fromisoformat(date_val.replace('Z', '+00:00'))
                                 except ValueError:
                                     raise ValueError(f"Formato de fecha '{date_key}' inválido: {date_val}")
                             elif isinstance(date_val, datetime):
                                 params_procesados[date_key] = date_val # Ya es datetime
                             # else: Ignorar o fallar si no es str ni datetime?

                # Copiar el resto de parámetros sin sobreescribir los convertidos
                for k, v in parametros.items():
                    if k not in params_procesados:
                        params_procesados[k] = v

                # Aquí irían validaciones de parámetros requeridos específicos por acción
                # ...

            except (ValueError, TypeError) as conv_err:
                logger.error(f"Error en parámetros para '{accion}': {conv_err}. Recibido: {parametros}")
                return func.HttpResponse(f"Parámetros inválidos para '{accion}': {conv_err}", status_code=400)

            # Llamar a la función auxiliar
            logger.info(f"Ejecutando {funcion_a_ejecutar.__name__} con params: {params_procesados}")
            try:
                resultado = funcion_a_ejecutar(**params_procesados)
                logger.info(f"Ejecución de '{accion}' completada.")
            except Exception as exec_err:
                logger.exception(f"Error durante la ejecución de la acción '{accion}': {exec_err}")
                return func.HttpResponse(f"Error interno al ejecutar la acción '{accion}'.", status_code=500)

            # Devolver resultado
            try:
                 return func.HttpResponse(json.dumps(resultado, default=str), mimetype="application/json")
            except TypeError as serialize_err:
                 logger.error(f"Error al serializar resultado JSON para '{accion}': {serialize_err}.")
                 return func.HttpResponse(f"Error interno: Respuesta no serializable para {accion}.", status_code=500)

        else:
            logger.warning(f"Acción '{accion}' no reconocida.")
            acciones_validas = list(acciones_disponibles.keys())
            return func.HttpResponse(f"Acción '{accion}' no reconocida. Válidas: {acciones_validas}", status_code=400)

    except Exception as e:
        logger.exception(f"Error GENERAL inesperado: {e}")
        return func.HttpResponse("Error interno del servidor.", status_code=500)

# --- FIN: Función Principal ---
