# HttpTrigger/__init__.py (Versión Completa y Unificada)

import json
import logging
import requests
import azure.functions as func
from typing import Dict, Any, Callable, List, Optional, Union, Mapping, Sequence
from datetime import datetime, timezone
import os

# --- Configuración de Logging ---
logger = logging.getLogger("azure.functions")
logger.setLevel(logging.INFO) # O logging.DEBUG para más detalle

# --- Variables de Entorno y Configuración ---
def get_config_or_raise(key: str, default: Optional[str] = None) -> str:
    value = os.environ.get(key, default)
    if value is None:
        logger.error(f"CONFIG ERROR: Falta la variable de entorno requerida: {key}")
        raise ValueError(f"Configuración esencial faltante: {key}")
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

# --- Constantes y Autenticación ---
BASE_URL = "https://graph.microsoft.com/v1.0"
# Headers globales que se actualizarán con el token
HEADERS: Dict[str, Optional[str]] = {
    'Authorization': None,
    'Content-Type': 'application/json'
}
# Timeout general para llamadas a la API de Graph (en segundos)
GRAPH_API_TIMEOUT = 45 # Aumentado ligeramente por si acaso

def obtener_token() -> str:
    """Obtiene un token de acceso de aplicación usando credenciales de cliente."""
    logger.info("Obteniendo token de acceso de aplicación...")
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {'client_id': CLIENT_ID, 'scope': GRAPH_SCOPE, 'client_secret': CLIENT_SECRET, 'grant_type': 'client_credentials'}
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    response = None
    try:
        response = requests.post(url, data=data, headers=headers, timeout=GRAPH_API_TIMEOUT) # Timeout añadido
        response.raise_for_status()
        token_data = response.json()
        token = token_data.get('access_token')
        if not token: logger.error(f"No se encontró 'access_token'. Respuesta: {token_data}"); raise Exception("No se pudo obtener el token de acceso.")
        return token
    except requests.exceptions.Timeout: logger.error(f"Timeout al obtener token desde {url}"); raise Exception("Timeout al contactar el servidor de autenticación.")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error de red/HTTP al obtener token: {e}. Detalles: {error_details}"); raise Exception(f"Error de red/HTTP al obtener token: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response'); logger.error(f"Error al decodificar JSON del token: {e}. Respuesta: {response_text}"); raise Exception(f"Error al decodificar JSON del token: {e}")
    except Exception as e: logger.error(f"Error inesperado al obtener token: {e}"); raise

def _actualizar_headers() -> None:
    """Obtiene un token fresco y actualiza los HEADERS globales."""
    try:
        token = obtener_token()
        HEADERS['Authorization'] = f'Bearer {token}'
        logger.info("Cabecera de autorización actualizada.")
    except Exception as e:
        logger.error(f"Falló la actualización de la cabecera: {e}")
        raise Exception(f"Fallo al actualizar la cabecera: {e}")

# --- FIN: Constantes y Autenticación ---

# --- INICIO: Funciones Auxiliares de Graph API (Todas Integradas Aquí) ---

# ---- CORREO ----
def listar_correos(top: int = 10, skip: int = 0, folder: str = 'Inbox', select: Optional[List[str]] = None, filter_query: Optional[str] = None, order_by: Optional[str] = None, mailbox: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/mailFolders/{folder}/messages"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
    response: Optional[requests.Response] = None
    try:
        clean_params = {k:v for k, v in params.items() if v is not None}; logger.info(f"API Call: GET {url} Params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data: Dict[str, Any] = response.json(); logger.info(f"Listados {len(data.get('value',[]))} correos."); return data
    except requests.exceptions.Timeout: logger.error(f"Timeout listando correos: {url}"); raise Exception("Timeout API Graph (listar correos).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (listar correos): {e}. Det: {error_details}"); raise Exception(f"Error API (listar correos): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response'); logger.error(f"Error JSON (listar correos): {e}. Resp: {response_text}"); raise Exception(f"Error JSON (listar correos): {e}")

def leer_correo(message_id: str, select: Optional[List[str]] = None, mailbox: Optional[str] = None) -> dict:
     _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
     params = {}; response: Optional[requests.Response] = None
     if select and isinstance(select, list): params['$select'] = ','.join(select)
     try:
         logger.info(f"API Call: GET {url} Params: {params}"); response = requests.get(url, headers=HEADERS, params=params or None, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Correo '{message_id}' leído."); return data
     except requests.exceptions.Timeout: logger.error(f"Timeout leyendo correo: {url}"); raise Exception("Timeout API Graph (leer correo).")
     except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (leer correo): {e}. Det: {error_details}"); raise Exception(f"Error API (leer correo): {e}")
     except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response'); logger.error(f"Error JSON (leer correo): {e}. Resp: {response_text}"); raise Exception(f"Error JSON (leer correo): {e}")

def enviar_correo(destinatario: Union[str, List[str]], asunto: str, mensaje: str, cc: Optional[Union[str, List[str]]] = None, bcc: Optional[Union[str, List[str]]] = None, attachments: Optional[List[dict]] = None, from_email: Optional[str] = None, is_draft: bool = False, mailbox: Optional[str] = None) -> dict:
    _actualizar_headers(); usuario = mailbox or MAILBOX
    if is_draft: url = f"{BASE_URL}/users/{usuario}/messages"
    else: url = f"{BASE_URL}/users/{usuario}/sendMail"
    if isinstance(destinatario, str): destinatario_list = [destinatario]
    elif isinstance(destinatario, list): destinatario_list = destinatario
    else: raise TypeError("Destinatario debe ser str o List[str]")
    to_recipients = [{"emailAddress": {"address": r}} for r in destinatario_list if r and isinstance(r, str)]
    cc_recipients = []
    if cc:
        if isinstance(cc, str): cc_list = [cc]
        elif isinstance(cc, list): cc_list = cc
        else: raise TypeError("CC debe ser str o List[str]")
        cc_recipients = [{"emailAddress": {"address": r}} for r in cc_list if r and isinstance(r, str)]
    bcc_recipients = []
    if bcc:
        if isinstance(bcc, str): bcc_list = [bcc]
        elif isinstance(bcc, list): bcc_list = bcc
        else: raise TypeError("BCC debe ser str o List[str]")
        bcc_recipients = [{"emailAddress": {"address": r}} for r in bcc_list if r and isinstance(r, str)]
    if not to_recipients: logging.error("No destinatarios válidos."); raise ValueError("Se requiere destinatario válido.")
    message_payload: Dict[str, Any] = {"subject": asunto, "body": {"contentType": "HTML", "content": mensaje},"toRecipients": to_recipients,}
    if cc_recipients: message_payload["ccRecipients"] = cc_recipients
    if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
    if attachments: message_payload["attachments"] = attachments
    if from_email: message_payload["from"] = {"emailAddress": {"address": from_email}}
    final_payload = {"message": message_payload, "saveToSentItems": "true"} if not is_draft else message_payload
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url}"); response = requests.post(url, headers=HEADERS, json=final_payload, timeout=GRAPH_API_TIMEOUT); response.raise_for_status()
        if not is_draft: logger.info(f"Correo enviado."); return {"status": "Enviado", "code": response.status_code}
        else: data = response.json(); message_id = data.get('id'); logger.info(f"Borrador guardado ID: {message_id}."); return {"status": "Borrador Guardado", "code": response.status_code, "id": message_id, "data": data}
    except requests.exceptions.Timeout: logger.error(f"Timeout enviando/guardando correo: {url}"); raise Exception("Timeout API Graph (enviar/guardar correo).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (enviar/guardar correo): {e}. Det: {error_details}"); raise Exception(f"Error API (enviar/guardar correo): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response'); logger.error(f"Error JSON (guardar borrador): {e}. Resp: {response_text}"); raise Exception(f"Error JSON (guardar borrador): {e}")

def guardar_borrador(destinatario: Union[str, List[str]], asunto: str, mensaje: str, cc: Optional[Union[str, List[str]]] = None, bcc: Optional[Union[str, List[str]]] = None, attachments: Optional[List[dict]] = None, from_email: Optional[str] = None, mailbox: Optional[str] = None) -> dict:
    logger.info(f"Guardando borrador para '{mailbox or MAILBOX}'. Asunto: '{asunto}'"); return enviar_correo(destinatario, asunto, mensaje, cc, bcc, attachments, from_email, is_draft=True, mailbox=mailbox)

def enviar_borrador(message_id: str, mailbox: Optional[str] = None) -> dict:
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/send"; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url}"); response = requests.post(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Borrador '{message_id}' enviado."); return {"status": "Borrador Enviado", "code": response.status_code}
    except requests.exceptions.Timeout: logger.error(f"Timeout enviando borrador: {url}"); raise Exception("Timeout API Graph (enviar borrador).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (enviar borrador '{message_id}'): {e}. Det: {error_details}"); raise Exception(f"Error API (enviar borrador '{message_id}'): {e}")

def responder_correo(message_id: str, mensaje_respuesta: str, to_recipients: Optional[List[dict]] = None, reply_all: bool = False, mailbox: Optional[str] = None) -> dict:
    _actualizar_headers(); usuario = mailbox or MAILBOX; action = "replyAll" if reply_all else "reply"; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/{action}"
    payload: Dict[str, Any] = {"comment": mensaje_respuesta}
    if to_recipients: payload["message"] = { "toRecipients": to_recipients }; logger.info(f"Respondiendo con destinatarios custom.")
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url}"); response = requests.post(url, headers=HEADERS, json=payload, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Respuesta {'a todos ' if reply_all else ''}enviada correo '{message_id}'."); return {"status": "Respondido", "code": response.status_code}
    except requests.exceptions.Timeout: logger.error(f"Timeout respondiendo correo: {url}"); raise Exception("Timeout API Graph (responder correo).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (responder correo '{message_id}'): {e}. Det: {error_details}"); raise Exception(f"Error API (responder correo): {e}")

def reenviar_correo(message_id: str, destinatarios: Union[str, List[str]], mensaje_reenvio: str = "FYI", mailbox: Optional[str] = None) -> dict:
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/forward"
    if isinstance(destinatarios, str): destinatarios = [destinatarios]
    to_recipients_list = [{"emailAddress": {"address": r}} for r in destinatarios if r and isinstance(r, str)]
    if not to_recipients_list: raise ValueError("Se requiere destinatario válido para reenviar.")
    payload = {"toRecipients": to_recipients_list, "comment": mensaje_reenvio}; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url}"); response = requests.post(url, headers=HEADERS, json=payload, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Correo '{message_id}' reenviado."); return {"status": "Reenviado", "code": response.status_code}
    except requests.exceptions.Timeout: logger.error(f"Timeout reenviando correo: {url}"); raise Exception("Timeout API Graph (reenviar correo).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (reenviar correo '{message_id}'): {e}. Det: {error_details}"); raise Exception(f"Error API (reenviar correo): {e}")

def eliminar_correo(message_id: str, mailbox: Optional[str] = None) -> dict:
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url}"); response = requests.delete(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Correo '{message_id}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.Timeout: logger.error(f"Timeout eliminando correo: {url}"); raise Exception("Timeout API Graph (eliminar correo).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (eliminar correo '{message_id}'): {e}. Det: {error_details}"); raise Exception(f"Error API (eliminar correo): {e}")

# ---- CALENDARIO ----
def listar_eventos(top: int = 10, start_date: Optional[datetime] = None, end_date: Optional[datetime] = None, filter_query: Optional[str] = None, order_by: Optional[str] = None, select: Optional[List[str]] = None, use_calendar_view: bool = True, mailbox: Optional[str] = 'me') -> Dict[str, Any]:
    _actualizar_headers(); usuario = mailbox if mailbox else 'me'; base_endpoint = f"https://graph.microsoft.com/v1.0/users/{usuario}"; params: Dict[str, Any] = {}; endpoint_suffix = ""
    if use_calendar_view and start_date and end_date:
        endpoint_suffix = "/calendarView";
        if isinstance(start_date, datetime) and start_date.tzinfo is None: start_date = start_date.replace(tzinfo=timezone.utc)
        if isinstance(end_date, datetime) and end_date.tzinfo is None: end_date = end_date.replace(tzinfo=timezone.utc)
        if isinstance(start_date, datetime): params['startDateTime'] = start_date.isoformat()
        if isinstance(end_date, datetime): params['endDateTime'] = end_date.isoformat()
        params['$top'] = int(top);
        if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
        if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
        if select and isinstance(select, list): params['$select'] = ','.join(select)
    else:
        endpoint_suffix = "/events"; params['$top'] = int(top); filters = []
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
    url = f"{base_endpoint}{endpoint_suffix}"; clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} Params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Listados eventos."); return data
    except requests.exceptions.Timeout: logger.error(f"Timeout listando eventos: {url}"); raise Exception("Timeout API Graph (listar eventos).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (listar eventos): {e}. Det: {error_details}"); raise Exception(f"Error API (listar eventos): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response'); logger.error(f"Error JSON (listar eventos): {e}. Resp: {response_text}"); raise Exception(f"Error JSON (listar eventos): {e}")

def crear_evento(titulo: str, inicio: datetime, fin: datetime, asistentes: Optional[List[Dict[str, Any]]] = None, cuerpo: Optional[str] = None, es_reunion_online: bool = False, proveedor_reunion_online: str = "teamsForBusiness", recordatorio_minutos: Optional[int] = 15, ubicacion: Optional[str] = None, mostrar_como: str = "busy", mailbox: Optional[str] = 'me') -> Dict[str, Any]:
    _actualizar_headers(); usuario = mailbox if mailbox else 'me'; url = f"https://graph.microsoft.com/v1.0/users/{usuario}/events"
    if not isinstance(inicio, datetime) or not isinstance(fin, datetime): raise ValueError("'inicio' y 'fin' deben ser datetimes.")
    if inicio.tzinfo is None: inicio = inicio.replace(tzinfo=timezone.utc);
    if fin.tzinfo is None: fin = fin.replace(tzinfo=timezone.utc);
    body: Dict[str, Any] = {"subject": titulo, "start": {"dateTime": inicio.isoformat(), "timeZone": "UTC"}, "end": {"dateTime": fin.isoformat(), "timeZone": "UTC"}, "showAs": mostrar_como}
    if asistentes is not None:
        if isinstance(asistentes, list) and all(isinstance(a, dict) for a in asistentes): body["attendees"] = [{"emailAddress": {"address": a.get('emailAddress')},"type": a.get('type', 'required')} for a in asistentes if a and a.get('emailAddress')]
        else: logger.warning(f"Tipo/Formato inesperado para 'asistentes': {type(asistentes)}")
    if cuerpo is not None and isinstance(cuerpo, str): body["body"] = {"contentType": "HTML", "content": cuerpo}
    if ubicacion is not None and isinstance(ubicacion, str): body["location"] = {"displayName": ubicacion}
    if es_reunion_online: body["isOnlineMeeting"] = True
    if proveedor_reunion_online: body["onlineMeetingProvider"] = proveedor_reunion_online
    if recordatorio_minutos is not None and isinstance(recordatorio_minutos, int): body["isReminderOn"] = True; body["reminderMinutesBeforeStart"] = recordatorio_minutos
    else: body["isReminderOn"] = False
    if mostrar_como: body["showAs"] = mostrar_como
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url}"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Evento '{titulo}' creado."); return data
    except requests.exceptions.Timeout: logger.error(f"Timeout creando evento: {url}"); raise Exception("Timeout API Graph (crear evento).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (crear evento): {e}. Det: {error_details}"); raise Exception(f"Error API (crear evento): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response'); logger.error(f"Error JSON (crear evento): {e}. Resp: {response_text}"); raise Exception(f"Error JSON (crear evento): {e}")

def actualizar_evento(evento_id: str, nuevos_valores: Dict[str, Any], mailbox: Optional[str] = 'me') -> Dict[str, Any]:
    _actualizar_headers(); usuario = mailbox if mailbox else 'me'; url = f"https://graph.microsoft.com/v1.0/users/{usuario}/events/{evento_id}"
    payload = nuevos_valores.copy()
    if 'start' in payload and isinstance(payload.get('start'), datetime): start_dt = payload['start'];
    if start_dt.tzinfo is None: start_dt = start_dt.replace(tzinfo=timezone.utc); payload['start'] = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}; logging.info(f"Proc. fecha 'start': {payload['start']}")
    if 'end' in payload and isinstance(payload.get('end'), datetime): end_dt = payload['end'];
    if end_dt.tzinfo is None: end_dt = end_dt.replace(tzinfo=timezone.utc); payload['end'] = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}; logging.info(f"Proc. fecha 'end': {payload['end']}")
    response: Optional[requests.Response] = None
    try:
        etag = payload.pop('@odata.etag', None); current_headers = HEADERS.copy()
        if etag: current_headers['If-Match'] = etag; logging.info(f"Usando ETag evento {evento_id}")
        logger.info(f"API Call: PATCH {url}"); response = requests.patch(url, headers=current_headers, json=payload, timeout=GRAPH_API_TIMEOUT); response.raise_for_status();
        logging.info(f"Evento '{evento_id}' actualizado.");
        if response.status_code == 204: return {"status": "Actualizado (No Content)", "id": evento_id}
        else: return response.json()
    except requests.exceptions.Timeout: logger.error(f"Timeout actualizando evento: {url}"); raise Exception("Timeout API Graph (actualizar evento).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (actualizar evento '{evento_id}'): {e}. Det: {error_details}"); raise Exception(f"Error API (actualizar evento): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response'); logger.error(f"Error JSON (actualizar evento): {e}. Resp: {response_text}"); raise Exception(f"Error JSON (actualizar evento): {e}")

def eliminar_evento(evento_id: str, mailbox: Optional[str] = 'me') -> Dict[str, Any]:
    _actualizar_headers(); usuario = mailbox if mailbox else 'me'; url = f"https://graph.microsoft.com/v1.0/users/{usuario}/events/{evento_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url}"); response = requests.delete(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Evento '{evento_id}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.Timeout: logger.error(f"Timeout eliminando evento: {url}"); raise Exception("Timeout API Graph (eliminar evento).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (eliminar evento '{evento_id}'): {e}. Det: {error_details}"); raise Exception(f"Error API (eliminar evento): {e}")

# ---- TEAMS y OTROS ----
def listar_chats(top: int = 20, skip: int = 0, filter_query: Optional[str] = None, order_by: Optional[str] = None, expand: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"https://graph.microsoft.com/v1.0/me/chats"; params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
    if expand is not None and isinstance(expand, str): params['$expand'] = expand
    clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} Params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Listados {len(data.get('value',[]))} chats."); return data
    except requests.exceptions.Timeout: logger.error(f"Timeout listando chats: {url}"); raise Exception("Timeout API Graph (listar chats).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (listar chats): {e}. Det: {error_details}"); raise Exception(f"Error API (listar chats): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response'); logger.error(f"Error JSON (listar chats): {e}. Resp: {response_text}"); raise Exception(f"Error JSON (listar chats): {e}")

def listar_equipos(top: int = 20, skip: int = 0, filter_query: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"https://graph.microsoft.com/v1.0/me/joinedTeams"; params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} Params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Listados {len(data.get('value',[]))} equipos."); return data
    except requests.exceptions.Timeout: logger.error(f"Timeout listando equipos: {url}"); raise Exception("Timeout API Graph (listar equipos).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (listar equipos): {e}. Det: {error_details}"); raise Exception(f"Error API (listar equipos): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response'); logger.error(f"Error JSON (listar equipos): {e}. Resp: {response_text}"); raise Exception(f"Error JSON (listar equipos): {e}")

def obtener_equipo(team_id: str, select: Optional[List[str]]=None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"https://graph.microsoft.com/v1.0/teams/{team_id}"; params: Dict[str, Any] = {}
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} Params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params or None, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Obtenido equipo ID: {team_id}."); return data
    except requests.exceptions.Timeout: logger.error(f"Timeout obteniendo equipo: {url}"); raise Exception("Timeout API Graph (obtener equipo).")
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"Error API (obtener equipo): {e}. Det: {error_details}"); raise Exception(f"Error API (obtener equipo): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response'); logger.error(f"Error JSON (obtener equipo): {e}. Resp: {response_text}"); raise Exception(f"Error JSON (obtener equipo): {e}")

# --- FIN: Funciones Auxiliares de Graph API ---


# --- INICIO: Función Principal de Azure Functions (Entry Point) ---

# Mapeo de nombres de acción a las funciones DEFINIDAS ARRIBA
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
    # Añade aquí cualquier otra acción/función auxiliar que hayas definido arriba
}
# Verificar que todas las funciones mapeadas existen (opcional, para desarrollo)
for accion_check, func_ref_check in acciones_disponibles.items():
    if not callable(func_ref_check):
        logger.error(f"Config Error: La función para la acción '{accion_check}' no es válida o no está definida.")

def main(req: func.HttpRequest) -> func.HttpResponse:
    """Punto de entrada principal. Maneja la solicitud HTTP, llama a la acción apropiada y devuelve la respuesta."""
    logging.info(f'Python HTTP trigger function procesando solicitud. Method: {req.method}, URL: {req.url}')
    invocation_id = req.headers.get('X-Azure-Functions-InvocationId', 'N/A')
    logging.info(f"Invocation ID: {invocation_id}")

    accion: Optional[str] = None
    parametros: Dict[str, Any] = {}
    funcion_a_ejecutar: Optional[Callable] = None

    # --- INICIO: Bloque Try-Except General ---
    try:
        # --- Leer accion/parametros ---
        req_body: Optional[Dict[str, Any]] = None
        if req.method in ('POST', 'PUT', 'PATCH'):
            try:
                req_body = req.get_json(); assert isinstance(req_body, dict)
                accion = req_body.get('accion'); params_input = req_body.get('parametros')
                if isinstance(params_input, dict): parametros = params_input
                elif params_input is not None: logger.warning(f"Invocation {invocation_id}: 'parametros' no es dict"); parametros = {}
                else: parametros = {}
            except ValueError: logger.warning(f'Invocation {invocation_id}: Cuerpo no es JSON válido.'); return func.HttpResponse("Cuerpo JSON inválido.", status_code=400)
            except AssertionError: logger.warning(f'Invocation {invocation_id}: Cuerpo JSON no es un objeto.'); return func.HttpResponse("Cuerpo JSON debe ser un objeto.", status_code=400)
        else: accion = req.params.get('accion'); parametros = dict(req.params)

        # --- Validar acción ---
        if not accion or not isinstance(accion, str): logger.warning(f"Invocation {invocation_id}: Clave 'accion' faltante o no es string."); return func.HttpResponse("Falta 'accion' (string).", status_code=400)
        logger.info(f"Invocation {invocation_id}: Acción solicitada: '{accion}'. Parámetros iniciales: {parametros}")

        # --- Buscar y ejecutar la función ---
        if accion in acciones_disponibles:
            funcion_a_ejecutar = acciones_disponibles[accion]
            logger.info(f"Invocation {invocation_id}: Mapeado a función: {funcion_a_ejecutar.__name__}")

            # --- Validar/Convertir parámetros ANTES de llamar ---
            params_procesados: Dict[str, Any] = {}
            try:
                params_procesados = parametros.copy()
                if accion in ["listar_correos", "listar_eventos", "listar_chats", "listar_equipos"]:
                    if 'top' in params_procesados: params_procesados['top'] = int(params_procesados['top'])
                    if 'skip' in params_procesados: params_procesados['skip'] = int(params_procesados['skip'])
                elif accion in ["crear_evento", "actualizar_evento"]:
                     for date_key in ['inicio', 'fin']:
                         if date_key in params_procesados:
                             date_val = params_procesados[date_key]
                             if isinstance(date_val, str):
                                 try: params_procesados[date_key] = datetime.fromisoformat(date_val.replace('Z', '+00:00'))
                                 except ValueError: raise ValueError(f"Formato fecha '{date_key}' inválido: {date_val}")
                             elif not isinstance(date_val, datetime): raise ValueError(f"Tipo inválido para '{date_key}'.")
                # ... (añadir más conversiones/validaciones aquí) ...

            except (ValueError, TypeError, KeyError) as conv_err:
                logger.error(f"Invocation {invocation_id}: Error en parámetros para '{accion}': {conv_err}. Recibido: {parametros}")
                return func.HttpResponse(f"Parámetros inválidos para '{accion}': {conv_err}", status_code=400)

            # !!!!! INICIO LOGGING ADICIONAL !!!!!
            logger.info(f"DEBUG Invocation {invocation_id}: Tipo de funcion_a_ejecutar: {type(funcion_a_ejecutar)}")
            logger.info(f"DEBUG Invocation {invocation_id}: Argumentos a pasar (params_procesados): {params_procesados}")
            logger.info(f"DEBUG Invocation {invocation_id}: Tipo de params_procesados: {type(params_procesados)}")
            # !!!!! FIN LOGGING ADICIONAL !!!!!

            # --- Llamar a la función auxiliar ---
            logger.info(f"Invocation {invocation_id}: Ejecutando {funcion_a_ejecutar.__name__}...")
            try:
                # Usamos la llamada genérica
                resultado = funcion_a_ejecutar(**params_procesados)
                logger.info(f"Invocation {invocation_id}: Ejecución de '{accion}' completada.")
            except Exception as exec_err:
                logger.exception(f"Invocation {invocation_id}: Error durante ejecución acción '{accion}': {exec_err}")
                return func.HttpResponse(f"Error interno al ejecutar '{accion}'.", status_code=500)

            # --- Devolver resultado ---
            try:
                 return func.HttpResponse(json.dumps(resultado, default=str), mimetype="application/json", status_code=200)
            except TypeError as serialize_err:
                 logger.error(f"Invocation {invocation_id}: Error al serializar resultado JSON para '{accion}': {serialize_err}.")
                 return func.HttpResponse(f"Error interno: Respuesta no serializable.", status_code=500)

        else: # Acción no encontrada
            logger.warning(f"Invocation {invocation_id}: Acción '{accion}' no reconocida."); acciones_validas = list(acciones_disponibles.keys());
            return func.HttpResponse(f"Acción '{accion}' no reconocida. Válidas: {acciones_validas}", status_code=400)

    # --- FIN: Bloque Try-Except General ---
    except Exception as e:
        func_name = getattr(funcion_a_ejecutar, '__name__', 'N/A') if funcion_a_ejecutar else 'N/A'
        logger.exception(f"Invocation {invocation_id}: Error GENERAL INESPERADO en main() procesando acción '{accion or 'desconocida'}' (Función: {func_name}): {e}")
        return func.HttpResponse("Error interno del servidor. Revise los logs.", status_code=500)

# --- FIN: Función Principal ---

