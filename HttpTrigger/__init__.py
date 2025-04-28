import json
import logging
import requests
import azure.functions as func
from typing import Dict, Any, Callable, List, Optional, Union, Mapping, Sequence
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

# (Función obtener_token - sin cambios)
def obtener_token() -> str:
    logger.info("Obteniendo token de acceso...")
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {'client_id': CLIENT_ID, 'scope': GRAPH_SCOPE, 'client_secret': CLIENT_SECRET, 'grant_type': 'client_credentials'}
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    response = None
    try:
        response = requests.post(url, data=data, headers=headers)
        response.raise_for_status()
        token_data = response.json()
        token = token_data.get('access_token')
        if not token: logger.error(f"❌ No se encontró 'access_token'. Respuesta: {token_data}"); raise Exception("No se pudo obtener el token de acceso.")
        return token
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error red/HTTP (token): {e}. Detalles: {error_details}"); raise Exception(f"Error red/HTTP (token): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (token): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (token): {e}")
    except Exception as e: logger.error(f"❌ Error inesperado (token): {e}"); raise

# (Función _actualizar_headers - sin cambios)
def _actualizar_headers() -> None:
    try:
        token = obtener_token()
        HEADERS['Authorization'] = f'Bearer {token}'
        logger.info("Cabecera de autorización actualizada.")
    except Exception as e:
        logger.error(f"❌ Falló la actualización de la cabecera: {e}")
        raise Exception(f"Fallo al actualizar la cabecera: {e}")

# --- FIN: Constantes y Autenticación ---


# --- INICIO: Funciones Auxiliares de Graph API ---

# ---- CORREO ----
# (Función listar_correos - sin cambios respecto a la última versión)
def listar_correos(top: int = 10, skip: int = 0, folder: str = 'Inbox', select: Optional[List[str]] = None, filter_query: Optional[str] = None, order_by: Optional[str] = None, mailbox: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/mailFolders/{folder}/messages"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
    response: Optional[requests.Response] = None
    try:
        clean_params = {k:v for k, v in params.items() if v is not None}; logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params); response.raise_for_status(); data: Dict[str, Any] = response.json(); logger.info(f"Listados {len(data.get('value',[]))} correos."); return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error listar correos: {e}. URL: {url}. Detalles: {error_details}"); raise Exception(f"Error al listar correos: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (listar correos): {e}. Respuesta: {response_text}"); raise Exception(f"Error al decodificar JSON (listar correos): {e}")

# (Función leer_correo - sin cambios respecto a la última versión)
def leer_correo(message_id: str, select: Optional[List[str]] = None, mailbox: Optional[str] = None) -> dict:
     _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
     params = {}; response: Optional[requests.Response] = None
     if select: params['$select'] = ','.join(select)
     try:
         logger.info(f"Llamando a Graph API: GET {url} con params: {params}"); response = requests.get(url, headers=HEADERS, params=params or None); response.raise_for_status(); data = response.json(); logger.info(f"Correo '{message_id}' leído."); return data
     except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error leer correo: {e}. URL: {url}. Detalles: {error_details}"); raise Exception(f"Error leer correo: {e}")
     except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (leer correo): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (leer correo): {e}")

# (Función enviar_correo - sin cambios respecto a la última versión)
def enviar_correo(destinatario: Union[str, List[str]], asunto: str, mensaje: str, cc: Optional[Union[str, List[str]]] = None, bcc: Optional[Union[str, List[str]]] = None, attachments: Optional[List[dict]] = None, is_draft: bool = False, mailbox: Optional[str] = None) -> dict:
     _actualizar_headers(); usuario = mailbox or MAILBOX; endpoint = "messages" if is_draft else "sendMail"; url = f"{BASE_URL}/users/{usuario}/{endpoint}"
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

# !!!!! INICIO: FUNCIONES DE CORREO QUE FALTABAN !!!!!
def guardar_borrador(
    destinatario: Union[str, List[str]],
    asunto: str,
    mensaje: str,
    cc: Optional[Union[str, List[str]]] = None,
    bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[dict]] = None,
    mailbox: Optional[str] = None
) -> dict:
    """Guarda un correo como borrador llamando a enviar_correo con is_draft=True."""
    logger.info(f"Intentando guardar borrador para usuario '{mailbox or MAILBOX}' con asunto: '{asunto}'")
    # Llama a la función principal de envío con el flag de borrador
    return enviar_correo(destinatario, asunto, mensaje, cc, bcc, attachments, is_draft=True, mailbox=mailbox)

def enviar_borrador(
    message_id: str,
    mailbox: Optional[str] = None
) -> dict:
    """Envía un mensaje de correo que fue guardado como borrador."""
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/send"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: POST {url}")
        response = requests.post(url, headers=HEADERS)
        response.raise_for_status() # 202 Accepted esperado
        logger.info(f"Borrador de correo '{message_id}' enviado por usuario '{usuario}'.")
        return {"status": "Borrador Enviado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al enviar borrador '{message_id}' por usuario '{usuario}': {e}. URL: {url}. Detalles: {error_details}")
        raise Exception(f"Error al enviar borrador: {e}")

def responder_correo(
    message_id: str,
    mensaje_respuesta: str,
    to_recipients: Optional[List[dict]] = None, # Permite especificar destinatarios si es necesario sobreescribir
    reply_all: bool = False,
    mailbox: Optional[str] = None
) -> dict:
    """Responde a un mensaje de correo existente."""
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    action = "replyAll" if reply_all else "reply"
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/{action}"
    # El payload básico solo necesita 'comment'. Para cambiar destinatarios, etc., se anida 'message'.
    payload: Dict[str, Any] = {"comment": mensaje_respuesta}
    if to_recipients:
         # Estructura esperada si quieres sobreescribir destinatarios, cc, etc.
         payload["message"] = { "toRecipients": to_recipients }
         logger.info(f"Respondiendo con destinatarios personalizados.")

    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: POST {url}")
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status() # 202 Accepted esperado
        logger.info(f"Respuesta {'a todos ' if reply_all else ''}enviada al correo '{message_id}' por usuario '{usuario}'.")
        return {"status": "Respondido", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al responder {'a todos ' if reply_all else ''}al correo '{message_id}' por usuario '{usuario}': {e}. URL: {url}. Detalles: {error_details}")
        raise Exception(f"Error al responder al correo: {e}")

def reenviar_correo(
    message_id: str,
    destinatarios: Union[str, List[str]], # Se espera email(s)
    mensaje_reenvio: str = "FYI", # Comentario a añadir
    mailbox: Optional[str] = None
) -> dict:
    """Reenvía un mensaje de correo existente."""
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/forward"

    if isinstance(destinatarios, str): destinatarios = [destinatarios]
    # Construir lista de destinatarios en formato Graph API
    to_recipients_list = [{"emailAddress": {"address": r}} for r in destinatarios if r and isinstance(r, str)]
    if not to_recipients_list:
        raise ValueError("Se requiere al menos un destinatario válido (string email) para reenviar.")

    payload = {"toRecipients": to_recipients_list, "comment": mensaje_reenvio}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: POST {url}")
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status() # 202 Accepted esperado
        logger.info(f"Correo '{message_id}' reenviado por usuario '{usuario}' a: {destinatarios}.")
        return {"status": "Reenviado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al reenviar el correo '{message_id}' por usuario '{usuario}': {e}. URL: {url}. Detalles: {error_details}")
        raise Exception(f"Error al reenviar el correo: {e}")

def eliminar_correo(
    message_id: str,
    mailbox: Optional[str] = None
) -> dict:
    """Elimina un mensaje de correo (lo mueve a Elementos Eliminados)."""
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: DELETE {url}")
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status() # 204 No Content esperado
        logger.info(f"Correo '{message_id}' eliminado por usuario '{usuario}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al eliminar el correo '{message_id}' por usuario '{usuario}': {e}. URL: {url}. Detalles: {error_details}")
        raise Exception(f"Error al eliminar el correo: {e}")
# !!!!! FIN: FUNCIONES DE CORREO QUE FALTABAN !!!!!


# ---- CALENDARIO ----
# (Función listar_eventos - sin cambios respecto a la última versión)
def listar_eventos(top: int = 10, start_date: Optional[datetime] = None, end_date: Optional[datetime] = None, filter_query: Optional[str] = None, order_by: Optional[str] = None, select: Optional[List[str]] = None, use_calendar_view: bool = True, mailbox: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); usuario = mailbox or MAILBOX; base_endpoint = f"{BASE_URL}/users/{usuario}"; params: Dict[str, Any] = {}; endpoint_suffix = ""
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
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params); response.raise_for_status(); data = response.json(); logger.info(f"Listados eventos."); return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error listar eventos: {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}"); raise Exception(f"Error listar eventos: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (listar eventos): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar eventos): {e}")

# (Función crear_evento - sin cambios respecto a la última versión)
def crear_evento(titulo: str, inicio: datetime, fin: datetime, asistentes: Optional[List[Dict[str, Any]]] = None, cuerpo: Optional[str] = None, es_reunion_online: bool = False, proveedor_reunion_online: str = "teamsForBusiness", recordatorio_minutos: Optional[int] = 15, ubicacion: Optional[str] = None, mostrar_como: str = "busy", mailbox: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/events"
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
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: POST {url}"); response = requests.post(url, headers=HEADERS, json=body); response.raise_for_status(); data = response.json(); logger.info(f"Evento '{titulo}' creado."); return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error crear evento: {e}. Detalles: {error_details}. URL: {url}"); raise Exception(f"Error crear evento: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (crear evento): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (crear evento): {e}")

# (Función actualizar_evento - versión CORREGIDA de multi-línea)
def actualizar_evento(evento_id: str, nuevos_valores: Dict[str, Any], mailbox: Optional[str] = None) -> Dict[str, Any]:
    """Actualiza un evento existente."""
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
    payload = nuevos_valores.copy()
    # Procesar fecha de inicio
    if 'start' in payload and isinstance(payload.get('start'), datetime):
        start_dt = payload['start']
        if start_dt.tzinfo is None: start_dt = start_dt.replace(tzinfo=timezone.utc)
        payload['start'] = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}
        logging.info(f"Procesada fecha 'start' para actualización: {payload['start']}")
    # Procesar fecha de fin
    if 'end' in payload and isinstance(payload.get('end'), datetime):
        end_dt = payload['end']
        if end_dt.tzinfo is None: end_dt = end_dt.replace(tzinfo=timezone.utc)
        payload['end'] = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}
        logging.info(f"Procesada fecha 'end' para actualización: {payload['end']}")
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: PATCH {url}"); response = requests.patch(url, headers=HEADERS, json=payload); response.raise_for_status(); data = response.json(); logger.info(f"Evento '{evento_id}' actualizado."); return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error actualizar evento: {e}. Detalles: {error_details}. URL: {url}"); raise Exception(f"Error actualizar evento: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (actualizar evento): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (actualizar evento): {e}")

# (Función eliminar_evento - sin cambios respecto a la última versión)
def eliminar_evento(evento_id: str, mailbox: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: DELETE {url}"); response = requests.delete(url, headers=HEADERS); response.raise_for_status(); logger.info(f"Evento '{evento_id}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error eliminar evento: {e}. URL: {url}. Detalles: {error_details}"); raise Exception(f"Error eliminar evento: {e}")

# ---- TEAMS y OTROS ----
# (Funciones listar_chats, listar_equipos, obtener_equipo - sin cambios respecto a la última versión)
def listar_chats(top: int = 20, skip: int = 0, filter_query: Optional[str] = None, order_by: Optional[str] = None, expand: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/me/chats"; params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
    if expand is not None and isinstance(expand, str): params['$expand'] = expand
    clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params); response.raise_for_status(); data = response.json(); logger.info(f"Listados {len(data.get('value',[]))} chats."); return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error listar chats: {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}"); raise Exception(f"Error listar chats: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (listar chats): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar chats): {e}")

def listar_equipos(top: int = 20, skip: int = 0, filter_query: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/me/joinedTeams"; params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params); response.raise_for_status(); data = response.json(); logger.info(f"Listados {len(data.get('value',[]))} equipos."); return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logger.error(f"❌ Error listar equipos: {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}"); raise Exception(f"Error listar equipos: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logger.error(f"❌ Error JSON (listar equipos): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar equipos): {e}")

def obtener_equipo(team_id: str, select: Optional[List[str]]=None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/teams/{team_id}"; params: Dict[str, Any] = {}
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
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
    "guardar_borrador": guardar_borrador,   # <-- Ahora definida arriba
    "enviar_borrador": enviar_borrador,     # <-- Ahora definida arriba
    "responder_correo": responder_correo,   # <-- Ahora definida arriba
    "reenviar_correo": reenviar_correo,     # <-- Ahora definida arriba
    "eliminar_correo": eliminar_correo,     # <-- Ahora definida arriba
    "listar_eventos": listar_eventos,
    "crear_evento": crear_evento,
    "actualizar_evento": actualizar_evento, # <-- Ahora definida arriba
    "eliminar_evento": eliminar_evento,     # <-- Ahora definida arriba
    "listar_chats": listar_chats,
    "listar_equipos": listar_equipos,
    "obtener_equipo": obtener_equipo,
}

# (Función main - sin cambios respecto a la última versión con validaciones/conversiones)
def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info(f'Python HTTP trigger function procesando una solicitud. Method: {req.method}')
    accion: Optional[str] = None; parametros: Dict[str, Any] = {}
    try:
        req_body: Optional[Dict[str, Any]] = None
        if req.method in ('POST', 'PUT', 'PATCH'):
            try:
                req_body = req.get_json(); assert isinstance(req_body, dict)
                accion = req_body.get('accion'); params_input = req_body.get('parametros')
                if isinstance(params_input, dict): parametros = params_input
                elif params_input is not None: logger.warning(f"'parametros' no es dict"); parametros = {}
                else: parametros = {}
            except ValueError: logger.warning('Cuerpo no es JSON válido.'); return func.HttpResponse("Cuerpo JSON inválido.", status_code=400)
            except AssertionError: logger.warning('Cuerpo JSON no es un objeto.'); return func.HttpResponse("Cuerpo JSON debe ser un objeto.", status_code=400)
        else: accion = req.params.get('accion'); parametros = dict(req.params) # Simplificado para GET

        if not accion or not isinstance(accion, str): logger.warning("Clave 'accion' faltante o no es string."); return func.HttpResponse("Falta 'accion' (string).", status_code=400)
        logger.info(f"Acción a ejecutar: '{accion}'. Parámetros iniciales: {parametros}")

        if accion in acciones_disponibles:
            funcion_a_ejecutar = acciones_disponibles[accion]; logger.info(f"Preparando ejecución de: {funcion_a_ejecutar.__name__}")
            params_procesados: Dict[str, Any] = {};
            try: # Bloque para procesar parámetros
                if accion in ["listar_correos", "listar_eventos", "listar_chats", "listar_equipos"]:
                    if 'top' in parametros: params_procesados['top'] = int(parametros['top'])
                    if 'skip' in parametros: params_procesados['skip'] = int(parametros['skip'])
                elif accion in ["crear_evento", "actualizar_evento"]:
                     for date_key in ['inicio', 'fin']:
                         if date_key in parametros:
                             date_val = parametros[date_key]
                             if isinstance(date_val, str):
                                 try: params_procesados[date_key] = datetime.fromisoformat(date_val.replace('Z', '+00:00'))
                                 except ValueError: raise ValueError(f"Formato fecha '{date_key}' inválido: {date_val}")
                             elif isinstance(date_val, datetime): params_procesados[date_key] = date_val
                # Copiar el resto
                for k, v in parametros.items():
                    if k not in params_procesados: params_procesados[k] = v
                # Validaciones específicas (ejemplo)
                # ... (añadir más validaciones si es necesario) ...
            except (ValueError, TypeError) as conv_err: logger.error(f"Error en parámetros para '{accion}': {conv_err}. Recibido: {parametros}"); return func.HttpResponse(f"Parámetros inválidos para '{accion}': {conv_err}", status_code=400)

            # Llamar función
            logger.info(f"Ejecutando {funcion_a_ejecutar.__name__} con params: {params_procesados}")
            try:
                resultado = funcion_a_ejecutar(**params_procesados); logger.info(f"Ejecución de '{accion}' completada.")
            except Exception as exec_err: logger.exception(f"Error durante ejecución acción '{accion}': {exec_err}"); return func.HttpResponse(f"Error interno al ejecutar '{accion}'.", status_code=500)
            # Devolver resultado
            try: return func.HttpResponse(json.dumps(resultado, default=str), mimetype="application/json")
            except TypeError as serialize_err: logger.error(f"Error al serializar resultado JSON para '{accion}': {serialize_err}."); return func.HttpResponse(f"Error interno: Respuesta no serializable.", status_code=500)
        else:
            logger.warning(f"Acción '{accion}' no reconocida."); acciones_validas = list(acciones_disponibles.keys()); return func.HttpResponse(f"Acción '{accion}' no reconocida. Válidas: {acciones_validas}", status_code=400)
    except Exception as e: logger.exception(f"Error GENERAL inesperado: {e}"); return func.HttpResponse("Error interno del servidor.", status_code=500)

# --- FIN: Función Principal ---
