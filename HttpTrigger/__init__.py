# HttpTrigger/__init__.py (Original + Funciones SP/OD Añadidas)

import json
import logging
import requests
import azure.functions as func
from typing import Dict, Any, Callable, List, Optional, Union, Mapping, Sequence # Añadido Union de archivos SP/OD
from datetime import datetime, timezone
import os
import io # Necesario para subidas

# --- Configuración de Logging (Original) ---
logger = logging.getLogger("azure.functions")
logger.setLevel(logging.INFO) # O logging.DEBUG para más detalle

# --- Variables de Entorno y Configuración (Original) ---
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
    MAILBOX = get_config_or_raise('MAILBOX', default='me') # Mantenido default 'me' del original
    GRAPH_SCOPE = os.environ.get('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')
    # Añadir opcionales de SP si se necesitan, si no, se usará lookup/default en funciones SP
    SHAREPOINT_DEFAULT_SITE_ID = os.environ.get('SHAREPOINT_DEFAULT_SITE_ID')
    SHAREPOINT_DEFAULT_DRIVE_ID = os.environ.get('SHAREPOINT_DEFAULT_DRIVE_ID', 'Documents')
    logger.info("Variables de entorno cargadas correctamente.")
except ValueError as e:
    logger.critical(f"Error CRÍTICO de configuración inicial: {e}. La función no puede operar.")
    raise

# --- Constantes y Autenticación (Original) ---
BASE_URL = "https://graph.microsoft.com/v1.0"
HEADERS: Dict[str, Optional[str]] = {
    'Authorization': None,
    'Content-Type': 'application/json'
}
GRAPH_API_TIMEOUT = 45

def obtener_token() -> str:
    """Obtiene un token de acceso de aplicación usando credenciales de cliente."""
    # Esta es la función original de tu primer archivo
    logger.info("Obteniendo token de acceso de aplicación...")
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {'client_id': CLIENT_ID, 'scope': GRAPH_SCOPE, 'client_secret': CLIENT_SECRET, 'grant_type': 'client_credentials'}
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    response = None
    try:
        response = requests.post(url, data=data, headers=headers, timeout=GRAPH_API_TIMEOUT)
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
    # Esta es la función original de tu primer archivo
    try:
        token = obtener_token()
        HEADERS['Authorization'] = f'Bearer {token}'
        logger.info("Cabecera de autorización actualizada.")
    except Exception as e:
        logger.error(f"Falló la actualización de la cabecera: {e}")
        raise Exception(f"Fallo al actualizar la cabecera: {e}")

# --- FIN: Constantes y Autenticación ---


# --- INICIO: Funciones Auxiliares de Graph API (Originales + Añadidas SP/OD) ---

# ---- CORREO (Original) ----
def listar_correos(top: int = 10, skip: int = 0, folder: str = 'Inbox', select: Optional[List[str]] = None, filter_query: Optional[str] = None, order_by: Optional[str] = None, mailbox: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/mailFolders/{folder}/messages"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
    response: Optional[requests.Response] = None
    try: clean_params = {k:v for k, v in params.items() if v is not None}; logger.info(f"API Call: GET {url} Params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data: Dict[str, Any] = response.json(); logger.info(f"Listados {len(data.get('value',[]))} correos."); return data
    except Exception as e: logger.error(f"Error en listar_correos: {e}", exc_info=True); raise # Simplificado error handling original

def leer_correo(message_id: str, select: Optional[List[str]] = None, mailbox: Optional[str] = None) -> dict:
     _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"; params = {}; response: Optional[requests.Response] = None
     if select and isinstance(select, list): params['$select'] = ','.join(select)
     try: logger.info(f"API Call: GET {url} Params: {params}"); response = requests.get(url, headers=HEADERS, params=params or None, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Correo '{message_id}' leído."); return data
     except Exception as e: logger.error(f"Error en leer_correo: {e}", exc_info=True); raise

def enviar_correo(destinatario: Union[str, List[str]], asunto: str, mensaje: str, cc: Optional[Union[str, List[str]]] = None, bcc: Optional[Union[str, List[str]]] = None, attachments: Optional[List[dict]] = None, from_email: Optional[str] = None, is_draft: bool = False, mailbox: Optional[str] = None) -> dict:
     _actualizar_headers(); usuario = mailbox or MAILBOX
     if is_draft: url = f"{BASE_URL}/users/{usuario}/messages"; log_action = "Guardando borrador"
     else: url = f"{BASE_URL}/users/{usuario}/sendMail"; log_action = "Enviando correo"
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
         logger.info(f"API Call: POST {url} ({log_action})"); response = requests.post(url, headers=HEADERS, json=final_payload, timeout=GRAPH_API_TIMEOUT); response.raise_for_status()
         if not is_draft: logger.info(f"Correo enviado."); return {"status": "Enviado", "code": response.status_code}
         else: data = response.json(); message_id = data.get('id'); logger.info(f"Borrador guardado ID: {message_id}."); return {"status": "Borrador Guardado", "code": response.status_code, "id": message_id, "data": data}
     except Exception as e: logger.error(f"Error en {log_action}: {e}", exc_info=True); raise

def guardar_borrador(destinatario: Union[str, List[str]], asunto: str, mensaje: str, cc: Optional[Union[str, List[str]]] = None, bcc: Optional[Union[str, List[str]]] = None, attachments: Optional[List[dict]] = None, from_email: Optional[str] = None, mailbox: Optional[str] = None) -> dict:
     logger.info(f"Guardando borrador para '{mailbox or MAILBOX}'. Asunto: '{asunto}'"); return enviar_correo(destinatario, asunto, mensaje, cc, bcc, attachments, from_email, is_draft=True, mailbox=mailbox)

def enviar_borrador(message_id: str, mailbox: Optional[str] = None) -> dict:
     _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/send"; response: Optional[requests.Response] = None
     try: logger.info(f"API Call: POST {url}"); response = requests.post(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Borrador '{message_id}' enviado."); return {"status": "Borrador Enviado", "code": response.status_code}
     except Exception as e: logger.error(f"Error en enviar_borrador: {e}", exc_info=True); raise

def responder_correo(message_id: str, mensaje_respuesta: str, to_recipients: Optional[List[dict]] = None, reply_all: bool = False, mailbox: Optional[str] = None) -> dict:
     _actualizar_headers(); usuario = mailbox or MAILBOX; action = "replyAll" if reply_all else "reply"; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/{action}"
     payload: Dict[str, Any] = {"comment": mensaje_respuesta}
     if to_recipients: payload["message"] = { "toRecipients": to_recipients }; logger.info(f"Respondiendo con destinatarios custom.")
     response: Optional[requests.Response] = None
     try: logger.info(f"API Call: POST {url}"); response = requests.post(url, headers=HEADERS, json=payload, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Respuesta {'a todos ' if reply_all else ''}enviada correo '{message_id}'."); return {"status": "Respondido", "code": response.status_code}
     except Exception as e: logger.error(f"Error en responder_correo: {e}", exc_info=True); raise

def reenviar_correo(message_id: str, destinatarios: Union[str, List[str]], mensaje_reenvio: str = "FYI", mailbox: Optional[str] = None) -> dict:
     _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/forward"
     if isinstance(destinatarios, str): destinatarios = [destinatarios]
     to_recipients_list = [{"emailAddress": {"address": r}} for r in destinatarios if r and isinstance(r, str)]
     if not to_recipients_list: raise ValueError("Se requiere destinatario válido para reenviar.")
     payload = {"toRecipients": to_recipients_list, "comment": mensaje_reenvio}; response: Optional[requests.Response] = None
     try: logger.info(f"API Call: POST {url}"); response = requests.post(url, headers=HEADERS, json=payload, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Correo '{message_id}' reenviado."); return {"status": "Reenviado", "code": response.status_code}
     except Exception as e: logger.error(f"Error en reenviar_correo: {e}", exc_info=True); raise

def eliminar_correo(message_id: str, mailbox: Optional[str] = None) -> dict:
     _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
     response: Optional[requests.Response] = None
     try: logger.info(f"API Call: DELETE {url}"); response = requests.delete(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Correo '{message_id}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
     except Exception as e: logger.error(f"Error en eliminar_correo: {e}", exc_info=True); raise

# ---- CALENDARIO (Original) ----
def listar_eventos(top: int = 10, start_date: Optional[datetime] = None, end_date: Optional[datetime] = None, filter_query: Optional[str] = None, order_by: Optional[str] = None, select: Optional[List[str]] = None, use_calendar_view: bool = True, mailbox: Optional[str] = None) -> Dict[str, Any]: # Cambiado mailbox default a None como en Correo
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
     try: logger.info(f"API Call: GET {url} Params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Listados eventos."); return data
     except Exception as e: logger.error(f"Error en listar_eventos: {e}", exc_info=True); raise

def crear_evento(titulo: str, inicio: datetime, fin: datetime, asistentes: Optional[List[Dict[str, Any]]] = None, cuerpo: Optional[str] = None, es_reunion_online: bool = False, proveedor_reunion_online: str = "teamsForBusiness", recordatorio_minutos: Optional[int] = 15, ubicacion: Optional[str] = None, mostrar_como: str = "busy", mailbox: Optional[str] = None) -> Dict[str, Any]: # Cambiado mailbox default a None
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
     if proveedor_reunion_online: body["onlineMeetingProvider"] = proveedor_reunion_online # Mantener original: sin check es_reunion_online
     if recordatorio_minutos is not None and isinstance(recordatorio_minutos, int): body["isReminderOn"] = True; body["reminderMinutesBeforeStart"] = recordatorio_minutos
     else: body["isReminderOn"] = False
     if mostrar_como: body["showAs"] = mostrar_como # Mantener original: if en lugar de usar default
     response: Optional[requests.Response] = None
     try: logger.info(f"API Call: POST {url}"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Evento '{titulo}' creado."); return data
     except Exception as e: logger.error(f"Error en crear_evento: {e}", exc_info=True); raise

def actualizar_evento(evento_id: str, nuevos_valores: Dict[str, Any], mailbox: Optional[str] = None) -> Dict[str, Any]: # Cambiado mailbox default a None
     _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
     payload = nuevos_valores.copy()
     if 'start' in payload and isinstance(payload.get('start'), datetime): start_dt = payload['start']; start_dt = start_dt.replace(tzinfo=timezone.utc) if start_dt.tzinfo is None else start_dt; payload['start'] = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}; logging.info(f"Proc. fecha 'start': {payload['start']}") # Mantener UTC
     if 'end' in payload and isinstance(payload.get('end'), datetime): end_dt = payload['end']; end_dt = end_dt.replace(tzinfo=timezone.utc) if end_dt.tzinfo is None else end_dt; payload['end'] = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}; logging.info(f"Proc. fecha 'end': {payload['end']}") # Mantener UTC
     response: Optional[requests.Response] = None
     try:
         etag = payload.pop('@odata.etag', None); current_headers = HEADERS.copy()
         if etag: current_headers['If-Match'] = etag; logging.info(f"Usando ETag evento {evento_id}")
         logger.info(f"API Call: PATCH {url}"); response = requests.patch(url, headers=current_headers, json=payload, timeout=GRAPH_API_TIMEOUT); response.raise_for_status();
         logging.info(f"Evento '{evento_id}' actualizado.");
         if response.status_code == 204: return {"status": "Actualizado (No Content)", "id": evento_id}
         else: return response.json() # Asume 200 OK con body
     except Exception as e: logger.error(f"Error en actualizar_evento: {e}", exc_info=True); raise

def eliminar_evento(evento_id: str, mailbox: Optional[str] = None) -> Dict[str, Any]: # Cambiado mailbox default a None
     _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
     response: Optional[requests.Response] = None
     try: logger.info(f"API Call: DELETE {url}"); response = requests.delete(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Evento '{evento_id}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
     except Exception as e: logger.error(f"Error en eliminar_evento: {e}", exc_info=True); raise

# ---- TEAMS y OTROS (Original - Usa /me) ----
def listar_chats(top: int = 20, skip: int = 0, filter_query: Optional[str] = None, order_by: Optional[str] = None, expand: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/me/chats"; params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)} # Usa /me
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
    if expand is not None and isinstance(expand, str): params['$expand'] = expand
    clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try: logger.info(f"API Call: GET {url} Params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Listados {len(data.get('value',[]))} chats."); return data
    except Exception as e: logger.error(f"Error en listar_chats: {e}", exc_info=True); raise

def listar_equipos(top: int = 20, skip: int = 0, filter_query: Optional[str] = None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/me/joinedTeams"; params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)} # Usa /me
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try: logger.info(f"API Call: GET {url} Params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Listados {len(data.get('value',[]))} equipos."); return data
    except Exception as e: logger.error(f"Error en listar_equipos: {e}", exc_info=True); raise

def obtener_equipo(team_id: str, select: Optional[List[str]]=None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/teams/{team_id}"; params: Dict[str, Any] = {} # No usa /me
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try: logger.info(f"API Call: GET {url} Params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params or None, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Obtenido equipo ID: {team_id}."); return data
    except Exception as e: logger.error(f"Error en obtener_equipo: {e}", exc_info=True); raise


# ---- SHAREPOINT (Añadido) ----
# Usan _actualizar_headers() y HEADERS globales del script principal.

_cached_root_site_id_sp: Optional[str] = None # Renombrado para evitar colisión si había otro

def obtener_site_id_sp() -> str: # Renombrado, asume lookup raíz o usa variable entorno
    """Obtiene el ID del sitio raíz de SharePoint o usa el configurado."""
    global _cached_root_site_id_sp
    if SHAREPOINT_DEFAULT_SITE_ID:
        logger.info(f"Usando Site ID de config: {SHAREPOINT_DEFAULT_SITE_ID}")
        return SHAREPOINT_DEFAULT_SITE_ID
    if _cached_root_site_id_sp:
        logger.info(f"Usando Site ID raíz cacheado: {_cached_root_site_id_sp}")
        return _cached_root_site_id_sp

    _actualizar_headers() # Solo si no se usa cache o config
    url = f"{BASE_URL}/sites/root"
    try:
        logger.info(f"API Call: GET {url} (Obteniendo Site ID raíz SP)")
        response = requests.get(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        site_data = response.json()
        site_id = site_data.get('id')
        if not site_id: raise Exception("Respuesta de sitio inválida, falta 'id'.")
        logger.info(f"Site ID raíz SP obtenido: {site_id}")
        _cached_root_site_id_sp = site_id
        return site_id
    except Exception as e: logger.error(f"Error en obtener_site_id_sp: {e}", exc_info=True); raise

# -- SHAREPOINT - Listas (Añadido) --
def sp_crear_lista(nombre_lista: str) -> dict:
    _actualizar_headers(); site_id = obtener_site_id_sp(); url = f"{BASE_URL}/sites/{site_id}/lists"
    body = {"displayName": nombre_lista, "columns": [{"name": "Clave", "text": {}}, {"name": "Valor", "text": {}}], "list": {"template": "genericList"}}
    try: logger.info(f"API Call: POST {url}"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info(f"Lista '{nombre_lista}' creada."); return data
    except Exception as e: logger.error(f"Error en sp_crear_lista: {e}", exc_info=True); raise

def sp_listar_listas() -> dict:
     _actualizar_headers(); site_id = obtener_site_id_sp(); url = f"{BASE_URL}/sites/{site_id}/lists"
     try: logger.info(f"API Call: GET {url}"); response = requests.get(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info("Listando listas."); return data
     except Exception as e: logger.error(f"Error en sp_listar_listas: {e}", exc_info=True); raise

def sp_agregar_elemento_lista(nombre_lista: str, clave: str, valor: str) -> dict: # Modificado params para coincidir con definicion original SP
     _actualizar_headers(); site_id = obtener_site_id_sp(); url = f"{BASE_URL}/sites/{site_id}/lists/{nombre_lista}/items"
     body = {"fields": {"Clave": clave, "Valor": valor}} # Usando Clave/Valor del ejemplo SP
     try: logger.info(f"API Call: POST {url}"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info(f"Elemento agregado a '{nombre_lista}'."); return data
     except Exception as e: logger.error(f"Error en sp_agregar_elemento_lista: {e}", exc_info=True); raise

def sp_listar_elementos_lista(nombre_lista: str, expand_fields: bool = True) -> dict: # Como en archivo SP
     _actualizar_headers(); site_id = obtener_site_id_sp(); url = f"{BASE_URL}/sites/{site_id}/lists/{nombre_lista}/items";
     if expand_fields: url += "?expand=fields"
     all_items = []; current_url: Optional[str] = url
     try:
         while current_url:
             logger.info(f"API Call: GET {current_url} (Listando elems SP lista '{nombre_lista}')")
             response = requests.get(current_url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); page_items = data.get('value', []); all_items.extend(page_items)
             current_url = data.get('@odata.nextLink')
             if current_url: _actualizar_headers() # Refrescar token para paginación
         logger.info(f"Total elems. lista SP '{nombre_lista}': {len(all_items)}"); return {'value': all_items}
     except Exception as e: logger.error(f"Error en sp_listar_elementos_lista: {e}", exc_info=True); raise

def sp_actualizar_elemento_lista(nombre_lista: str, item_id: str, nuevos_valores: dict) -> dict: # Como en archivo SP
     _actualizar_headers(); site_id = obtener_site_id_sp(); url = f"{BASE_URL}/sites/{site_id}/lists/{nombre_lista}/items/{item_id}/fields"
     try: logger.info(f"API Call: PATCH {url}"); response = requests.patch(url, headers=HEADERS, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info(f"Elem. '{item_id}' actualizado lista '{nombre_lista}'."); return data
     except Exception as e: logger.error(f"Error en sp_actualizar_elemento_lista: {e}", exc_info=True); raise

def sp_eliminar_elemento_lista(nombre_lista: str, item_id: str) -> dict: # Como en archivo SP
     _actualizar_headers(); site_id = obtener_site_id_sp(); url = f"{BASE_URL}/sites/{site_id}/lists/{nombre_lista}/items/{item_id}"
     try: logger.info(f"API Call: DELETE {url}"); response = requests.delete(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Elem. '{item_id}' eliminado lista '{nombre_lista}'."); return {"status": "Eliminado", "code": response.status_code}
     except Exception as e: logger.error(f"Error en sp_eliminar_elemento_lista: {e}", exc_info=True); raise

# -- SHAREPOINT - Documentos (Añadido) --
# Usan SHAREPOINT_DEFAULT_DRIVE_ID si 'biblioteca' no se especifica
def sp_listar_documentos_biblioteca(biblioteca: Optional[str] = None) -> dict: # biblioteca opcional
     _actualizar_headers(); site_id = obtener_site_id_sp(); drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID; url = f"{BASE_URL}/sites/{site_id}/drives/{drive}/root/children"
     all_files = []; current_url: Optional[str] = url
     try:
         while current_url:
             logger.info(f"API Call: GET {current_url} (Listando docs SP biblioteca '{drive}')")
             response = requests.get(current_url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); page_items = data.get('value', []); all_files.extend(page_items)
             current_url = data.get('@odata.nextLink')
             if current_url: _actualizar_headers()
         logger.info(f"Total docs SP biblioteca '{drive}': {len(all_files)}"); return {'value': all_files}
     except Exception as e: logger.error(f"Error en sp_listar_documentos_biblioteca: {e}", exc_info=True); raise

def sp_subir_documento(nombre_archivo: str, contenido_base64: Union[str, bytes], biblioteca: Optional[str] = None) -> dict: # biblioteca opcional
     _actualizar_headers(); site_id = obtener_site_id_sp(); drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID; url = f"{BASE_URL}/sites/{site_id}/drives/{drive}/root:/{nombre_archivo}:/content"
     if isinstance(contenido_base64, str): contenido_bytes = contenido_base64.encode('utf-8') # Asume UTF-8 si es string
     else: contenido_bytes = contenido_base64 # Ya son bytes
     upload_headers = HEADERS.copy(); upload_headers['Content-Type'] = 'application/octet-stream' # Asumir binario
     try:
        logger.info(f"API Call: PUT {url} (Subiendo doc SP '{nombre_archivo}' a biblioteca '{drive}')")
        if len(contenido_bytes) > 4*1024*1024: logger.warning(f"Archivo SP '{nombre_archivo}' > 4MB.")
        response = requests.put(url, headers=upload_headers, data=contenido_bytes, timeout=GRAPH_API_TIMEOUT * 3); response.raise_for_status(); data=response.json(); logger.info(f"Doc SP '{nombre_archivo}' subido."); return data
     except Exception as e: logger.error(f"Error en sp_subir_documento: {e}", exc_info=True); raise

# Renombrado para claridad, usa la lógica del segundo eliminar_archivo del código SP
def sp_eliminar_archivo_biblioteca(nombre_archivo: str, biblioteca: Optional[str] = None) -> dict: # biblioteca opcional
     _actualizar_headers(); site_id = obtener_site_id_sp(); drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID; url = f"{BASE_URL}/sites/{site_id}/drives/{drive}/root:/{nombre_archivo}"
     try: logger.info(f"API Call: DELETE {url} (Eliminando archivo SP '{nombre_archivo}' de biblioteca '{drive}')"); response = requests.delete(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Archivo SP '{nombre_archivo}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
     except Exception as e: logger.error(f"Error en sp_eliminar_archivo_biblioteca: {e}", exc_info=True); raise

def sp_crear_carpeta_biblioteca(nombre_carpeta: str, biblioteca: Optional[str] = None) -> dict: # biblioteca opcional
    _actualizar_headers(); site_id = obtener_site_id_sp(); drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID; url = f"{BASE_URL}/sites/{site_id}/drives/{drive}/root/children"
    body = {"name": nombre_carpeta, "folder": {}}
    try: logger.info(f"API Call: POST {url} (Creando carpeta SP '{nombre_carpeta}' en biblioteca '{drive}')"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info(f"Carpeta SP '{nombre_carpeta}' creada."); return data
    except Exception as e: logger.error(f"Error en sp_crear_carpeta_biblioteca: {e}", exc_info=True); raise

def sp_mover_archivo(nombre_archivo: str, nueva_ubicacion: str, biblioteca: Optional[str] = None) -> dict: # biblioteca opcional
     _actualizar_headers(); site_id = obtener_site_id_sp(); drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID; url = f"{BASE_URL}/sites/{site_id}/drives/{drive}/root:/{nombre_archivo}"
     # Necesita ID del drive destino para path, obtenerlo primero
     try: drive_resp = requests.get(f"{BASE_URL}/sites/{site_id}/drives/{drive}", headers=HEADERS, params={'$select':'id'}, timeout=GRAPH_API_TIMEOUT); drive_resp.raise_for_status(); actual_drive_id = drive_resp.json().get('id'); assert actual_drive_id
     except Exception as e: logger.error(f"Error obteniendo ID drive '{drive}' para mover: {e}"); raise Exception(f"Error obteniendo ID drive destino: {e}")
     parent_path = f"/drives/{actual_drive_id}/root:{nueva_ubicacion.strip()}" if nueva_ubicacion != '/' else f"/drives/{actual_drive_id}/root"
     body = {"parentReference": {"path": parent_path}}
     try: logger.info(f"API Call: PATCH {url} (Moviendo SP '{nombre_archivo}' a '{nueva_ubicacion}' en biblioteca '{drive}')"); response = requests.patch(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info(f"Archivo SP '{nombre_archivo}' movido."); return data
     except Exception as e: logger.error(f"Error en sp_mover_archivo: {e}", exc_info=True); raise

def sp_copiar_archivo(nombre_archivo: str, nueva_ubicacion: str, biblioteca: Optional[str] = None) -> dict: # biblioteca opcional
     _actualizar_headers(); site_id = obtener_site_id_sp(); drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID; url = f"{BASE_URL}/sites/{site_id}/drives/{drive}/root:/{nombre_archivo}:/copy" # Endpoint de copia
     # Necesita ID del drive destino para parentReference
     try: drive_resp = requests.get(f"{BASE_URL}/sites/{site_id}/drives/{drive}", headers=HEADERS, params={'$select':'id'}, timeout=GRAPH_API_TIMEOUT); drive_resp.raise_for_status(); actual_drive_id = drive_resp.json().get('id'); assert actual_drive_id
     except Exception as e: logger.error(f"Error obteniendo ID drive '{drive}' para copiar: {e}"); raise Exception(f"Error obteniendo ID drive destino: {e}")
     parent_path = f"/drive/root:{nueva_ubicacion.strip()}" if nueva_ubicacion != '/' else f"/drive/root"
     body = {"parentReference": {"driveId": actual_drive_id, "path": parent_path}, "name": nombre_archivo } # Mantener nombre por defecto
     try: logger.info(f"API Call: POST {url} (Copiando SP '{nombre_archivo}' a '{nueva_ubicacion}' en biblioteca '{drive}')"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); monitor_url = response.headers.get('Location'); logger.info(f"Copia SP '{nombre_archivo}' iniciada. Monitor: {monitor_url}"); return {"status": "Copia Iniciada", "code": response.status_code, "monitorUrl": monitor_url}
     except Exception as e: logger.error(f"Error en sp_copiar_archivo: {e}", exc_info=True); raise

def sp_obtener_metadatos_archivo(nombre_archivo: str, biblioteca: Optional[str] = None) -> dict: # biblioteca opcional
     _actualizar_headers(); site_id = obtener_site_id_sp(); drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID; url = f"{BASE_URL}/sites/{site_id}/drives/{drive}/root:/{nombre_archivo}"
     try: logger.info(f"API Call: GET {url} (Obteniendo metadatos SP '{nombre_archivo}' de biblioteca '{drive}')"); response = requests.get(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info(f"Metadatos SP '{nombre_archivo}' obtenidos."); return data
     except Exception as e: logger.error(f"Error en sp_obtener_metadatos_archivo: {e}", exc_info=True); raise

# Las siguientes funciones SP estaban en el archivo SP pero no se mapearon antes, añadiendo por completitud
def sp_actualizar_metadatos_archivo(nombre_archivo: str, nuevos_valores: dict, biblioteca: Optional[str] = None) -> dict:
     _actualizar_headers(); site_id = obtener_site_id_sp(); drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID; url = f"{BASE_URL}/sites/{site_id}/drives/{drive}/root:/{nombre_archivo}"
     try: logger.info(f"API Call: PATCH {url} (Actualizando metadatos SP '{nombre_archivo}' en biblioteca '{drive}')"); response = requests.patch(url, headers=HEADERS, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info(f"Metadatos SP '{nombre_archivo}' actualizados."); return data
     except Exception as e: logger.error(f"Error en sp_actualizar_metadatos_archivo: {e}", exc_info=True); raise

def sp_obtener_contenido_archivo(nombre_archivo: str, biblioteca: Optional[str] = None) -> bytes:
     _actualizar_headers(); site_id = obtener_site_id_sp(); drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID; url = f"{BASE_URL}/sites/{site_id}/drives/{drive}/root:/{nombre_archivo}/content"
     try: logger.info(f"API Call: GET {url} (Obteniendo contenido SP '{nombre_archivo}' de biblioteca '{drive}')"); response = requests.get(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT * 2); response.raise_for_status(); logger.info(f"Contenido SP '{nombre_archivo}' obtenido."); return response.content
     except Exception as e: logger.error(f"Error en sp_obtener_contenido_archivo: {e}", exc_info=True); raise

def sp_actualizar_contenido_archivo(nombre_archivo: str, nuevo_contenido: bytes, biblioteca: Optional[str] = None) -> dict:
     _actualizar_headers(); site_id = obtener_site_id_sp(); drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID; url = f"{BASE_URL}/sites/{site_id}/drives/{drive}/root:/{nombre_archivo}/content"
     upload_headers = HEADERS.copy(); upload_headers['Content-Type'] = 'application/octet-stream'
     try: logger.info(f"API Call: PUT {url} (Actualizando contenido SP '{nombre_archivo}' en biblioteca '{drive}')"); response = requests.put(url, headers=upload_headers, data=nuevo_contenido, timeout=GRAPH_API_TIMEOUT * 3); response.raise_for_status(); data=response.json(); logger.info(f"Contenido SP '{nombre_archivo}' actualizado."); return data
     except Exception as e: logger.error(f"Error en sp_actualizar_contenido_archivo: {e}", exc_info=True); raise

def sp_crear_enlace_compartido_archivo(nombre_archivo: str, tipo_enlace: str = "view", alcance: str = "anonymous", biblioteca: Optional[str] = None) -> dict:
     _actualizar_headers(); site_id = obtener_site_id_sp(); drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID; url = f"{BASE_URL}/sites/{site_id}/drives/{drive}/root:/{nombre_archivo}/createLink"
     body = {"type": tipo_enlace, "scope": alcance}
     try: logger.info(f"API Call: POST {url} (Creando enlace SP '{nombre_archivo}' en biblioteca '{drive}')"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info(f"Enlace SP creado para '{nombre_archivo}'."); return data
     except Exception as e: logger.error(f"Error en sp_crear_enlace_compartido_archivo: {e}", exc_info=True); raise


# ---- ONEDRIVE (Añadido - usa /me) ----
# Usan _actualizar_headers() y HEADERS globales. Usan /me/drive/...

def od_listar_archivos(ruta: str = "/") -> dict:
     _actualizar_headers(); url = f"{BASE_URL}/me/drive/root:{ruta}:/children" # Usa /me
     all_items = []; current_url: Optional[str] = url
     try:
         while current_url:
             logger.info(f"API Call: GET {current_url} (Listando OD /me ruta '{ruta}')")
             response = requests.get(current_url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); page_items = data.get('value', []); all_items.extend(page_items)
             current_url = data.get('@odata.nextLink')
             if current_url: _actualizar_headers()
         logger.info(f"Total items OD /me en '{ruta}': {len(all_items)}"); return {'value': all_items}
     except Exception as e: logger.error(f"Error en od_listar_archivos: {e}", exc_info=True); raise

def od_subir_archivo(nombre_archivo: str, contenido: Union[str, bytes], ruta: str = "/") -> dict: # Nombre param 'contenido' como en archivo OD
     _actualizar_headers(); url = f"{BASE_URL}/me/drive/root:{ruta}/{nombre_archivo}:/content" # Usa /me
     if isinstance(contenido, str): contenido_bytes = contenido.encode('utf-8')
     else: contenido_bytes = contenido
     # Corregido: Pasar HEADERS completos, no solo Auth
     upload_headers = HEADERS.copy(); upload_headers['Content-Type'] = 'application/octet-stream'
     try:
         logger.info(f"API Call: PUT {url} (Subiendo OD /me '{nombre_archivo}' a ruta '{ruta}')")
         if len(contenido_bytes) > 4*1024*1024: logger.warning(f"Archivo OD '{nombre_archivo}' > 4MB.")
         response = requests.put(url, headers=upload_headers, data=contenido_bytes, timeout=GRAPH_API_TIMEOUT * 3); response.raise_for_status(); data=response.json(); logger.info(f"Archivo OD '{nombre_archivo}' subido."); return data
     except Exception as e: logger.error(f"Error en od_subir_archivo: {e}", exc_info=True); raise

def od_descargar_archivo(nombre_archivo: str, ruta: str = "/") -> bytes:
     _actualizar_headers(); url = f"{BASE_URL}/me/drive/root:{ruta}/{nombre_archivo}:/content" # Usa /me
     try: logger.info(f"API Call: GET {url} (Descargando OD /me '{nombre_archivo}' de ruta '{ruta}')"); response = requests.get(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT * 2); response.raise_for_status(); logger.info(f"Archivo OD '{nombre_archivo}' descargado."); return response.content
     except Exception as e: logger.error(f"Error en od_descargar_archivo: {e}", exc_info=True); raise

def od_eliminar_archivo(nombre_archivo: str, ruta: str = "/") -> dict: # Renombrado para consistencia
     _actualizar_headers(); url = f"{BASE_URL}/me/drive/root:{ruta}/{nombre_archivo}" # Usa /me
     try: logger.info(f"API Call: DELETE {url} (Eliminando OD /me '{nombre_archivo}' de ruta '{ruta}')"); response = requests.delete(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Archivo OD '{nombre_archivo}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
     except Exception as e: logger.error(f"Error en od_eliminar_archivo: {e}", exc_info=True); raise

def od_crear_carpeta(nombre_carpeta: str, ruta: str = "/") -> dict:
     _actualizar_headers(); url = f"{BASE_URL}/me/drive/root:{ruta}:/children" # Usa /me
     body = {"name": nombre_carpeta, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"}
     try: logger.info(f"API Call: POST {url} (Creando OD /me carpeta '{nombre_carpeta}' en ruta '{ruta}')"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info(f"Carpeta OD '{nombre_carpeta}' creada."); return data
     except Exception as e: logger.error(f"Error en od_crear_carpeta: {e}", exc_info=True); raise

def od_mover_archivo(nombre_archivo: str, ruta_origen: str = "/", ruta_destino: str = "/NuevaCarpeta") -> dict: # Nombre param como en OD file
     _actualizar_headers(); url = f"{BASE_URL}/me/drive/root:{ruta_origen}/{nombre_archivo}" # Usa /me
     parent_path = f"/drive/root:{ruta_destino.strip()}" if ruta_destino != '/' else "/drive/root" # Path relativo a /drive/root
     body = {"parentReference": {"path": parent_path}, "name": nombre_archivo}
     try: logger.info(f"API Call: PATCH {url} (Moviendo OD /me '{nombre_archivo}' de '{ruta_origen}' a '{ruta_destino}')"); response = requests.patch(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info(f"Archivo/Carpeta OD '{nombre_archivo}' movido."); return data
     except Exception as e: logger.error(f"Error en od_mover_archivo: {e}", exc_info=True); raise

def od_copiar_archivo(nombre_archivo: str, ruta_origen: str = "/", ruta_destino: str = "/Copias") -> dict: # Nombre param como en OD file
     _actualizar_headers()
     # Necesita ID del drive /me
     try: drive_resp = requests.get(f"{BASE_URL}/me/drive", headers=HEADERS, params={'$select':'id'}, timeout=GRAPH_API_TIMEOUT); drive_resp.raise_for_status(); actual_drive_id = drive_resp.json().get('id'); assert actual_drive_id
     except Exception as e: logger.error(f"Error obteniendo ID drive /me para copiar: {e}"); raise Exception(f"Error obteniendo ID drive /me: {e}")
     url = f"{BASE_URL}/me/drive/root:{ruta_origen}/{nombre_archivo}/copy" # Usa /me
     parent_path = f"/drive/root:{ruta_destino.strip()}" if ruta_destino != '/' else "/drive/root"
     body = {"parentReference": {"driveId": actual_drive_id, "path": parent_path}, "name": f"Copia_{nombre_archivo}"} # Como en OD file
     try: logger.info(f"API Call: POST {url} (Copiando OD /me '{nombre_archivo}' de '{ruta_origen}' a '{ruta_destino}')"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); monitor_url = response.headers.get('Location'); logger.info(f"Copia OD '{nombre_archivo}' iniciada. Monitor: {monitor_url}"); return {"status": "Copia Iniciada", "code": response.status_code, "monitorUrl": monitor_url}
     except Exception as e: logger.error(f"Error en od_copiar_archivo: {e}", exc_info=True); raise

def od_obtener_metadatos_archivo(nombre_archivo: str, ruta: str = "/") -> dict:
     _actualizar_headers(); url = f"{BASE_URL}/me/drive/root:{ruta}/{nombre_archivo}" # Usa /me
     try: logger.info(f"API Call: GET {url} (Obteniendo metadatos OD /me '{nombre_archivo}' ruta '{ruta}')"); response = requests.get(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info(f"Metadatos OD '{nombre_archivo}' obtenidos."); return data
     except Exception as e: logger.error(f"Error en od_obtener_metadatos_archivo: {e}", exc_info=True); raise

def od_actualizar_metadatos_archivo(nombre_archivo: str, nuevos_valores: dict, ruta:str = "/")->dict:
     _actualizar_headers(); url = f"{BASE_URL}/me/drive/root:{ruta}/{nombre_archivo}" # Usa /me
     try: logger.info(f"API Call: PATCH {url} (Actualizando metadatos OD /me '{nombre_archivo}' ruta '{ruta}')"); response = requests.patch(url, headers=HEADERS, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data=response.json(); logger.info(f"Metadatos OD '{nombre_archivo}' actualizados."); return data
     except Exception as e: logger.error(f"Error en od_actualizar_metadatos_archivo: {e}", exc_info=True); raise


# --- FIN: Funciones Auxiliares de Graph API ---


# --- INICIO: Función Principal de Azure Functions (Entry Point - Original + Mapeo Añadido) ---

# Mapeo de nombres de acción a las funciones DEFINIDAS ARRIBA
acciones_disponibles: Dict[str, Callable[..., Any]] = {
    # Originales
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
    "listar_chats": listar_chats,       # Usa /me como en original
    "listar_equipos": listar_equipos,   # Usa /me como en original
    "obtener_equipo": obtener_equipo,   # Usa /teams/{id} como en original
    # SharePoint - Listas (Añadidas)
    "sp_crear_lista": sp_crear_lista,
    "sp_listar_listas": sp_listar_listas,
    "sp_agregar_elemento_lista": sp_agregar_elemento_lista,
    "sp_listar_elementos_lista": sp_listar_elementos_lista,
    "sp_actualizar_elemento_lista": sp_actualizar_elemento_lista,
    "sp_eliminar_elemento_lista": sp_eliminar_elemento_lista,
    # SharePoint - Documentos/Bibliotecas (Añadidas)
    "sp_listar_documentos_biblioteca": sp_listar_documentos_biblioteca,
    "sp_subir_documento": sp_subir_documento,
    "sp_eliminar_archivo_biblioteca": sp_eliminar_archivo_biblioteca, # Nombre unificado
    "sp_crear_carpeta_biblioteca": sp_crear_carpeta_biblioteca,
    "sp_mover_archivo": sp_mover_archivo,
    "sp_copiar_archivo": sp_copiar_archivo,
    "sp_obtener_metadatos_archivo": sp_obtener_metadatos_archivo,
    "sp_actualizar_metadatos_archivo": sp_actualizar_metadatos_archivo,
    "sp_obtener_contenido_archivo": sp_obtener_contenido_archivo,
    "sp_actualizar_contenido_archivo": sp_actualizar_contenido_archivo,
    "sp_crear_enlace_compartido_archivo": sp_crear_enlace_compartido_archivo,
    # OneDrive (Añadidas - usan /me)
    "od_listar_archivos": od_listar_archivos,
    "od_subir_archivo": od_subir_archivo,
    "od_descargar_archivo": od_descargar_archivo,
    "od_eliminar_archivo": od_eliminar_archivo, # Nombre unificado
    "od_crear_carpeta": od_crear_carpeta,
    "od_mover_archivo": od_mover_archivo,
    "od_copiar_archivo": od_copiar_archivo,
    "od_obtener_metadatos_archivo": od_obtener_metadatos_archivo,
    "od_actualizar_metadatos_archivo": od_actualizar_metadatos_archivo,
}

# Verificar mapeo (Original)
for accion_check, func_ref_check in acciones_disponibles.items():
    if not callable(func_ref_check):
        logger.error(f"Config Error: La función para la acción '{accion_check}' no es válida o no está definida.")

def main(req: func.HttpRequest) -> func.HttpResponse:
    """Punto de entrada principal. Maneja la solicitud HTTP, llama a la acción apropiada y devuelve la respuesta."""
    # Esta es la función main() original, con la lógica de parseo y ejecución que tenías.
    # Se ha mantenido intacta excepto por la adición de más conversiones/validaciones de params si fuesen necesarias.
    logging.info(f'Python HTTP trigger function procesando solicitud. Method: {req.method}, URL: {req.url}')
    invocation_id = req.headers.get('X-Azure-Functions-InvocationId', 'N/A')
    logging.info(f"Invocation ID: {invocation_id}")

    accion: Optional[str] = None
    parametros: Dict[str, Any] = {}
    funcion_a_ejecutar: Optional[Callable] = None

    # --- INICIO: Bloque Try-Except General (Original) ---
    try:
        # --- Leer accion/parametros (Original) ---
        req_body: Optional[Dict[str, Any]] = None
        # Adaptado ligeramente para manejar form-data para subidas, como en versión anterior
        content_type = req.headers.get('Content-Type', '').lower()
        if req.method in ('POST', 'PUT', 'PATCH'):
            if 'application/json' in content_type:
                try:
                    req_body = req.get_json(); assert isinstance(req_body, dict)
                    accion = req_body.get('accion'); params_input = req_body.get('parametros')
                    if isinstance(params_input, dict): parametros = params_input
                    elif params_input is not None: logger.warning(f"Invocation {invocation_id}: 'parametros' no es dict"); parametros = {}
                    else: parametros = {}
                except (ValueError, AssertionError): logger.warning(f'Invocation {invocation_id}: Cuerpo no es JSON válido o no objeto.'); return func.HttpResponse("Cuerpo JSON inválido o no es objeto.", status_code=400)
            elif 'multipart/form-data' in content_type:
                 logger.info(f"Invocation {invocation_id}: Procesando multipart/form-data."); parametros = {}
                 accion = req.form.get('accion')
                 for key, value in req.form.items():
                     if key not in ['accion', 'file']: parametros[key] = value
                 file = req.files.get('file')
                 if file: logger.info(f"Archivo '{file.filename}' en form-data."); parametros['contenido_bytes'] = file.read(); parametros['nombre_archivo_original'] = file.filename # Necesario para funciones de subida
                 else: logger.info("No archivo ('file') en form-data.")
            else: # Fallback a query params si no es JSON o form-data conocido
                 accion = req.params.get('accion'); parametros = dict(req.params)
        else: # GET y otros
             accion = req.params.get('accion'); parametros = dict(req.params)
        if 'accion' in parametros: del parametros['accion'] # Limpiar

        # --- Validar acción (Original) ---
        if not accion or not isinstance(accion, str): logger.warning(f"Invocation {invocation_id}: Clave 'accion' faltante o no es string."); return func.HttpResponse("Falta 'accion' (string).", status_code=400)
        log_params = {k: f"<bytes len={len(v)}>" if isinstance(v, bytes) else v for k, v in parametros.items()} # Log seguro
        logger.info(f"Invocation {invocation_id}: Acción solicitada: '{accion}'. Parámetros iniciales: {log_params}")

        # --- Buscar y ejecutar la función (Original) ---
        if accion in acciones_disponibles:
            funcion_a_ejecutar = acciones_disponibles[accion]
            logger.info(f"Invocation {invocation_id}: Mapeado a función: {funcion_a_ejecutar.__name__}")

            # --- Validar/Convertir parámetros ANTES de llamar (Original + Añadidos si es necesario) ---
            params_procesados: Dict[str, Any] = {}
            try:
                params_procesados = parametros.copy()
                # Conversiones originales
                if accion in ["listar_correos", "listar_eventos", "listar_chats", "listar_equipos"]: # Añadir nuevas listas si usan top/skip
                    if 'top' in params_procesados and params_procesados['top'] is not None: params_procesados['top'] = int(params_procesados['top'])
                    if 'skip' in params_procesados and params_procesados['skip'] is not None: params_procesados['skip'] = int(params_procesados['skip'])
                elif accion in ["crear_evento", "actualizar_evento"]:
                     for date_key in ['inicio', 'fin']:
                         if date_key in params_procesados and params_procesados[date_key] is not None:
                             date_val = params_procesados[date_key]
                             if isinstance(date_val, str): params_procesados[date_key] = datetime.fromisoformat(date_val.replace('Z', '+00:00'))
                             elif not isinstance(date_val, datetime): raise ValueError(f"Tipo inválido para '{date_key}'.")
                # Añadir conversiones/validaciones para SP/OD si hace falta
                # Ejemplo: asegurar que 'nuevos_valores' sea dict, etc.
                if accion in ["sp_subir_documento", "od_subir_archivo"]:
                     if 'contenido_bytes' not in params_procesados: raise ValueError(f"Acción '{accion}' requiere 'contenido_bytes'.")
                     # Usar nombre original si no se pasa nombre_archivo y se subió por form
                     if 'nombre_archivo' not in params_procesados and 'nombre_archivo_destino' not in params_procesados and 'nombre_archivo_original' in params_procesados:
                         if accion == "sp_subir_documento": params_procesados['nombre_archivo'] = params_procesados['nombre_archivo_original']
                         if accion == "od_subir_archivo": params_procesados['nombre_archivo'] = params_procesados['nombre_archivo_original']


            except (ValueError, TypeError, KeyError) as conv_err:
                logger.error(f"Invocation {invocation_id}: Error en parámetros para '{accion}': {conv_err}. Recibido: {log_params}", exc_info=True)
                return func.HttpResponse(f"Parámetros inválidos para '{accion}': {conv_err}", status_code=400)

            # --- Llamar a la función auxiliar (Original) ---
            logger.info(f"Invocation {invocation_id}: Ejecutando {funcion_a_ejecutar.__name__}...")
            resultado: Any = None
            try:
                # Pasar parámetros desempaquetados
                resultado = funcion_a_ejecutar(**params_procesados)
                logger.info(f"Invocation {invocation_id}: Ejecución de '{accion}' completada.")
            except TypeError as type_err: # Error común si faltan args o hay extras
                 logger.error(f"Invocation {invocation_id}: Error de argumento al llamar a {funcion_a_ejecutar.__name__}: {type_err}. Params pasados: {params_procesados}", exc_info=True)
                 return func.HttpResponse(f"Error en argumentos para acción '{accion}': {type_err}", status_code=400)
            except Exception as exec_err:
                logger.exception(f"Invocation {invocation_id}: Error durante ejecución acción '{accion}': {exec_err}")
                return func.HttpResponse(f"Error interno al ejecutar '{accion}': {exec_err}", status_code=500)

            # --- Devolver resultado (Original + manejo bytes) ---
            if isinstance(resultado, bytes): # Añadido para descargas
                 logger.info(f"Invocation {invocation_id}: Devolviendo contenido binario (bytes).")
                 filename = os.path.basename(parametros.get('nombre_archivo') or parametros.get('ruta_archivo') or 'downloaded_file')
                 return func.HttpResponse(resultado, mimetype="application/octet-stream", headers={'Content-Disposition': f'attachment; filename="{filename}"'}, status_code=200)
            elif isinstance(resultado, (dict, list)): # Manejo original JSON
                try:
                    return func.HttpResponse(json.dumps(resultado, default=str), mimetype="application/json", status_code=200)
                except TypeError as serialize_err:
                    logger.error(f"Invocation {invocation_id}: Error al serializar resultado JSON para '{accion}': {serialize_err}.", exc_info=True)
                    return func.HttpResponse(f"Error interno: Respuesta no serializable.", status_code=500)
            else: # Otros tipos como strings o status dicts
                 return func.HttpResponse(str(resultado), mimetype="text/plain", status_code=200)


        else: # Acción no encontrada (Original)
            logger.warning(f"Invocation {invocation_id}: Acción '{accion}' no reconocida."); acciones_validas = list(acciones_disponibles.keys());
            return func.HttpResponse(f"Acción '{accion}' no reconocida. Válidas: {acciones_validas}", status_code=400)

    # --- FIN: Bloque Try-Except General (Original) ---
    except Exception as e:
        func_name = getattr(funcion_a_ejecutar, '__name__', 'N/A')
        logger.exception(f"Invocation {invocation_id}: Error GENERAL INESPERADO en main() procesando acción '{accion or 'desconocida'}' (Función: {func_name}): {e}")
        return func.HttpResponse("Error interno del servidor. Revise los logs.", status_code=500)

# --- FIN: Función Principal ---
