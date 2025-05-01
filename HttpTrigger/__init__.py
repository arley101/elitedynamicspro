# HttpTrigger/__init__.py (Versión Corregida - Restaurado /me para OD/Teams)

import json
import logging
import requests
import azure.functions as func
from typing import Dict, Any, Callable, List, Optional, Union, Mapping, Sequence
from datetime import datetime, timezone
import os
import io # Necesario para subir contenido binario si se recibe como bytes

# --- Configuración de Logging ---
logger = logging.getLogger("azure.functions")
logger.setLevel(logging.INFO) # O logging.DEBUG para más detalle

# --- Variables de Entorno y Configuración ---
# CLIENT_ID, TENANT_ID, CLIENT_SECRET: Necesarias para obtener_token (incluso si algunas llamadas usan /me con token delegado obtenido externamente, el mecanismo de token está aquí).
# MAILBOX: Usado por defecto para funciones de Correo y Calendario (/users/{MAILBOX}/...).
# GRAPH_SCOPE: Alcance para el token (ej: 'https://graph.microsoft.com/.default' para app, o scopes delegados si el token se obtiene de otra forma).
# SHAREPOINT_DEFAULT_SITE_ID: (Opcional) ID del sitio raíz de SharePoint por defecto.
# SHAREPOINT_DEFAULT_DRIVE_ID: (Opcional) ID o nombre de la biblioteca de documentos por defecto (ej: 'Documents').

def get_config_or_raise(key: str, default: Optional[str] = None) -> str:
    """Obtiene un valor de configuración de las variables de entorno o lanza un error."""
    value = os.environ.get(key, default)
    if value is None:
        # Solo loguear error aquí, la validación crítica se hará en main() o al inicio
        logger.warning(f"CONFIG WARNING: Variable de entorno no encontrada: {key}. Se usará default si existe, o podría fallar.")
        # Permitir que continúe y falle en la operación si es realmente requerida
        if default is None: # Si no hay default, es potencialmente un problema
             raise ValueError(f"Configuración esencial faltante y sin default: {key}")
    return value if value is not None else default # Devolver default si value es None


try:
    # Cargar configuración. Algunas pueden ser opcionales dependiendo del flujo de auth real.
    CLIENT_ID = os.environ.get('CLIENT_ID')
    TENANT_ID = os.environ.get('TENANT_ID')
    CLIENT_SECRET = os.environ.get('CLIENT_SECRET') # Puede no ser usado si se inyecta un token delegado
    MAILBOX = os.environ.get('MAILBOX', 'me') # Default a 'me' si no se especifica; usado para Mail/Calendar
    GRAPH_SCOPE = os.environ.get('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')
    SHAREPOINT_DEFAULT_SITE_ID = os.environ.get('SHAREPOINT_DEFAULT_SITE_ID')
    SHAREPOINT_DEFAULT_DRIVE_ID = os.environ.get('SHAREPOINT_DEFAULT_DRIVE_ID', 'Documents') # Default a 'Documents'
    logger.info("Variables de entorno cargadas (validación pendiente de uso).")
except Exception as e:
     # Error inesperado al leer env vars
     logger.critical(f"Error CRÍTICO leyendo configuración inicial: {e}. La función podría no operar.")
     raise

# --- Constantes y Autenticación ---
BASE_URL = "https://graph.microsoft.com/v1.0"
# Headers globales - Se actualizarán con el token obtenido o potencialmente uno inyectado
HEADERS: Dict[str, Optional[str]] = {
    'Authorization': None,
    'Content-Type': 'application/json'
}
GRAPH_API_TIMEOUT = 45

_cached_token: Optional[Dict[str, Any]] = None

def obtener_token() -> str:
    """
    Obtiene un token de acceso de aplicación usando credenciales de cliente.
    NOTA: Si usas permisos delegados, necesitarás un mecanismo diferente para obtener
          e inyectar el token en los HEADERS. Esta función asume Client Credentials.
    """
    global _cached_token
    # TODO: Implementar lógica real de caché y expiración de token
    # TODO: Manejar caso donde CLIENT_ID/SECRET no estén si se usa auth delegada

    # Verificar si tenemos credenciales de cliente
    if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET]):
        logger.error("Faltan CLIENT_ID, TENANT_ID o CLIENT_SECRET para obtener token de aplicación.")
        # Si se espera un token delegado, este debería haberse inyectado en HEADERS externamente.
        # Verificar si ya existe un token en HEADERS.
        if HEADERS.get('Authorization'):
            logger.warning("Credenciales de cliente faltantes, pero ya existe un token en HEADERS. Se usará el existente.")
            # Extraer y devolver el token existente para consistencia, aunque _actualizar_headers lo re-inyectará.
            auth_header = HEADERS['Authorization']
            if auth_header.lower().startswith('bearer '):
                return auth_header[7:]
            else:
                 raise Exception("Token existente en HEADERS no tiene el formato 'Bearer <token>'.")
        else:
            raise Exception("Credenciales de cliente (CLIENT_ID, TENANT_ID, CLIENT_SECRET) no configuradas y no hay token existente.")

    logger.info("Obteniendo NUEVO token de acceso de aplicación (Client Credentials)...")
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        'client_id': CLIENT_ID,
        'scope': GRAPH_SCOPE,
        'client_secret': CLIENT_SECRET,
        'grant_type': 'client_credentials'
    }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    response = None
    try:
        response = requests.post(url, data=data, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        token_data = response.json()
        token = token_data.get('access_token')
        if not token:
            logger.error(f"No se encontró 'access_token' en la respuesta de autenticación. Respuesta: {token_data}")
            raise Exception("No se pudo obtener el token de acceso de la respuesta.")
        logger.info("Token de acceso de aplicación obtenido.")
        return token
    except requests.exceptions.Timeout:
        logger.error(f"Timeout al obtener token desde {url}")
        raise Exception("Timeout al contactar el servidor de autenticación.")
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"Error de red/HTTP al obtener token: {e}. Detalles: {error_details}")
        raise Exception(f"Error de red/HTTP al obtener token: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object')
        logger.error(f"Error al decodificar JSON del token: {e}. Respuesta recibida: {response_text}")
        raise Exception(f"Error al decodificar JSON del token: {e}")
    except Exception as e:
        logger.error(f"Error inesperado al obtener token: {e}", exc_info=True)
        raise

def _actualizar_headers(inyectar_token: Optional[str] = None) -> None:
    """
    Actualiza los HEADERS globales.
    Si se provee 'inyectar_token', lo usa. Sino, llama a obtener_token().
    """
    global HEADERS
    try:
        token_a_usar = None
        if inyectar_token:
            logger.info("Usando token inyectado externamente.")
            token_a_usar = inyectar_token
        else:
            # Intentar obtener token (puede fallar si no hay credenciales y no hay token existente)
            token_a_usar = obtener_token()

        HEADERS['Authorization'] = f'Bearer {token_a_usar}'
        # logger.info("Cabecera de autorización actualizada.")
    except Exception as e:
        logger.error(f"Falló la obtención/actualización del token para la cabecera: {e}")
        HEADERS['Authorization'] = None # Asegurar que no quede un token viejo si falla
        raise Exception(f"Fallo crítico al configurar la cabecera de autorización: {e}")

# --- FIN: Constantes y Autenticación ---


# --- INICIO: Funciones Auxiliares de Graph API ---

# ---- CORREO (Outlook) ----
# Estas funciones usan /users/{MAILBOX}/... por lo que requieren que MAILBOX esté configurado
# y que el token (sea de app o delegado) tenga permisos sobre ese buzón.
def listar_correos(top: int = 10, skip: int = 0, folder: str = 'Inbox', select: Optional[List[str]] = None, filter_query: Optional[str] = None, order_by: Optional[str] = None, mailbox: Optional[str] = None) -> Dict[str, Any]:
    """Lista correos de una carpeta específica."""
    _actualizar_headers()
    usuario = mailbox or MAILBOX # Usa el mailbox configurado o el provisto
    if not usuario: raise ValueError("Mailbox no especificado o configurado.")
    url = f"{BASE_URL}/users/{usuario}/mailFolders/{folder}/messages"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by
    response: Optional[requests.Response] = None
    try:
        clean_params = {k:v for k, v in params.items() if v is not None}
        logger.info(f"API Call: GET {url} Params: {clean_params}")
        response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data: Dict[str, Any] = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} correos de {usuario}/{folder}.")
        return data
    except Exception as e: logger.error(f"Error en listar_correos: {e}", exc_info=True); raise

def leer_correo(message_id: str, select: Optional[List[str]] = None, mailbox: Optional[str] = None) -> dict:
    """Lee un correo específico por su ID."""
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    if not usuario: raise ValueError("Mailbox no especificado o configurado.")
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
    params = {}
    response: Optional[requests.Response] = None
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    try:
        logger.info(f"API Call: GET {url} Params: {params}")
        response = requests.get(url, headers=HEADERS, params=params or None, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Correo '{message_id}' leído de {usuario}.")
        return data
    except Exception as e: logger.error(f"Error en leer_correo: {e}", exc_info=True); raise

def enviar_correo(destinatario: Union[str, List[str]], asunto: str, mensaje: str, cc: Optional[Union[str, List[str]]] = None, bcc: Optional[Union[str, List[str]]] = None, attachments: Optional[List[dict]] = None, from_email: Optional[str] = None, is_draft: bool = False, mailbox: Optional[str] = None) -> dict:
    """Envía un correo o guarda un borrador."""
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    if not usuario: raise ValueError("Mailbox no especificado o configurado.")
    if is_draft: url = f"{BASE_URL}/users/{usuario}/messages"; log_action = "Guardando borrador"
    else: url = f"{BASE_URL}/users/{usuario}/sendMail"; log_action = "Enviando correo"
    def normalize_recipients(rec_input: Union[str, List[str]], type_name: str) -> List[Dict[str, Any]]:
        if isinstance(rec_input, str): rec_list = [rec_input]
        elif isinstance(rec_input, list): rec_list = rec_input
        else: raise TypeError(f"{type_name} debe ser str o List[str]")
        return [{"emailAddress": {"address": r}} for r in rec_list if r and isinstance(r, str)]
    try: to_recipients = normalize_recipients(destinatario, "Destinatario"); cc_recipients = normalize_recipients(cc, "CC") if cc else []; bcc_recipients = normalize_recipients(bcc, "BCC") if bcc else []
    except TypeError as e: logger.error(f"Error en formato de destinatarios: {e}"); raise ValueError(f"Formato de destinatario inválido: {e}")
    if not to_recipients: logging.error("No destinatarios válidos."); raise ValueError("Destinatario válido requerido.")
    message_payload: Dict[str, Any] = {"subject": asunto, "body": {"contentType": "HTML", "content": mensaje},"toRecipients": to_recipients,}
    if cc_recipients: message_payload["ccRecipients"] = cc_recipients
    if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
    if attachments: message_payload["attachments"] = attachments
    if from_email: message_payload["from"] = {"emailAddress": {"address": from_email}}
    final_payload = {"message": message_payload, "saveToSentItems": "true"} if not is_draft else message_payload
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} ({log_action} desde {usuario})")
        response = requests.post(url, headers=HEADERS, json=final_payload, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        if not is_draft: logger.info(f"Correo enviado desde {usuario}."); return {"status": "Enviado", "code": response.status_code}
        else: data = response.json(); message_id = data.get('id'); logger.info(f"Borrador guardado en {usuario}. ID: {message_id}."); return {"status": "Borrador Guardado", "code": response.status_code, "id": message_id, "data": data}
    except Exception as e: logger.error(f"Error en {log_action}: {e}", exc_info=True); raise

def guardar_borrador(destinatario: Union[str, List[str]], asunto: str, mensaje: str, cc: Optional[Union[str, List[str]]] = None, bcc: Optional[Union[str, List[str]]] = None, attachments: Optional[List[dict]] = None, from_email: Optional[str] = None, mailbox: Optional[str] = None) -> dict:
    """Guarda un correo como borrador."""
    usuario = mailbox or MAILBOX; logger.info(f"Llamando a guardar_borrador para '{usuario}'. Asunto: '{asunto}'"); return enviar_correo(destinatario, asunto, mensaje, cc, bcc, attachments, from_email, is_draft=True, mailbox=usuario)

def enviar_borrador(message_id: str, mailbox: Optional[str] = None) -> dict:
    """Envía un borrador previamente guardado."""
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/send"; response: Optional[requests.Response] = None
    try: logger.info(f"API Call: POST {url} (Enviando borrador {message_id} de {usuario})"); response = requests.post(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Borrador '{message_id}' enviado desde {usuario}."); return {"status": "Borrador Enviado", "code": response.status_code}
    except Exception as e: logger.error(f"Error en enviar_borrador: {e}", exc_info=True); raise

def responder_correo(message_id: str, mensaje_respuesta: str, to_recipients: Optional[List[dict]] = None, reply_all: bool = False, mailbox: Optional[str] = None) -> dict:
    """Responde a un correo."""
    _actualizar_headers(); usuario = mailbox or MAILBOX; action = "replyAll" if reply_all else "reply"; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/{action}"
    payload: Dict[str, Any] = {"comment": mensaje_respuesta}
    if to_recipients and isinstance(to_recipients, list): payload["message"] = { "toRecipients": to_recipients }; logger.info(f"Respondiendo a {message_id} con destinatarios custom.")
    response: Optional[requests.Response] = None
    try: logger.info(f"API Call: POST {url} (Respondiendo{' a todos' if reply_all else ''} correo {message_id} de {usuario})"); response = requests.post(url, headers=HEADERS, json=payload, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Respuesta {'a todos ' if reply_all else ''}enviada correo '{message_id}' de {usuario}."); return {"status": "Respondido", "code": response.status_code}
    except Exception as e: logger.error(f"Error en responder_correo: {e}", exc_info=True); raise

def reenviar_correo(message_id: str, destinatarios: Union[str, List[str]], mensaje_reenvio: str = "FYI", mailbox: Optional[str] = None) -> dict:
    """Reenvía un correo."""
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/forward"
    if isinstance(destinatarios, str): destinatarios = [destinatarios]
    to_recipients_list = [{"emailAddress": {"address": r}} for r in destinatarios if r and isinstance(r, str)]
    if not to_recipients_list: raise ValueError("Destinatario válido requerido.")
    payload = {"toRecipients": to_recipients_list, "comment": mensaje_reenvio}; response: Optional[requests.Response] = None
    try: logger.info(f"API Call: POST {url} (Reenviando correo {message_id} de {usuario})"); response = requests.post(url, headers=HEADERS, json=payload, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Correo '{message_id}' reenviado desde {usuario}."); return {"status": "Reenviado", "code": response.status_code}
    except Exception as e: logger.error(f"Error en reenviar_correo: {e}", exc_info=True); raise

def eliminar_correo(message_id: str, mailbox: Optional[str] = None) -> dict:
    """Elimina un correo (mueve a Elementos Eliminados)."""
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
    response: Optional[requests.Response] = None
    try: logger.info(f"API Call: DELETE {url} (Eliminando correo {message_id} de {usuario})"); response = requests.delete(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Correo '{message_id}' eliminado de {usuario}."); return {"status": "Eliminado", "code": response.status_code}
    except Exception as e: logger.error(f"Error en eliminar_correo: {e}", exc_info=True); raise

# ---- CALENDARIO (Outlook) ----
# También usan /users/{MAILBOX}/...
def listar_eventos(top: int = 10, start_date: Optional[datetime] = None, end_date: Optional[datetime] = None, filter_query: Optional[str] = None, order_by: Optional[str] = None, select: Optional[List[str]] = None, use_calendar_view: bool = True, mailbox: Optional[str] = None) -> Dict[str, Any]:
    """Lista eventos del calendario."""
    _actualizar_headers(); usuario = mailbox or MAILBOX; base_endpoint = f"{BASE_URL}/users/{usuario}"
    params: Dict[str, Any] = {}; endpoint_suffix = ""
    def ensure_timezone(dt: Optional[datetime]) -> Optional[datetime]: return dt.replace(tzinfo=timezone.utc) if dt and isinstance(dt, datetime) and dt.tzinfo is None else dt
    start_date = ensure_timezone(start_date); end_date = ensure_timezone(end_date)
    if use_calendar_view and start_date and end_date:
        endpoint_suffix = "/calendarView"; params['startDateTime'] = start_date.isoformat(); params['endDateTime'] = end_date.isoformat()
        params['$top'] = int(top)
        if filter_query: params['$filter'] = filter_query
        if order_by: params['$orderby'] = order_by
        if select: params['$select'] = ','.join(select)
        log_msg = f"Listando eventos (CalendarView) de {usuario}"
    else:
        endpoint_suffix = "/events"; params['$top'] = int(top); filters = []
        if start_date: filters.append(f"start/dateTime ge '{start_date.isoformat()}'")
        if end_date: filters.append(f"end/dateTime le '{end_date.isoformat()}'")
        if filter_query: filters.append(f"({filter_query})")
        if filters: params['$filter'] = " and ".join(filters)
        if order_by: params['$orderby'] = order_by
        if select: params['$select'] = ','.join(select)
        log_msg = f"Listando eventos (/events) de {usuario}"
    url = f"{base_endpoint}{endpoint_suffix}"; clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try: logger.info(f"API Call: GET {url} Params: {clean_params}"); response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"{log_msg}. Obtenidos {len(data.get('value',[]))}."); return data
    except Exception as e: logger.error(f"Error en listar_eventos: {e}", exc_info=True); raise

def crear_evento(titulo: str, inicio: datetime, fin: datetime, asistentes: Optional[List[Dict[str, Any]]] = None, cuerpo: Optional[str] = None, es_reunion_online: bool = False, proveedor_reunion_online: str = "teamsForBusiness", recordatorio_minutos: Optional[int] = 15, ubicacion: Optional[str] = None, mostrar_como: str = "busy", mailbox: Optional[str] = None) -> Dict[str, Any]:
    """Crea un nuevo evento en el calendario."""
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/events"
    if not isinstance(inicio, datetime) or not isinstance(fin, datetime): raise ValueError("'inicio' y 'fin' deben ser datetimes.")
    if inicio.tzinfo is None: inicio = inicio.replace(tzinfo=timezone.utc)
    if fin.tzinfo is None: fin = fin.replace(tzinfo=timezone.utc)
    body: Dict[str, Any] = {"subject": titulo, "start": {"dateTime": inicio.isoformat(), "timeZone": "UTC"}, "end": {"dateTime": fin.isoformat(), "timeZone": "UTC"}, "showAs": mostrar_como}
    if asistentes is not None:
        if isinstance(asistentes, list) and all(isinstance(a, dict) and 'emailAddress' in a for a in asistentes): body["attendees"] = asistentes
        else: logger.warning(f"Formato inválido para 'asistentes': {asistentes}")
    if cuerpo: body["body"] = {"contentType": "HTML", "content": cuerpo}
    if ubicacion: body["location"] = {"displayName": ubicacion}
    if es_reunion_online: body["isOnlineMeeting"] = True
    if es_reunion_online and proveedor_reunion_online: body["onlineMeetingProvider"] = proveedor_reunion_online
    if recordatorio_minutos is not None: body["isReminderOn"] = True; body["reminderMinutesBeforeStart"] = recordatorio_minutos
    else: body["isReminderOn"] = False
    response: Optional[requests.Response] = None
    try: logger.info(f"API Call: POST {url} (Creando evento '{titulo}' para {usuario})"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Evento '{titulo}' creado para {usuario}. ID: {data.get('id')}"); return data
    except Exception as e: logger.error(f"Error en crear_evento: {e}", exc_info=True); raise

def actualizar_evento(evento_id: str, nuevos_valores: Dict[str, Any], mailbox: Optional[str] = None) -> Dict[str, Any]:
    """Actualiza un evento existente usando PATCH."""
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
    payload = nuevos_valores.copy()
    for date_key in ['start', 'end']:
        if date_key in payload and isinstance(payload[date_key], datetime):
            dt = payload[date_key]; dt = dt.replace(tzinfo=timezone.utc) if dt.tzinfo is None else dt; payload[date_key] = {"dateTime": dt.isoformat(), "timeZone": "UTC"}; logging.info(f"Proc. fecha '{date_key}': {payload[date_key]}")
    response: Optional[requests.Response] = None
    try: logger.info(f"API Call: PATCH {url} (Actualizando evento {evento_id} para {usuario})"); response = requests.patch(url, headers=HEADERS, json=payload, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Evento '{evento_id}' actualizado para {usuario}."); return response.json()
    except Exception as e: logger.error(f"Error en actualizar_evento: {e}", exc_info=True); raise

def eliminar_evento(evento_id: str, mailbox: Optional[str] = None) -> Dict[str, Any]:
    """Elimina un evento del calendario."""
    _actualizar_headers(); usuario = mailbox or MAILBOX; url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
    response: Optional[requests.Response] = None
    try: logger.info(f"API Call: DELETE {url} (Eliminando evento {evento_id} para {usuario})"); response = requests.delete(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Evento '{evento_id}' eliminado para {usuario}."); return {"status": "Eliminado", "code": response.status_code}
    except Exception as e: logger.error(f"Error en eliminar_evento: {e}", exc_info=True); raise

# ---- TEAMS y OTROS ----
# ¡¡RESTAURADO a /me !! Asegúrate de que el token tiene permisos delegados adecuados.
def listar_chats(top: int = 20, skip: int = 0, filter_query: Optional[str] = None, order_by: Optional[str] = None, expand: Optional[str] = None) -> Dict[str, Any]:
    """Lista los chats del usuario actual (/me). Requiere permisos delegados (ej: Chat.Read)."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/chats" # Restaurado a /me
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query: params['$filter'] = filter_query
    if order_by: params['$orderby'] = order_by
    if expand: params['$expand'] = expand
    clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try: logger.info(f"API Call: GET {url} Params: {clean_params} (Listando chats para /me)"); response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Listados {len(data.get('value',[]))} chats para /me."); return data
    except Exception as e: logger.error(f"Error en listar_chats (/me): {e}", exc_info=True); raise

def listar_equipos(top: int = 20, skip: int = 0, filter_query: Optional[str] = None) -> Dict[str, Any]:
    """Lista los equipos a los que pertenece el usuario actual (/me). Requiere permisos delegados (ej: Team.ReadBasic.All)."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/joinedTeams" # Restaurado a /me
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query: params['$filter'] = filter_query
    clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try: logger.info(f"API Call: GET {url} Params: {clean_params} (Listando equipos para /me)"); response = requests.get(url, headers=HEADERS, params=clean_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Listados {len(data.get('value',[]))} equipos para /me."); return data
    except Exception as e: logger.error(f"Error en listar_equipos (/me): {e}", exc_info=True); raise

def obtener_equipo(team_id: str, select: Optional[List[str]]=None) -> Dict[str, Any]:
    """Obtiene detalles de un equipo específico por su ID. Requiere Team.ReadBasic.All (app o delegado)."""
    _actualizar_headers()
    url = f"{BASE_URL}/teams/{team_id}" # Este endpoint es por ID, no usa /me o /users
    params: Dict[str, Any] = {}
    if select: params['$select'] = ','.join(select)
    clean_params = {k:v for k, v in params.items() if v is not None}; response: Optional[requests.Response] = None
    try: logger.info(f"API Call: GET {url} Params: {clean_params} (Obteniendo equipo {team_id})"); response = requests.get(url, headers=HEADERS, params=clean_params or None, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Obtenido equipo ID: {team_id}."); return data
    except Exception as e: logger.error(f"Error en obtener_equipo: {e}", exc_info=True); raise

# ---- SHAREPOINT ----
# Estas funciones usan /sites/{site_id}/... Requieren Site ID y permisos adecuados (app o delegado).
_cached_root_site_id: Optional[str] = None
def obtener_site_id(site_url_relative: Optional[str] = None, hostname: Optional[str] = None) -> str:
    """Obtiene el ID de un sitio de SharePoint por su URL relativa y hostname, o el sitio raíz."""
    global _cached_root_site_id; _actualizar_headers()
    if site_url_relative and hostname: site_path = f"{hostname}:/{site_url_relative.strip('/')}"; url = f"{BASE_URL}/sites/{site_path}"; log_msg = f"Obteniendo Site ID para {site_path}"
    elif SHAREPOINT_DEFAULT_SITE_ID: logger.info(f"Usando Site ID de config: {SHAREPOINT_DEFAULT_SITE_ID}"); return SHAREPOINT_DEFAULT_SITE_ID
    else:
        if _cached_root_site_id: logger.info(f"Usando Site ID raíz cacheado: {_cached_root_site_id}"); return _cached_root_site_id
        url = f"{BASE_URL}/sites/root"; log_msg = "Obteniendo Site ID raíz"
    try:
        logger.info(f"API Call: GET {url} ({log_msg})"); response = requests.get(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); site_data = response.json(); site_id = site_data.get('id')
        if not site_id: logger.error(f"Respuesta sitio inválida, falta 'id': {site_data}"); raise Exception("Respuesta sitio inválida.")
        logger.info(f"Site ID obtenido: {site_id}");
        if url.endswith("/sites/root"): _cached_root_site_id = site_id
        return site_id
    except requests.exceptions.RequestException as e:
        status_code = getattr(e.response, 'status_code', None); error_details = getattr(e.response, 'text', str(e))
        if status_code == 404: logger.error(f"Error 404: Sitio no encontrado en {url}. Det: {error_details}"); raise Exception(f"Sitio no encontrado: {url}")
        logger.error(f"Error API (obtener Site ID): {e}. Det: {error_details}"); raise Exception(f"Error API (obtener Site ID): {e}")
    except Exception as e: logger.error(f"Error inesperado obtener Site ID: {e}", exc_info=True); raise

# -- SHAREPOINT - Listas --
def sp_crear_lista(nombre_lista: str, site_id: Optional[str] = None) -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); url = f"{BASE_URL}/sites/{target_site_id}/lists"
    body = {"displayName": nombre_lista, "columns": [{"name": "Title", "text": {}},{"name": "Clave", "text": {}},{"name": "Valor", "text": {}}],"list": {"template": "genericList"}}
    try: logger.info(f"API Call: POST {url} (Creando lista '{nombre_lista}' sitio {target_site_id})"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Lista '{nombre_lista}' creada sitio {target_site_id}. ID: {data.get('id')}"); return data
    except Exception as e: logger.error(f"Error sp_crear_lista: {e}", exc_info=True); raise

def sp_listar_listas(site_id: Optional[str] = None) -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); url = f"{BASE_URL}/sites/{target_site_id}/lists"; params = {'$select': 'id,name,displayName,webUrl'}
    try: logger.info(f"API Call: GET {url} (Listando listas sitio {target_site_id})"); response = requests.get(url, headers=HEADERS, params=params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Listadas {len(data.get('value',[]))} listas sitio {target_site_id}."); return data
    except Exception as e: logger.error(f"Error sp_listar_listas: {e}", exc_info=True); raise

def sp_agregar_elemento_lista(lista_id_o_nombre: str, campos: Dict[str, Any], site_id: Optional[str] = None) -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items"; body = {"fields": campos}
    try: logger.info(f"API Call: POST {url} (Agregando elem. lista '{lista_id_o_nombre}' sitio {target_site_id})"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Elem. agregado lista '{lista_id_o_nombre}'. ID: {data.get('id')}"); return data
    except Exception as e: logger.error(f"Error sp_agregar_elemento_lista: {e}", exc_info=True); raise

def sp_listar_elementos_lista(lista_id_o_nombre: str, site_id: Optional[str] = None, expand_fields: bool = True, top: int = 100, filter_query: Optional[str] = None, select: Optional[List[str]] = None) -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items"
    params: Dict[str, Any] = {'$top': min(int(top), 999)}
    if expand_fields: params['$expand'] = 'fields'
    if filter_query: params['$filter'] = filter_query
    if select: params['$select'] = ','.join(select + (['id'] if 'id' not in select else []))
    all_items = []; current_url: Optional[str] = url
    try:
        page_count = 0
        while current_url:
            page_count += 1; logger.info(f"API Call: GET {current_url} Page: {page_count} (Listando elems lista '{lista_id_o_nombre}', sitio {target_site_id})")
            current_params = params if page_count == 1 else None; response = requests.get(current_url, headers=HEADERS, params=current_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); page_items = data.get('value', []); all_items.extend(page_items); logger.info(f"Página {page_count}: {len(page_items)} elems.")
            current_url = data.get('@odata.nextLink')
            if current_url: logger.info("SP Paginación: Sig."); _actualizar_headers()
            else: logger.info("SP Paginación: Fin.")
        logger.info(f"Total elems. lista '{lista_id_o_nombre}': {len(all_items)}"); return {'value': all_items}
    except Exception as e: logger.error(f"Error sp_listar_elementos_lista: {e}", exc_info=True); raise

def sp_actualizar_elemento_lista(lista_id_o_nombre: str, item_id: str, nuevos_valores_campos: dict, site_id: Optional[str] = None) -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items/{item_id}/fields"
    try: logger.info(f"API Call: PATCH {url} (Actualizando item {item_id} lista '{lista_id_o_nombre}', sitio {target_site_id})"); response = requests.patch(url, headers=HEADERS, json=nuevos_valores_campos, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Elem. '{item_id}' actualizado lista '{lista_id_o_nombre}'."); return data
    except Exception as e: logger.error(f"Error sp_actualizar_elemento_lista: {e}", exc_info=True); raise

def sp_eliminar_elemento_lista(lista_id_o_nombre: str, item_id: str, site_id: Optional[str] = None) -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items/{item_id}"
    try: logger.info(f"API Call: DELETE {url} (Eliminando item {item_id} lista '{lista_id_o_nombre}', sitio {target_site_id})"); response = requests.delete(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Elem. '{item_id}' eliminado lista '{lista_id_o_nombre}'."); return {"status": "Eliminado", "code": response.status_code}
    except Exception as e: logger.error(f"Error sp_eliminar_elemento_lista: {e}", exc_info=True); raise

# -- SHAREPOINT - Documentos (Bibliotecas / Drives) --
def _get_sp_drive_endpoint(site_id: str, drive_id_or_name: Optional[str] = None) -> str:
    target_drive = drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID; return f"{BASE_URL}/sites/{site_id}/drives/{target_drive}"
def _get_sp_item_path_endpoint(site_id: str, item_path: str, drive_id_or_name: Optional[str] = None) -> str:
    drive_endpoint = _get_sp_drive_endpoint(site_id, drive_id_or_name); safe_path = item_path.strip()
    if not safe_path.startswith('/'): safe_path = '/' + safe_path
    return f"{drive_endpoint}/root" if safe_path == '/' else f"{drive_endpoint}/root:{safe_path}"

def sp_listar_archivos_carpeta(ruta_carpeta: str = '/', site_id: Optional[str] = None, drive_id_or_name: Optional[str] = None, top: int = 100) -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); item_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_carpeta, drive_id_or_name); url = f"{item_endpoint}/children"; params = {'$top': min(int(top), 999)}
    all_items = []; current_url: Optional[str] = url
    try:
        page_count = 0
        while current_url:
            page_count += 1; logger.info(f"API Call: GET {current_url} Page: {page_count} (Listando SP '{ruta_carpeta}' drive '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}', sitio {target_site_id})")
            current_params = params if page_count == 1 else None; response = requests.get(current_url, headers=HEADERS, params=current_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); page_items = data.get('value', []); all_items.extend(page_items); logger.info(f"Página {page_count}: {len(page_items)} items.")
            current_url = data.get('@odata.nextLink')
            if current_url: logger.info("SP Paginación: Sig."); _actualizar_headers()
            else: logger.info("SP Paginación: Fin.")
        logger.info(f"Total items SP en '{ruta_carpeta}': {len(all_items)}"); return {'value': all_items}
    except Exception as e: logger.error(f"Error sp_listar_archivos_carpeta: {e}", exc_info=True); raise

def sp_subir_archivo(nombre_archivo_destino: str, contenido_bytes: bytes, ruta_carpeta_destino: str = '/', site_id: Optional[str] = None, drive_id_or_name: Optional[str] = None, conflict_behavior: str = "rename") -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); target_file_path = os.path.join(ruta_carpeta_destino, nombre_archivo_destino).replace('\\', '/'); item_endpoint = _get_sp_item_path_endpoint(target_site_id, target_file_path, drive_id_or_name); url = f"{item_endpoint}:/content?@microsoft.graph.conflictBehavior={conflict_behavior}"
    upload_headers = HEADERS.copy(); upload_headers['Content-Type'] = 'application/octet-stream'
    try:
        logger.info(f"API Call: PUT {item_endpoint}:/content (Subiendo SP '{nombre_archivo_destino}' a '{ruta_carpeta_destino}' drive '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}', sitio {target_site_id})")
        if len(contenido_bytes) > 4 * 1024 * 1024: logger.warning(f"Archivo SP '{nombre_archivo_destino}' > 4MB. Considerar sesión de carga.")
        response = requests.put(url, headers=upload_headers, data=contenido_bytes, timeout=GRAPH_API_TIMEOUT * 3); response.raise_for_status(); data = response.json(); logger.info(f"Archivo SP '{nombre_archivo_destino}' subido. ID: {data.get('id')}"); return data
    except Exception as e: logger.error(f"Error sp_subir_archivo: {e}", exc_info=True); raise

def sp_descargar_archivo(ruta_archivo: str, site_id: Optional[str] = None, drive_id_or_name: Optional[str] = None) -> bytes:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); item_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_archivo, drive_id_or_name); url = f"{item_endpoint}/content"
    try: logger.info(f"API Call: GET {url} (Descargando SP '{ruta_archivo}' drive '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}', sitio {target_site_id})"); response = requests.get(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT * 2); response.raise_for_status(); logger.info(f"Contenido SP '{ruta_archivo}' descargado."); return response.content
    except Exception as e: logger.error(f"Error sp_descargar_archivo: {e}", exc_info=True); raise

def sp_eliminar_archivo_o_carpeta(ruta_item: str, site_id: Optional[str] = None, drive_id_or_name: Optional[str] = None) -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); item_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_item, drive_id_or_name); url = item_endpoint
    try: logger.info(f"API Call: DELETE {url} (Eliminando SP '{ruta_item}' drive '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}', sitio {target_site_id})"); response = requests.delete(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Item SP '{ruta_item}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
    except Exception as e: logger.error(f"Error sp_eliminar_archivo_o_carpeta: {e}", exc_info=True); raise

def sp_crear_carpeta(nombre_nueva_carpeta: str, ruta_carpeta_padre: str = '/', site_id: Optional[str] = None, drive_id_or_name: Optional[str] = None, conflict_behavior: str = "rename") -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); parent_folder_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_carpeta_padre, drive_id_or_name); url = f"{parent_folder_endpoint}/children"
    body = {"name": nombre_nueva_carpeta, "folder": {}, "@microsoft.graph.conflictBehavior": conflict_behavior}
    try: logger.info(f"API Call: POST {url} (Creando SP carpeta '{nombre_nueva_carpeta}' en '{ruta_carpeta_padre}', drive '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}', sitio {target_site_id})"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Carpeta SP '{nombre_nueva_carpeta}' creada. ID: {data.get('id')}"); return data
    except Exception as e: logger.error(f"Error sp_crear_carpeta: {e}", exc_info=True); raise

def sp_mover_archivo_o_carpeta(ruta_item_origen: str, ruta_carpeta_destino: str, nuevo_nombre: Optional[str] = None, site_id: Optional[str] = None, drive_id_or_name: Optional[str] = None) -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); target_drive_id = drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID
    drive_info_url = _get_sp_drive_endpoint(target_site_id, target_drive_id)
    try: drive_resp = requests.get(drive_info_url, headers=HEADERS, params={'$select':'id'}, timeout=GRAPH_API_TIMEOUT); drive_resp.raise_for_status(); actual_drive_id = drive_resp.json().get('id'); assert actual_drive_id
    except Exception as e: logger.error(f"Error obteniendo ID drive '{target_drive_id}': {e}"); raise Exception(f"Error obteniendo ID drive destino: {e}")
    item_origen_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_item_origen, target_drive_id); url = item_origen_endpoint
    parent_reference_path = f"/drives/{actual_drive_id}/root:{ruta_carpeta_destino.strip()}" if ruta_carpeta_destino != '/' else f"/drives/{actual_drive_id}/root"
    body = {"parentReference": {"path": parent_reference_path}};
    if nuevo_nombre: body["name"] = nuevo_nombre
    try: logger.info(f"API Call: PATCH {url} (Moviendo SP '{ruta_item_origen}' a '{ruta_carpeta_destino}' drive {actual_drive_id}, sitio {target_site_id})"); response = requests.patch(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Item SP '{ruta_item_origen}' movido a '{ruta_carpeta_destino}'. ID: {data.get('id')}"); return data
    except Exception as e: logger.error(f"Error sp_mover_archivo_o_carpeta: {e}", exc_info=True); raise

def sp_copiar_archivo_o_carpeta(ruta_item_origen: str, ruta_carpeta_destino: str, nuevo_nombre: Optional[str] = None, site_id: Optional[str] = None, drive_id_or_name: Optional[str] = None) -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); target_drive_id = drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID
    drive_info_url = _get_sp_drive_endpoint(target_site_id, target_drive_id)
    try: drive_resp = requests.get(drive_info_url, headers=HEADERS, params={'$select':'id'}, timeout=GRAPH_API_TIMEOUT); drive_resp.raise_for_status(); actual_drive_id = drive_resp.json().get('id'); assert actual_drive_id
    except Exception as e: logger.error(f"Error obteniendo ID drive '{target_drive_id}' copia: {e}"); raise Exception(f"Error obteniendo ID drive destino copia: {e}")
    item_origen_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_item_origen, target_drive_id); url = f"{item_origen_endpoint}/copy"
    parent_reference_path = f"/drive/root:{ruta_carpeta_destino.strip()}" if ruta_carpeta_destino != '/' else "/drive/root"
    body = {"parentReference": {"driveId": actual_drive_id, "path": parent_reference_path}};
    if nuevo_nombre: body["name"] = nuevo_nombre
    try: logger.info(f"API Call: POST {url} (Iniciando copia SP '{ruta_item_origen}' a '{ruta_carpeta_destino}' drive {actual_drive_id}, sitio {target_site_id})"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); monitor_url = response.headers.get('Location'); logger.info(f"Copia SP '{ruta_item_origen}' iniciada. Monitor: {monitor_url}"); return {"status": "Copia Iniciada", "code": response.status_code, "monitorUrl": monitor_url}
    except Exception as e: logger.error(f"Error sp_copiar_archivo_o_carpeta: {e}", exc_info=True); raise

def sp_obtener_metadatos_item(ruta_item: str, site_id: Optional[str] = None, drive_id_or_name: Optional[str] = None, select: Optional[List[str]] = None, expand: Optional[List[str]] = None) -> dict:
    _actualizar_headers(); target_site_id = site_id or obtener_site_id(); item_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_item, drive_id_or_name); url = item_endpoint; params = {}
    if select: params['$select'] = ','.join(select)
    if expand: params['$expand'] = ','.join(expand)
    try: logger.info(f"API Call: GET {url} (Obteniendo metadatos SP '{ruta_item}' drive '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}', sitio {target_site_id})"); response = requests.get(url, headers=HEADERS, params=params or None, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Metadatos SP '{ruta_item}' obtenidos."); return data
    except Exception as e: logger.error(f"Error sp_obtener_metadatos_item: {e}", exc_info=True); raise


# ---- ONEDRIVE ----
# ¡¡RESTAURADO a /me !! Interactúa con el OneDrive del usuario autenticado (/me).
# Requiere permisos delegados (ej: Files.ReadWrite).

def _get_od_me_drive_endpoint() -> str:
    """Helper para obtener el endpoint del drive /me."""
    return f"{BASE_URL}/me/drive"

def _get_od_me_item_path_endpoint(ruta_relativa: str) -> str:
    """Helper para construir la URL a un item por ruta en OneDrive /me."""
    drive_endpoint = _get_od_me_drive_endpoint()
    safe_path = ruta_relativa.strip()
    if not safe_path.startswith('/'): safe_path = '/' + safe_path
    return f"{drive_endpoint}/root" if safe_path == '/' else f"{drive_endpoint}/root:{safe_path}"

def od_listar_archivos(ruta: str = "/", top: int = 100) -> dict:
    """Lista archivos y carpetas en una ruta del OneDrive del usuario (/me)."""
    _actualizar_headers(); item_endpoint = _get_od_me_item_path_endpoint(ruta); url = f"{item_endpoint}/children"; params = {'$top': min(int(top), 999)}
    all_items = []; current_url: Optional[str] = url
    try:
        page_count = 0
        while current_url:
            page_count += 1; logger.info(f"API Call: GET {current_url} Page: {page_count} (Listando OD /me '{ruta}')")
            current_params = params if page_count == 1 else None; response = requests.get(current_url, headers=HEADERS, params=current_params, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); page_items = data.get('value', []); all_items.extend(page_items); logger.info(f"Página {page_count}: {len(page_items)} items.")
            current_url = data.get('@odata.nextLink')
            if current_url: logger.info("OD Paginación: Sig."); _actualizar_headers()
            else: logger.info("OD Paginación: Fin.")
        logger.info(f"Total items OD /me en '{ruta}': {len(all_items)}"); return {'value': all_items}
    except Exception as e: logger.error(f"Error od_listar_archivos (/me): {e}", exc_info=True); raise

def od_subir_archivo(nombre_archivo_destino: str, contenido_bytes: bytes, ruta_carpeta_destino: str = "/", conflict_behavior: str = "rename") -> dict:
    """Sube un archivo al OneDrive del usuario (/me)."""
    _actualizar_headers(); target_file_path = os.path.join(ruta_carpeta_destino, nombre_archivo_destino).replace('\\', '/'); item_endpoint = _get_od_me_item_path_endpoint(target_file_path); url = f"{item_endpoint}:/content?@microsoft.graph.conflictBehavior={conflict_behavior}"
    upload_headers = HEADERS.copy(); upload_headers['Content-Type'] = 'application/octet-stream'
    try:
        logger.info(f"API Call: PUT {item_endpoint}:/content (Subiendo OD /me '{nombre_archivo_destino}' a '{ruta_carpeta_destino}')")
        if len(contenido_bytes) > 4 * 1024 * 1024: logger.warning(f"Archivo OD /me '{nombre_archivo_destino}' > 4MB. Considerar sesión de carga.")
        response = requests.put(url, headers=upload_headers, data=contenido_bytes, timeout=GRAPH_API_TIMEOUT * 3); response.raise_for_status(); data = response.json(); logger.info(f"Archivo OD /me '{nombre_archivo_destino}' subido. ID: {data.get('id')}"); return data
    except Exception as e: logger.error(f"Error od_subir_archivo (/me): {e}", exc_info=True); raise

def od_descargar_archivo(ruta_archivo: str) -> bytes:
    """Descarga el contenido binario de un archivo del OneDrive del usuario (/me)."""
    _actualizar_headers(); item_endpoint = _get_od_me_item_path_endpoint(ruta_archivo); url = f"{item_endpoint}/content"
    try: logger.info(f"API Call: GET {url} (Descargando OD /me '{ruta_archivo}')"); response = requests.get(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT * 2); response.raise_for_status(); logger.info(f"Contenido OD /me '{ruta_archivo}' descargado."); return response.content
    except Exception as e: logger.error(f"Error od_descargar_archivo (/me): {e}", exc_info=True); raise

def od_eliminar_archivo_o_carpeta(ruta_item: str) -> dict:
    """Elimina un archivo o carpeta del OneDrive del usuario (/me)."""
    _actualizar_headers(); item_endpoint = _get_od_me_item_path_endpoint(ruta_item); url = item_endpoint
    try: logger.info(f"API Call: DELETE {url} (Eliminando OD /me '{ruta_item}')"); response = requests.delete(url, headers=HEADERS, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); logger.info(f"Item OD /me '{ruta_item}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
    except Exception as e: logger.error(f"Error od_eliminar_archivo_o_carpeta (/me): {e}", exc_info=True); raise

def od_crear_carpeta(nombre_nueva_carpeta: str, ruta_carpeta_padre: str = "/", conflict_behavior: str = "rename") -> dict:
    """Crea una nueva carpeta en el OneDrive del usuario (/me)."""
    _actualizar_headers(); parent_folder_endpoint = _get_od_me_item_path_endpoint(ruta_carpeta_padre); url = f"{parent_folder_endpoint}/children"
    body = {"name": nombre_nueva_carpeta, "folder": {}, "@microsoft.graph.conflictBehavior": conflict_behavior}
    try: logger.info(f"API Call: POST {url} (Creando OD /me carpeta '{nombre_nueva_carpeta}' en '{ruta_carpeta_padre}')"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Carpeta OD /me '{nombre_nueva_carpeta}' creada. ID: {data.get('id')}"); return data
    except Exception as e: logger.error(f"Error od_crear_carpeta (/me): {e}", exc_info=True); raise

def od_mover_archivo_o_carpeta(ruta_item_origen: str, ruta_carpeta_destino: str, nuevo_nombre: Optional[str] = None) -> dict:
    """Mueve un archivo o carpeta en el OneDrive del usuario (/me)."""
    _actualizar_headers(); item_origen_endpoint = _get_od_me_item_path_endpoint(ruta_item_origen); url = item_origen_endpoint
    parent_reference_path = f"/drive/root:{ruta_carpeta_destino.strip()}" if ruta_carpeta_destino != '/' else "/drive/root"
    body = {"parentReference": { "path": parent_reference_path }};
    if nuevo_nombre: body["name"] = nuevo_nombre
    try: logger.info(f"API Call: PATCH {url} (Moviendo OD /me '{ruta_item_origen}' a '{ruta_carpeta_destino}')"); response = requests.patch(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Item OD /me '{ruta_item_origen}' movido a '{ruta_carpeta_destino}'. ID: {data.get('id')}"); return data
    except Exception as e: logger.error(f"Error od_mover_archivo_o_carpeta (/me): {e}", exc_info=True); raise

def od_copiar_archivo_o_carpeta(ruta_item_origen: str, ruta_carpeta_destino: str, nuevo_nombre: Optional[str] = None) -> dict:
    """Inicia la copia de un archivo o carpeta en el OneDrive del usuario (/me)."""
    _actualizar_headers()
    drive_endpoint = _get_od_me_drive_endpoint()
    try: drive_resp = requests.get(drive_endpoint, headers=HEADERS, params={'$select':'id'}, timeout=GRAPH_API_TIMEOUT); drive_resp.raise_for_status(); actual_drive_id = drive_resp.json().get('id'); assert actual_drive_id
    except Exception as e: logger.error(f"Error obteniendo ID drive OD /me: {e}"); raise Exception(f"Error obteniendo ID drive OD /me: {e}")
    item_origen_endpoint = _get_od_me_item_path_endpoint(ruta_item_origen); url = f"{item_origen_endpoint}/copy"
    parent_reference_path = f"/drive/root:{ruta_carpeta_destino.strip()}" if ruta_carpeta_destino != '/' else "/drive/root"
    body = {"parentReference": {"driveId": actual_drive_id, "path": parent_reference_path }};
    if nuevo_nombre: body["name"] = nuevo_nombre
    try: logger.info(f"API Call: POST {url} (Iniciando copia OD /me '{ruta_item_origen}' a '{ruta_carpeta_destino}')"); response = requests.post(url, headers=HEADERS, json=body, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); monitor_url = response.headers.get('Location'); logger.info(f"Copia OD /me '{ruta_item_origen}' iniciada. Monitor: {monitor_url}"); return {"status": "Copia Iniciada", "code": response.status_code, "monitorUrl": monitor_url}
    except Exception as e: logger.error(f"Error od_copiar_archivo_o_carpeta (/me): {e}", exc_info=True); raise

def od_obtener_metadatos_item(ruta_item: str, select: Optional[List[str]] = None) -> dict:
    """Obtiene los metadatos de un archivo o carpeta en el OneDrive del usuario (/me)."""
    _actualizar_headers(); item_endpoint = _get_od_me_item_path_endpoint(ruta_item); url = item_endpoint; params = {}
    if select: params['$select'] = ','.join(select)
    try: logger.info(f"API Call: GET {url} (Obteniendo metadatos OD /me '{ruta_item}')"); response = requests.get(url, headers=HEADERS, params=params or None, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Metadatos OD /me '{ruta_item}' obtenidos."); return data
    except Exception as e: logger.error(f"Error od_obtener_metadatos_item (/me): {e}", exc_info=True); raise

def od_actualizar_metadatos_item(ruta_item: str, nuevos_valores: dict) -> dict:
    """Actualiza los metadatos de un archivo o carpeta en el OneDrive del usuario (/me)."""
    _actualizar_headers(); item_endpoint = _get_od_me_item_path_endpoint(ruta_item); url = item_endpoint
    try: logger.info(f"API Call: PATCH {url} (Actualizando metadatos OD /me '{ruta_item}')"); response = requests.patch(url, headers=HEADERS, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT); response.raise_for_status(); data = response.json(); logger.info(f"Metadatos OD /me '{ruta_item}' actualizados."); return data
    except Exception as e: logger.error(f"Error od_actualizar_metadatos_item (/me): {e}", exc_info=True); raise


# --- FIN: Funciones Auxiliares de Graph API ---


# --- INICIO: Función Principal de Azure Functions (Entry Point) ---

# Mapeo de nombres de acción a las funciones DEFINIDAS ARRIBA
acciones_disponibles: Dict[str, Callable[..., Any]] = {
    # Correo (usa /users/{MAILBOX}/...)
    "mail_listar": listar_correos, "mail_leer": leer_correo, "mail_enviar": enviar_correo,
    "mail_guardar_borrador": guardar_borrador, "mail_enviar_borrador": enviar_borrador,
    "mail_responder": responder_correo, "mail_reenviar": reenviar_correo, "mail_eliminar": eliminar_correo,
    # Calendario (usa /users/{MAILBOX}/...)
    "cal_listar_eventos": listar_eventos, "cal_crear_evento": crear_evento,
    "cal_actualizar_evento": actualizar_evento, "cal_eliminar_evento": eliminar_evento,
    # Teams/Otros (usa /me/... o /teams/{id})
    "team_listar_chats": listar_chats,          # Usa /me
    "team_listar_equipos": listar_equipos,      # Usa /me
    "team_obtener_equipo": obtener_equipo,      # Usa /teams/{id}
    # SharePoint - Listas (usa /sites/{site_id}/...)
    "sp_crear_lista": sp_crear_lista, "sp_listar_listas": sp_listar_listas,
    "sp_agregar_elemento_lista": sp_agregar_elemento_lista, "sp_listar_elementos_lista": sp_listar_elementos_lista,
    "sp_actualizar_elemento_lista": sp_actualizar_elemento_lista, "sp_eliminar_elemento_lista": sp_eliminar_elemento_lista,
    # SharePoint - Documentos/Bibliotecas (usa /sites/{site_id}/drives/{drive_id}/...)
    "sp_listar_archivos_carpeta": sp_listar_archivos_carpeta, "sp_subir_archivo": sp_subir_archivo,
    "sp_descargar_archivo": sp_descargar_archivo, "sp_eliminar_archivo_o_carpeta": sp_eliminar_archivo_o_carpeta,
    "sp_crear_carpeta": sp_crear_carpeta, "sp_mover_archivo_o_carpeta": sp_mover_archivo_o_carpeta,
    "sp_copiar_archivo_o_carpeta": sp_copiar_archivo_o_carpeta, "sp_obtener_metadatos_item": sp_obtener_metadatos_item,
    # OneDrive (Usuario) (usa /me/drive/...)
    "od_listar_archivos": od_listar_archivos, "od_subir_archivo": od_subir_archivo,
    "od_descargar_archivo": od_descargar_archivo, "od_eliminar_archivo_o_carpeta": od_eliminar_archivo_o_carpeta,
    "od_crear_carpeta": od_crear_carpeta, "od_mover_archivo_o_carpeta": od_mover_archivo_o_carpeta,
    "od_copiar_archivo_o_carpeta": od_copiar_archivo_o_carpeta, "od_obtener_metadatos_item": od_obtener_metadatos_item,
    "od_actualizar_metadatos_item": od_actualizar_metadatos_item,
}

# Verificar funciones mapeadas
for accion_check, func_ref_check in acciones_disponibles.items():
    if not callable(func_ref_check): logger.error(f"CONFIG ERROR INTERNO: Acción '{accion_check}' mapeada a no-callable.")

def main(req: func.HttpRequest) -> func.HttpResponse:
    """Punto de entrada principal. Maneja la solicitud HTTP, llama a la acción apropiada y devuelve la respuesta."""
    logging.info(f'Python HTTP trigger function procesando solicitud. Method: {req.method}, URL: {req.url}')
    invocation_id = req.headers.get('X-Azure-Functions-InvocationId', 'N/A')
    logging.info(f"Invocation ID: {invocation_id}")

    accion: Optional[str] = None
    parametros: Dict[str, Any] = {}
    funcion_a_ejecutar: Optional[Callable] = None
    resultado: Any = None
    inyectado_token: Optional[str] = None # Para posible token externo

    # --- INICIO: Bloque Try-Except General ---
    try:
        # --- Opcional: Permitir inyección de token vía header ---
        # Útil si la autenticación se maneja fuera y se pasa el token delegado
        auth_header = req.headers.get('Authorization')
        if auth_header and auth_header.lower().startswith('bearer '):
            inyectado_token = auth_header[7:]
            logger.info(f"Invocation {invocation_id}: Token 'Bearer' detectado en header. Se intentará usar este token.")
            # Actualizar HEADERS globales inmediatamente si se detecta token externo
            # Esto sobreescribirá cualquier token obtenido por client credentials si _actualizar_headers se llama sin argumento
            HEADERS['Authorization'] = auth_header

        # --- Leer accion y parametros ---
        req_body: Optional[Dict[str, Any]] = None
        content_type = req.headers.get('Content-Type', '').lower()

        if req.method in ('POST', 'PUT', 'PATCH'):
            if 'application/json' in content_type:
                try:
                    req_body = req.get_json(); assert isinstance(req_body, dict)
                    accion = req_body.get('accion'); parametros = req_body.get('parametros', {}); assert isinstance(parametros, dict)
                except (ValueError, AssertionError) as e: logger.warning(f'Invocation {invocation_id}: Cuerpo JSON inválido/no objeto: {e}'); return func.HttpResponse("Cuerpo JSON inválido o no es objeto.", status_code=400)
            elif 'multipart/form-data' in content_type:
                 logger.info(f"Invocation {invocation_id}: Procesando multipart/form-data."); parametros = {}
                 accion = req.form.get('accion')
                 for key, value in req.form.items():
                     if key not in ['accion', 'file']: parametros[key] = value
                 file = req.files.get('file')
                 if file: logger.info(f"Archivo '{file.filename}' en form-data."); parametros['contenido_bytes'] = file.read(); parametros['nombre_archivo_original'] = file.filename
                 else: logger.info("No archivo ('file') en form-data.")
            elif 'application/x-www-form-urlencoded' in content_type:
                 logger.info(f"Invocation {invocation_id}: Procesando x-www-form-urlencoded."); accion = req.form.get('accion'); parametros = {k: v for k, v in req.form.items() if k != 'accion'}
            else: accion = req.params.get('accion'); parametros = dict(req.params); logger.warning(f"Invocation {invocation_id}: {req.method} sin Content-Type conocido. Usando query params.")
        elif req.method == 'GET': accion = req.params.get('accion'); parametros = dict(req.params)
        else: accion = req.params.get('accion'); parametros = dict(req.params)
        if 'accion' in parametros: del parametros['accion'] # Limpiar acción de params

        # --- Validar acción ---
        if not accion or not isinstance(accion, str): logger.warning(f"Invocation {invocation_id}: 'accion' faltante o inválida."); return func.HttpResponse("Falta parámetro 'accion' (string).", status_code=400)
        logger.info(f"Invocation {invocation_id}: Acción solicitada: '{accion}'.")
        log_params = {k: f"<bytes len={len(v)}>" if isinstance(v, bytes) else v for k, v in parametros.items()} # Log seguro
        logger.info(f"Invocation {invocation_id}: Parámetros iniciales: {log_params}")

        # --- Buscar y preparar la función ---
        if accion in acciones_disponibles:
            funcion_a_ejecutar = acciones_disponibles[accion]
            logger.info(f"Invocation {invocation_id}: Mapeado a función: {funcion_a_ejecutar.__name__}")

            # --- Validar/Convertir parámetros ANTES de llamar ---
            params_procesados: Dict[str, Any] = {}
            try:
                params_procesados = parametros.copy()
                # Conversiones comunes
                int_params = ['top', 'skip', 'recordatorio_minutos']; bool_params = ['reply_all', 'es_reunion_online', 'expand_fields', 'use_calendar_view']; datetime_params = ['start_date', 'end_date', 'inicio', 'fin']
                for p in int_params:
                    if p in params_procesados and params_procesados[p] is not None: params_procesados[p] = int(params_procesados[p])
                for p in bool_params:
                    if p in params_procesados and params_procesados[p] is not None: val = str(params_procesados[p]).lower(); params_procesados[p] = val in ['true', '1', 'yes']
                for p in datetime_params:
                     if p in params_procesados and params_procesados[p] is not None:
                         if isinstance(params_procesados[p], str): params_procesados[p] = datetime.fromisoformat(params_procesados[p].replace('Z', '+00:00'))
                         elif not isinstance(params_procesados[p], datetime): raise ValueError(f"Tipo inválido para '{p}'.")
                # Validaciones específicas
                if accion in ["sp_subir_archivo", "od_subir_archivo"]:
                     if 'contenido_bytes' not in params_procesados: raise ValueError(f"Acción '{accion}' requiere 'contenido_bytes'.")
                     if 'nombre_archivo_destino' not in params_procesados: params_procesados['nombre_archivo_destino'] = params_procesados.get('nombre_archivo_original', f"subido_{datetime.now().strftime('%Y%m%d%H%M%S')}") # Usar original o generar nombre
            except (ValueError, TypeError, KeyError) as conv_err: logger.error(f"Invocation {invocation_id}: Error parámetros '{accion}': {conv_err}. Recibido: {log_params}", exc_info=True); return func.HttpResponse(f"Parámetros inválidos '{accion}': {conv_err}", status_code=400)
            except Exception as pre_exec_err: logger.error(f"Invocation {invocation_id}: Error pre-ejecución '{accion}': {pre_exec_err}. Parámetros: {log_params}", exc_info=True); return func.HttpResponse(f"Error interno preparando '{accion}'.", status_code=500)

            # --- Llamar a la función auxiliar ---
            logger.info(f"Invocation {invocation_id}: Ejecutando {funcion_a_ejecutar.__name__}...")
            try:
                # Pasar el token inyectado (si existe) a _actualizar_headers antes de la llamada real
                # Nota: _actualizar_headers dentro de cada función auxiliar se llamará de nuevo,
                # pero si el token inyectado ya está en HEADERS, debería reutilizarse (si obtener_token no fuerza recarga).
                # Por simplicidad, dejamos que las funciones auxiliares llamen a _actualizar_headers() sin argumento.
                # Si se inyectó un token, ya está en HEADERS. Si no, obtener_token() se ejecutará.

                resultado = funcion_a_ejecutar(**params_procesados)
                logger.info(f"Invocation {invocation_id}: Ejecución '{accion}' completada.")
            except TypeError as type_err: logger.error(f"Invocation {invocation_id}: Error argumento {funcion_a_ejecutar.__name__}: {type_err}. Params: {params_procesados}", exc_info=True); return func.HttpResponse(f"Error argumentos '{accion}': {type_err}", status_code=400)
            except Exception as exec_err: logger.exception(f"Invocation {invocation_id}: Error ejecución '{accion}': {exec_err}"); return func.HttpResponse(f"Error al ejecutar '{accion}': {exec_err}", status_code=500)

            # --- Devolver resultado ---
            if isinstance(resultado, bytes):
                 logger.info(f"Invocation {invocation_id}: Devolviendo bytes."); filename = os.path.basename(parametros.get('ruta_archivo') or parametros.get('ruta_item') or 'download'); return func.HttpResponse(resultado, mimetype="application/octet-stream", headers={'Content-Disposition': f'attachment; filename="{filename}"'}, status_code=200)
            elif isinstance(resultado, (dict, list)):
                 logger.info(f"Invocation {invocation_id}: Devolviendo JSON.");
                 try: json_response = json.dumps(resultado, default=str); return func.HttpResponse(json_response, mimetype="application/json", status_code=200)
                 except TypeError as serialize_err: logger.error(f"Invocation {invocation_id}: Error serializar JSON '{accion}': {serialize_err}.", exc_info=True); return func.HttpResponse(f"Error interno: Respuesta no serializable.", status_code=500)
            else: logger.warning(f"Invocation {invocation_id}: Tipo resultado inesperado '{accion}': {type(resultado)}. String."); return func.HttpResponse(str(resultado), mimetype="text/plain", status_code=200)
        else: # Acción no encontrada
            logger.warning(f"Invocation {invocation_id}: Acción '{accion}' no reconocida."); acciones_validas = list(acciones_disponibles.keys()); return func.HttpResponse(f"Acción '{accion}' no reconocida. Válidas: {acciones_validas}", status_code=400)
    except Exception as e: # Captura errores generales
        func_name = getattr(funcion_a_ejecutar, '__name__', 'N/A'); logger.exception(f"Invocation {invocation_id}: Error GENERAL main() acción '{accion or '?'}' (Func: {func_name}): {e}"); return func.HttpResponse("Error interno del servidor.", status_code=500)

# --- FIN: Función Principal ---
