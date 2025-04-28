import json
import logging
import requests
import azure.functions as func # Asegúrate que azure.functions está importado
from typing import Dict, Any, Callable, List, Optional, Union
from datetime import datetime, timezone
import os

# Configuración de logging
logger = logging.getLogger("azure.functions")
logger.setLevel(logging.INFO)

# --- INICIO: Variables de Entorno y Configuración ---
# (Tu código existente para leer variables y validar)
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
    # Si esto falla, la función no puede ni empezar. El error se loguea y se relanza.
    logger.critical(f"Error CRÍTICO de configuración inicial: {e}. La función no puede operar.")
    # En un escenario real, podrías querer manejar esto de forma diferente,
    # pero para Azure Functions, fallar rápido si falta config esencial está bien.
    raise # Relanzar la excepción detendrá la ejecución de la función.

# --- FIN: Variables de Entorno y Configuración ---


# --- INICIO: Constantes y Autenticación ---
# (Tu código existente para BASE_URL, HEADERS, obtener_token, _actualizar_headers)
BASE_URL = "https://graph.microsoft.com/v1.0"
HEADERS = {
    'Authorization': None,
    'Content-Type': 'application/json'
}

def obtener_token() -> str:
    logger.info("Obteniendo token de acceso...")
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
        response = requests.post(url, data=data, headers=headers)
        response.raise_for_status()
        token_data = response.json()
        token = token_data.get('access_token')
        if not token:
            logger.error(f"❌ No se encontró 'access_token' en la respuesta. Respuesta: {token_data}")
            raise Exception("No se pudo obtener el token de acceso de la respuesta.")
        logger.info(f"Token obtenido correctamente.")
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
    # Intenta obtener y establecer el token. Si falla, lanza excepción.
    try:
        token = obtener_token()
        HEADERS['Authorization'] = f'Bearer {token}'
        logger.info("Cabecera de autorización actualizada.")
    except Exception as e:
        # El error ya se loguea en obtener_token, relanzamos para detener la operación dependiente.
        logger.error(f"❌ Falló la actualización de la cabecera de autorización.")
        raise Exception(f"Fallo al actualizar la cabecera: {e}")

# --- FIN: Constantes y Autenticación ---


# --- INICIO: Funciones Auxiliares de Graph API ---
# (Todo tu código existente para listar_correos, leer_correo, enviar_correo,
#  guardar_borrador, enviar_borrador, responder_correo, reenviar_correo,
#  eliminar_correo, listar_eventos, crear_evento, actualizar_evento,
#  eliminar_evento, listar_chats, listar_equipos, obtener_equipo)

# ---- CORREO ----
def listar_correos(
    top: int = 10,
    skip: int = 0,
    folder: str = 'Inbox',
    select: Optional[List[str]] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None,
    mailbox: Optional[str] = None
) -> dict:
    # Asegura cabecera fresca ANTES de la llamada
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/mailFolders/{folder}/messages"
    params = {'$top': top, '$skip': skip}
    if select: params['$select'] = ','.join(select)
    if filter_query: params['$filter'] = filter_query
    if order_by: params['$orderby'] = order_by

    try:
        # Limpia params de valores None antes de la llamada
        clean_params = {k:v for k, v in params.items() if v is not None}
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}")
        response = requests.get(url, headers=HEADERS, params=clean_params)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} correos de la carpeta '{folder}' para usuario '{usuario}'.")
        return data
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al listar correos para '{usuario}' en carpeta '{folder}': {e}. URL: {url}. Detalles: {error_details}")
        # Considera devolver un error específico o relanzar una excepción más descriptiva
        raise Exception(f"Error al listar correos: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"❌ Error al decodificar JSON (listar correos): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (listar correos): {e}")


def leer_correo(
    message_id: str,
    select: Optional[List[str]] = None,
    mailbox: Optional[str] = None
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
    params = {}
    if select: params['$select'] = ','.join(select)

    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {params}")
        response = requests.get(url, headers=HEADERS, params=params or None)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Correo '{message_id}' leído para usuario '{usuario}'.")
        return data
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al leer el correo '{message_id}' para usuario '{usuario}': {e}. URL: {url}. Detalles: {error_details}")
        raise Exception(f"Error al leer el correo: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"❌ Error al decodificar JSON (leer correo): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (leer correo): {e}")


def enviar_correo(
    destinatario: Union[str, List[str]],
    asunto: str,
    mensaje: str,
    cc: Optional[Union[str, List[str]]] = None,
    bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[dict]] = None,
    is_draft: bool = False,
    mailbox: Optional[str] = None
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    endpoint = "messages" if is_draft else "sendMail"
    url = f"{BASE_URL}/users/{usuario}/{endpoint}"

    if isinstance(destinatario, str): destinatario = [destinatario]
    if isinstance(cc, str): cc = [cc]
    if isinstance(bcc, str): bcc = [bcc]

    to_recipients = [{"emailAddress": {"address": r}} for r in destinatario if r] # Filtra vacíos
    cc_recipients = [{"emailAddress": {"address": r}} for r in cc if r] if cc else []
    bcc_recipients = [{"emailAddress": {"address": r}} for r in bcc if r] if bcc else []

    if not to_recipients:
        raise ValueError("Se requiere al menos un destinatario.")

    message_payload = {
        "subject": asunto,
        "body": {"contentType": "HTML", "content": mensaje},
        "toRecipients": to_recipients,
    }
    if cc_recipients: message_payload["ccRecipients"] = cc_recipients
    if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
    if attachments: message_payload["attachments"] = attachments

    payload = {"message": message_payload, "saveToSentItems": "true"} if not is_draft else message_payload

    try:
        logger.info(f"Llamando a Graph API: POST {url}")
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status()
        if not is_draft:
            logger.info(f"Correo enviado por '{usuario}' con asunto '{asunto}'.")
            return {"status": "Enviado", "code": response.status_code}
        else:
            data = response.json()
            message_id = data.get('id')
            logger.info(f"Correo guardado como borrador por '{usuario}' con ID: {message_id}.")
            # Devuelve más info al guardar borrador, puede ser útil
            return {"status": "Borrador Guardado", "code": response.status_code, "id": message_id, "data": data}
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al {'enviar' if not is_draft else 'guardar borrador'} correo por '{usuario}': {e}. Detalles: {error_details}. URL: {url}")
        raise Exception(f"Error al {'enviar' if not is_draft else 'guardar borrador'} correo: {e}")
    except json.JSONDecodeError as e:
        # Solo aplica si 'is_draft' es True y la respuesta es inválida
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"❌ Error al decodificar JSON (guardar borrador): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (guardar borrador): {e}")


def guardar_borrador(
    destinatario: Union[str, List[str]],
    asunto: str,
    mensaje: str,
    cc: Optional[Union[str, List[str]]] = None,
    bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[dict]] = None,
    mailbox: Optional[str] = None
) -> dict:
    logger.info(f"Intentando guardar borrador para usuario '{mailbox or MAILBOX}' con asunto: '{asunto}'")
    return enviar_correo(destinatario, asunto, mensaje, cc, bcc, attachments, is_draft=True, mailbox=mailbox)


def enviar_borrador(
    message_id: str,
    mailbox: Optional[str] = None
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/send"
    try:
        logger.info(f"Llamando a Graph API: POST {url}")
        response = requests.post(url, headers=HEADERS)
        response.raise_for_status()
        logger.info(f"Borrador de correo '{message_id}' enviado por usuario '{usuario}'.")
        return {"status": "Borrador Enviado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al enviar borrador '{message_id}' por usuario '{usuario}': {e}. URL: {url}. Detalles: {error_details}")
        raise Exception(f"Error al enviar borrador: {e}")


def responder_correo(
    message_id: str,
    mensaje_respuesta: str,
    to_recipients: Optional[List[dict]] = None,
    reply_all: bool = False,
    mailbox: Optional[str] = None
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    action = "replyAll" if reply_all else "reply"
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/{action}"
    payload = {"comment": mensaje_respuesta}
    # Añadir destinatarios si se especifican explícitamente
    # if to_recipients:
    #    payload["message"] = {"toRecipients": to_recipients}

    try:
        logger.info(f"Llamando a Graph API: POST {url}")
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status()
        logger.info(f"Respuesta {'a todos ' if reply_all else ''}enviada al correo '{message_id}' por usuario '{usuario}'.")
        return {"status": "Respondido", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al responder {'a todos ' if reply_all else ''}al correo '{message_id}' por usuario '{usuario}': {e}. URL: {url}. Detalles: {error_details}")
        raise Exception(f"Error al responder al correo: {e}")


def reenviar_correo(
    message_id: str,
    destinatarios: Union[str, List[str]],
    mensaje_reenvio: str = "FYI",
    mailbox: Optional[str] = None
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/forward"

    if isinstance(destinatarios, str): destinatarios = [destinatarios]
    to_recipients_list = [{"emailAddress": {"address": r}} for r in destinatarios if r] # Filtra vacíos
    if not to_recipients_list:
        raise ValueError("Se requiere al menos un destinatario para reenviar.")

    payload = {"toRecipients": to_recipients_list, "comment": mensaje_reenvio}

    try:
        logger.info(f"Llamando a Graph API: POST {url}")
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status()
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
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
    try:
        logger.info(f"Llamando a Graph API: DELETE {url}")
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logger.info(f"Correo '{message_id}' eliminado por usuario '{usuario}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al eliminar el correo '{message_id}' por usuario '{usuario}': {e}. URL: {url}. Detalles: {error_details}")
        raise Exception(f"Error al eliminar el correo: {e}")


# ---- CALENDARIO ----
# (Incluye aquí las definiciones corregidas de listar_eventos, crear_evento,
#  actualizar_evento, eliminar_evento si las tienes separadas, o usa las
#  versiones ya incluidas en este bloque si este es tu único archivo)
def listar_eventos(
    top: int = 10,
    start_date: Optional[datetime] = None,
    end_date: Optional[datetime] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None,
    select: Optional[List[str]] = None,
    use_calendar_view: bool = True,
    mailbox: Optional[str] = None
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    base_endpoint = f"{BASE_URL}/users/{usuario}"
    params = {}
    endpoint_suffix = ""

    # Lógica para determinar endpoint y parámetros (CalendarView o Events)
    if use_calendar_view and start_date and end_date:
        endpoint_suffix = "/calendarView"
        if start_date.tzinfo is None: start_date = start_date.replace(tzinfo=timezone.utc)
        if end_date.tzinfo is None: end_date = end_date.replace(tzinfo=timezone.utc)
        params['startDateTime'] = start_date.isoformat()
        params['endDateTime'] = end_date.isoformat()
        params['$top'] = top
        if filter_query: params['$filter'] = filter_query
        if order_by: params['$orderby'] = order_by
        if select: params['$select'] = ','.join(select)
    else:
        endpoint_suffix = "/events"
        params['$top'] = top
        filters = []
        if start_date:
             if start_date.tzinfo is None: start_date = start_date.replace(tzinfo=timezone.utc)
             filters.append(f"start/dateTime ge '{start_date.isoformat()}'")
        if end_date:
             if end_date.tzinfo is None: end_date = end_date.replace(tzinfo=timezone.utc)
             filters.append(f"end/dateTime le '{end_date.isoformat()}'")
        if filter_query: filters.append(f"({filter_query})")

        if filters: params['$filter'] = " and ".join(filters)
        if order_by: params['$orderby'] = order_by
        if select: params['$select'] = ','.join(select)

    url = f"{base_endpoint}{endpoint_suffix}"
    clean_params = {k:v for k, v in params.items() if v is not None}

    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}")
        response = requests.get(url, headers=HEADERS, params=clean_params)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados eventos para usuario '{usuario}'. Endpoint: {endpoint_suffix}. Params: {clean_params}")
        return data
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al listar eventos para usuario '{usuario}': {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}")
        raise Exception(f"Error al listar eventos: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"❌ Error al decodificar JSON (listar eventos): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (listar eventos): {e}")

# ... (Define aquí las otras funciones: crear_evento, actualizar_evento, eliminar_evento)
# Asegúrate de que usen _actualizar_headers() y manejen errores de forma similar
def crear_evento(
    titulo: str,
    inicio: datetime,
    fin: datetime,
    asistentes: Optional[List[Dict[str, Any]]] = None,
    cuerpo: Optional[str] = None,
    es_reunion_online: bool = False,
    proveedor_reunion_online: str = "teamsForBusiness",
    recordatorio_minutos: Optional[int] = 15,
    ubicacion: Optional[str] = None,
    mostrar_como: str = "busy",
    mailbox: Optional[str] = None
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/events"

    # Validación y formato de body (similar a como estaba antes)
    if inicio.tzinfo is None: inicio = inicio.replace(tzinfo=timezone.utc)
    if fin.tzinfo is None: fin = fin.replace(tzinfo=timezone.utc)
    body = { "subject": titulo, "start": {"dateTime": inicio.isoformat(), "timeZone": "UTC"}, "end": {"dateTime": fin.isoformat(), "timeZone": "UTC"}, "showAs": mostrar_como }
    if asistentes: body["attendees"] = [{"emailAddress": {"address": a.get('emailAddress') if isinstance(a, dict) else a},"type": a.get('type', 'required') if isinstance(a, dict) else 'required'} for a in asistentes if a and (a.get('emailAddress') if isinstance(a, dict) else a)]
    if cuerpo: body["body"] = {"contentType": "HTML", "content": cuerpo}
    if ubicacion: body["location"] = {"displayName": ubicacion}
    if es_reunion_online: body["isOnlineMeeting"] = True; body["onlineMeetingProvider"] = proveedor_reunion_online
    if recordatorio_minutos is not None: body["isReminderOn"] = True; body["reminderMinutesBeforeStart"] = recordatorio_minutos
    else: body["isReminderOn"] = False

    try:
        logger.info(f"Llamando a Graph API: POST {url}")
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Evento '{titulo}' creado para usuario '{usuario}' con ID: {data.get('id')}.")
        return data
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al crear evento para usuario '{usuario}': {e}. Detalles: {error_details}. URL: {url}")
        raise Exception(f"Error al crear evento: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"❌ Error al decodificar JSON (crear evento): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (crear evento): {e}")

# ... (actualizar_evento, eliminar_evento similar) ...
def actualizar_evento(evento_id: str, nuevos_valores: dict, mailbox: Optional[str] = None) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
    # Convertir fechas a formato ISO si están presentes... (como estaba antes)
    if 'start' in nuevos_valores and isinstance(nuevos_valores['start'], datetime):
        start_dt = nuevos_valores['start']
        if start_dt.tzinfo is None: start_dt = start_dt.replace(tzinfo=timezone.utc)
        nuevos_valores['start'] = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}
    if 'end' in nuevos_valores and isinstance(nuevos_valores['end'], datetime):
        end_dt = nuevos_valores['end']
        if end_dt.tzinfo is None: end_dt = end_dt.replace(tzinfo=timezone.utc)
        nuevos_valores['end'] = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}

    try:
        logger.info(f"Llamando a Graph API: PATCH {url}")
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Evento '{evento_id}' actualizado para usuario '{usuario}'.")
        return data
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al actualizar evento '{evento_id}' para usuario '{usuario}': {e}. Detalles: {error_details}. URL: {url}")
        raise Exception(f"Error al actualizar evento: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"❌ Error al decodificar JSON (actualizar evento): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (actualizar evento): {e}")


def eliminar_evento(evento_id: str, mailbox: Optional[str] = None) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX
    url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
    try:
        logger.info(f"Llamando a Graph API: DELETE {url}")
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logger.info(f"Evento '{evento_id}' eliminado para usuario '{usuario}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al eliminar el evento '{evento_id}' para usuario '{usuario}': {e}. URL: {url}. Detalles: {error_details}")
        raise Exception(f"Error al eliminar el evento: {e}")


# ---- TEAMS y OTROS ----
# (Incluye aquí las definiciones corregidas de listar_chats, listar_equipos,
#  obtener_equipo y CUALQUIER OTRA función de acción que tengas)
def listar_chats(top: int = 20, skip: int = 0, filter_query: Optional[str] = None,
                 order_by: Optional[str] = None, expand: Optional[str] = None) -> dict:
    _actualizar_headers()
    url = f"{BASE_URL}/me/chats"
    params = {'$top': top, '$skip': skip}
    if filter_query: params['$filter'] = filter_query
    if order_by: params['$orderby'] = order_by
    if expand: params['$expand'] = expand
    clean_params = {k:v for k, v in params.items() if v is not None}
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}")
        response = requests.get(url, headers=HEADERS, params=clean_params)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} chats del usuario (me). Params: {clean_params}")
        return data
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al listar chats (/me/chats): {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}")
        raise Exception(f"Error al listar chats: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"❌ Error al decodificar JSON (listar chats): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (listar chats): {e}")


def listar_equipos(top: int = 20, skip: int = 0, filter_query: Optional[str] = None) -> dict:
    _actualizar_headers()
    url = f"{BASE_URL}/me/joinedTeams"
    params = {'$top': top, '$skip': skip}
    if filter_query: params['$filter'] = filter_query
    clean_params = {k:v for k, v in params.items() if v is not None}
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}")
        response = requests.get(url, headers=HEADERS, params=clean_params)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} equipos del usuario (me). Params: {clean_params}")
        return data
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al listar equipos (/me/joinedTeams): {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}")
        raise Exception(f"Error al listar equipos: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"❌ Error al decodificar JSON (listar equipos): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (listar equipos): {e}")


def obtener_equipo(team_id: str, select: Optional[List[str]]=None) -> dict:
    _actualizar_headers()
    url = f"{BASE_URL}/teams/{team_id}"
    params = {}
    if select: params['$select'] = ','.join(select)
    clean_params = {k:v for k, v in params.items() if v is not None}
    try:
        logger.info(f"Llamando a Graph API: GET {url} con params: {clean_params}")
        response = requests.get(url, headers=HEADERS, params=clean_params or None)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Obtenido equipo con ID: {team_id}.")
        return data
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al obtener equipo {team_id}: {e}. URL: {url}, Params: {clean_params}. Detalles: {error_details}")
        raise Exception(f"Error al obtener equipo {team_id}: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logger.error(f"❌ Error al decodificar JSON (obtener equipo): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (obtener equipo): {e}")

# Agrega aquí CUALQUIER OTRA FUNCIÓN que tengas en `acciones/` y quieras poder llamar
# Ejemplo:
# def mi_otra_funcion(param1: str, param2: int = 0) -> dict:
#     logger.info(f"Ejecutando mi_otra_funcion con param1={param1}, param2={param2}")
#     # ... tu lógica ...
#     return {"resultado": "ok", "valor": param1 * param2}

# --- FIN: Funciones Auxiliares de Graph API ---


# --- INICIO: Función Principal de Azure Functions (Entry Point) ---

# Mapeo de nombres de acción (en JSON) a las funciones Python reales
# ¡IMPORTANTE! Añade aquí todas las acciones que quieras soportar
acciones_disponibles = {
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
    # Añade aquí otras acciones que definas arriba, por ejemplo:
    # "mi_accion_personalizada": mi_otra_funcion,
}

def main(req: func.HttpRequest) -> func.HttpResponse:
    """
    Punto de entrada principal para la Azure Function HTTP Trigger.
    Espera un cuerpo JSON con 'accion' y 'parametros'.
    """
    logging.info('Python HTTP trigger function procesando una solicitud.')

    # Verificar configuración inicial (se hace al cargar el módulo, si falla aquí no llega)
    # Pero podemos añadir un chequeo extra si queremos ser muy defensivos
    try:
        assert CLIENT_ID and TENANT_ID and CLIENT_SECRET
    except Exception as config_err:
         logger.critical(f"Re-verificación de configuración falló: {config_err}")
         return func.HttpResponse("Error interno de configuración del servidor.", status_code=500)


    # Obtener acción y parámetros del cuerpo de la solicitud
    accion = None
    parametros = {}
    try:
        # Intentar obtener JSON del cuerpo
        try:
            req_body = req.get_json()
        except ValueError:
            logger.warning('No se pudo decodificar JSON del cuerpo de la solicitud.')
            return func.HttpResponse(
                 "Por favor, envíe un cuerpo JSON válido en la solicitud.",
                 status_code=400
            )

        # Obtener acción y parámetros del JSON
        if req_body:
            accion = req_body.get('accion')
            parametros = req_body.get('parametros', {}) # Default a dict vacío si no viene
        
        if not accion:
            logger.warning("No se especificó 'accion' en el cuerpo JSON.")
            return func.HttpResponse(
                 "Por favor, especifique una 'accion' en el cuerpo JSON de la solicitud.",
                 status_code=400
            )

        logger.info(f"Acción recibida: '{accion}' con parámetros: {parametros}")

        # Buscar la función correspondiente a la acción
        if accion in acciones_disponibles:
            funcion_a_ejecutar = acciones_disponibles[accion]
            
            # Ejecutar la función con los parámetros desempaquetados
            logger.info(f"Ejecutando función: {funcion_a_ejecutar.__name__}")
            resultado = funcion_a_ejecutar(**parametros) # Usa ** para pasar el dict como kwargs
            logger.info(f"Ejecución de '{accion}' completada exitosamente.")

            # Devolver el resultado como JSON
            # Asegurarse de que el resultado sea serializable a JSON
            # Si tus funciones devuelven objetos complejos, podrías necesitar un serializador custom
            try:
                 return func.HttpResponse(
                     json.dumps(resultado, default=str), # default=str para manejar fechas, etc.
                     mimetype="application/json"
                 )
            except TypeError as serialize_err:
                 logger.error(f"Error al serializar el resultado a JSON para la acción '{accion}': {serialize_err}")
                 return func.HttpResponse(f"Error interno: No se pudo serializar la respuesta para {accion}.", status_code=500)

        else:
            # Acción no encontrada en nuestro mapeo
            logger.warning(f"Acción '{accion}' no reconocida.")
            acciones_validas = list(acciones_disponibles.keys())
            return func.HttpResponse(
                 f"Acción '{accion}' no reconocida. Acciones válidas: {acciones_validas}",
                 status_code=400 # Bad Request porque la acción no existe
            )

    except ValueError as ve:
         # Errores específicos que podemos controlar (ej. de validaciones)
         logger.error(f"Error de valor durante el procesamiento: {ve}")
         return func.HttpResponse(f"Error en la solicitud: {ve}", status_code=400)

    except Exception as e:
        # Captura general para cualquier otro error inesperado durante la ejecución
        # Incluye errores al obtener token, llamadas a Graph API, etc.
        logger.exception(f"Error interno inesperado al procesar la acción '{accion}': {e}") # logger.exception incluye stack trace
        return func.HttpResponse(
             f"Error interno del servidor al procesar la acción '{accion}'. Por favor, revise los logs.",
             status_code=500
        )

# --- FIN: Función Principal de Azure Functions ---
