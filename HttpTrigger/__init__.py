import json
import logging
import requests
import azure.functions as func
from typing import Dict, Any, Callable, List, Optional, Union
from datetime import datetime, timezone
import os

# Configuración de logging
logger = logging.getLogger("azure.functions")
logger.setLevel(logging.INFO)

# Variables de entorno
def get_config_or_raise(key: str, default: Optional[str] = None) -> str:
    value = os.environ.get(key, default)
    if value is None: # Chequea None explícitamente porque 'me' es un valor válido
        logger.error(f"Falta la variable de entorno requerida: {key}")
        raise ValueError(f"Falta la variable de entorno: {key}")
    return value

try:
    CLIENT_ID = get_config_or_raise('CLIENT_ID')
    TENANT_ID = get_config_or_raise('TENANT_ID')
    CLIENT_SECRET = get_config_or_raise('CLIENT_SECRET')
    # CORRECCIÓN: Establecer 'me' como default si no se especifica MAILBOX
    MAILBOX = get_config_or_raise('MAILBOX', default='me')
    GRAPH_SCOPE = os.environ.get('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')
except ValueError as e:
    logger.critical(f"Error de configuración inicial: {e}. La función no puede continuar.")
    raise

# Constantes y Headers globales
BASE_URL = "https://graph.microsoft.com/v1.0"
HEADERS = {
    'Authorization': None,
    'Content-Type': 'application/json'
}

# --- Funciones de Autenticación y Headers (sin cambios) ---
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
        logger.error(f"❌ Error de red/HTTP al obtener el token: {e}")
        raise Exception(f"Error de red/HTTP al obtener el token: {e}")
    except json.JSONDecodeError as e:
        logger.error(f"❌ Error al decodificar la respuesta JSON del token: {e}. Respuesta: {response.text}")
        raise Exception(f"Error al decodificar la respuesta JSON del token: {e}")
    except Exception as e:
        logger.error(f"❌ Error inesperado al obtener el token: {e}")
        raise

def _actualizar_headers() -> None:
    try:
        token = obtener_token()
        HEADERS['Authorization'] = f'Bearer {token}'
        logger.info("Cabecera de autorización actualizada.")
    except Exception as e:
        logger.error(f"❌ Falló la actualización de la cabecera de autorización.")
        raise Exception(f"Fallo al actualizar la cabecera: {e}")

# ---- CORREO (Adaptado) ----
def listar_correos(
    top: int = 10,
    skip: int = 0,
    folder: str = 'Inbox',
    select: Optional[List[str]] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None,
    mailbox: Optional[str] = None # <--- Parámetro añadido
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX # <--- Lógica de selección
    url = f"{BASE_URL}/users/{usuario}/mailFolders/{folder}/messages?$top={top}&$skip={skip}"
    # ... (resto del código igual, usando la url construida)
    params = {'$top': top, '$skip': skip}
    if select: params['$select'] = ','.join(select)
    if filter_query: params['$filter'] = filter_query
    if order_by: params['$orderby'] = order_by

    try:
        response = requests.get(url, headers=HEADERS, params={k:v for k, v in params.items() if v is not None}) # Limpia Nones de params
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} correos de la carpeta '{folder}' para usuario '{usuario}'.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Error al listar correos para '{usuario}' en carpeta '{folder}': {e}. URL: {url}")
        raise Exception(f"Error al listar correos: {e}")


def leer_correo(
    message_id: str,
    select: Optional[List[str]] = None,
    mailbox: Optional[str] = None # <--- Parámetro añadido
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX # <--- Lógica de selección
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
    # ... (resto del código igual)
    params = {}
    if select: params['$select'] = ','.join(select)

    try:
        response = requests.get(url, headers=HEADERS, params=params or None)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Correo '{message_id}' leído para usuario '{usuario}'.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Error al leer el correo '{message_id}' para usuario '{usuario}': {e}. URL: {url}")
        raise Exception(f"Error al leer el correo: {e}")


def enviar_correo(
    destinatario: Union[str, List[str]],
    asunto: str,
    mensaje: str,
    cc: Optional[Union[str, List[str]]] = None,
    bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[dict]] = None,
    # from_email: Optional[str] = None, # 'from' es complejo con app permissions, requiere config específica
    is_draft: bool = False,
    mailbox: Optional[str] = None # <--- Parámetro añadido
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX # <--- Lógica de selección
    endpoint = "messages" if is_draft else "sendMail"
    url = f"{BASE_URL}/users/{usuario}/{endpoint}"
    # ... (resto del código igual, construyendo payload)
    if isinstance(destinatario, str): destinatario = [destinatario]
    if isinstance(cc, str): cc = [cc]
    if isinstance(bcc, str): bcc = [bcc]

    to_recipients = [{"emailAddress": {"address": r}} for r in destinatario]
    cc_recipients = [{"emailAddress": {"address": r}} for r in cc] if cc else []
    bcc_recipients = [{"emailAddress": {"address": r}} for r in bcc] if bcc else []

    message_payload = {
        "subject": asunto,
        "body": {"contentType": "HTML", "content": mensaje},
        "toRecipients": to_recipients,
    }
    if cc_recipients: message_payload["ccRecipients"] = cc_recipients
    if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
    if attachments: message_payload["attachments"] = attachments
    # if from_email: message_payload["from"] = {"emailAddress": {"address": from_email}}

    payload = {"message": message_payload, "saveToSentItems": "true"} if not is_draft else message_payload

    try:
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status()
        if not is_draft:
            logger.info(f"Correo enviado por '{usuario}' con asunto '{asunto}'.")
            return {"status": "Enviado", "code": response.status_code}
        else:
            data = response.json()
            message_id = data.get('id')
            logger.info(f"Correo guardado como borrador por '{usuario}' con ID: {message_id}.")
            return {"status": "Borrador Guardado", "code": response.status_code, "id": message_id, "data": data}
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al {'enviar' if not is_draft else 'guardar borrador'} correo por '{usuario}': {e}. Detalles: {error_details}. URL: {url}")
        raise Exception(f"Error al {'enviar' if not is_draft else 'guardar borrador'} correo: {e}")


def guardar_borrador(
    destinatario: Union[str, List[str]],
    asunto: str,
    mensaje: str,
    cc: Optional[Union[str, List[str]]] = None,
    bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[dict]] = None,
    # from_email: Optional[str] = None,
    mailbox: Optional[str] = None # <--- Parámetro añadido
) -> dict:
    logger.info(f"Guardando borrador para usuario '{mailbox or MAILBOX}' con asunto: '{asunto}'")
    # Pasar el mailbox a enviar_correo
    return enviar_correo(destinatario, asunto, mensaje, cc, bcc, attachments, is_draft=True, mailbox=mailbox)


def enviar_borrador(
    message_id: str,
    mailbox: Optional[str] = None # <--- Parámetro añadido
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX # <--- Lógica de selección
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/send"
    # ... (resto del código igual)
    try:
        response = requests.post(url, headers=HEADERS)
        response.raise_for_status()
        logger.info(f"Borrador de correo '{message_id}' enviado por usuario '{usuario}'.")
        return {"status": "Borrador Enviado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Error al enviar borrador '{message_id}' por usuario '{usuario}': {e}. URL: {url}")
        raise Exception(f"Error al enviar borrador: {e}")


def responder_correo(
    message_id: str,
    mensaje_respuesta: str,
    to_recipients: Optional[List[dict]] = None, # Mantener opcional para sobreescribir
    reply_all: bool = False, # Añadir opción para replyAll
    mailbox: Optional[str] = None # <--- Parámetro añadido
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX # <--- Lógica de selección
    action = "replyAll" if reply_all else "reply"
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/{action}"
    payload = {"comment": mensaje_respuesta}
    # Si se proporcionan to_recipients, hay que incluirlos en una estructura 'message'
    # if to_recipients:
    #    payload["message"] = {"toRecipients": to_recipients}

    try:
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status()
        logger.info(f"Respuesta {'a todos ' if reply_all else ''}enviada al correo '{message_id}' por usuario '{usuario}'.")
        return {"status": "Respondido", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Error al responder {'a todos ' if reply_all else ''}al correo '{message_id}' por usuario '{usuario}': {e}. URL: {url}")
        raise Exception(f"Error al responder al correo: {e}")


def reenviar_correo(
    message_id: str,
    destinatarios: Union[str, List[str]],
    mensaje_reenvio: str = "FYI",
    mailbox: Optional[str] = None # <--- Parámetro añadido
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX # <--- Lógica de selección
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/forward"
    # ... (resto del código igual, construyendo payload)
    if isinstance(destinatarios, str): destinatarios = [destinatarios]
    to_recipients_list = [{"emailAddress": {"address": r}} for r in destinatarios]
    payload = {"toRecipients": to_recipients_list, "comment": mensaje_reenvio}

    try:
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status()
        logger.info(f"Correo '{message_id}' reenviado por usuario '{usuario}' a: {destinatarios}.")
        return {"status": "Reenviado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Error al reenviar el correo '{message_id}' por usuario '{usuario}': {e}. URL: {url}")
        raise Exception(f"Error al reenviar el correo: {e}")


def eliminar_correo(
    message_id: str,
    mailbox: Optional[str] = None # <--- Parámetro añadido
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX # <--- Lógica de selección
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
    # ... (resto del código igual)
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logger.info(f"Correo '{message_id}' eliminado por usuario '{usuario}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Error al eliminar el correo '{message_id}' por usuario '{usuario}': {e}. URL: {url}")
        raise Exception(f"Error al eliminar el correo: {e}")


# ---- CALENDARIO (Adaptado) ----
def listar_eventos(
    top: int = 10,
    start_date: Optional[datetime] = None,
    end_date: Optional[datetime] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None,
    select: Optional[List[str]] = None,
    use_calendar_view: bool = True,
    mailbox: Optional[str] = None # <--- Parámetro añadido
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX # <--- Lógica de selección
    base_endpoint = f"{BASE_URL}/users/{usuario}"
    params = {}
    endpoint_suffix = ""

    if use_calendar_view and start_date and end_date:
        endpoint_suffix = "/calendarView"
        # Asegurar UTC para calendarView
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
        if filter_query: filters.append(f"({filter_query})") # Paréntesis por seguridad

        if filters: params['$filter'] = " and ".join(filters)
        if order_by: params['$orderby'] = order_by
        if select: params['$select'] = ','.join(select)

    url = f"{base_endpoint}{endpoint_suffix}"

    try:
        # Limpia params de valores None antes de la llamada
        clean_params = {k:v for k, v in params.items() if v is not None}
        response = requests.get(url, headers=HEADERS, params=clean_params)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados eventos para usuario '{usuario}'. Endpoint: {endpoint_suffix}. Params: {clean_params}")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Error al listar eventos para usuario '{usuario}': {e}. URL: {url}, Params: {clean_params}")
        raise Exception(f"Error al listar eventos: {e}")


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
    mailbox: Optional[str] = None # <--- Parámetro añadido
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX # <--- Lógica de selección
    url = f"{BASE_URL}/users/{usuario}/events"
    # ... (resto del código igual, construyendo body)
    if inicio.tzinfo is None: inicio = inicio.replace(tzinfo=timezone.utc)
    if fin.tzinfo is None: fin = fin.replace(tzinfo=timezone.utc)

    body = {
        "subject": titulo,
        "start": {"dateTime": inicio.isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": fin.isoformat(), "timeZone": "UTC"},
        "showAs": mostrar_como
    }
    if asistentes:
        body["attendees"] = [
            {"emailAddress": {"address": a.get('emailAddress') if isinstance(a, dict) else a},
             "type": a.get('type', 'required') if isinstance(a, dict) else 'required'}
            for a in asistentes if a and (a.get('emailAddress') if isinstance(a, dict) else a) # Filtra Nones/vacíos
        ]
    if cuerpo: body["body"] = {"contentType": "HTML", "content": cuerpo}
    if ubicacion: body["location"] = {"displayName": ubicacion}
    if es_reunion_online:
        body["isOnlineMeeting"] = True
        body["onlineMeetingProvider"] = proveedor_reunion_online
    if recordatorio_minutos is not None:
        body["isReminderOn"] = True
        body["reminderMinutesBeforeStart"] = recordatorio_minutos
    else:
        body["isReminderOn"] = False

    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Evento '{titulo}' creado para usuario '{usuario}' con ID: {data.get('id')}.")
        return data
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al crear evento para usuario '{usuario}': {e}. Detalles: {error_details}. URL: {url}")
        raise Exception(f"Error al crear evento: {e}")


def actualizar_evento(
    evento_id: str,
    nuevos_valores: dict,
    mailbox: Optional[str] = None # <--- Parámetro añadido
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX # <--- Lógica de selección
    url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
    # ... (resto del código igual, procesando nuevos_valores y haciendo PATCH)
    if 'start' in nuevos_valores and isinstance(nuevos_valores['start'], datetime):
        start_dt = nuevos_valores['start']
        if start_dt.tzinfo is None: start_dt = start_dt.replace(tzinfo=timezone.utc)
        nuevos_valores['start'] = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}
    if 'end' in nuevos_valores and isinstance(nuevos_valores['end'], datetime):
        end_dt = nuevos_valores['end']
        if end_dt.tzinfo is None: end_dt = end_dt.replace(tzinfo=timezone.utc)
        nuevos_valores['end'] = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}

    try:
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Evento '{evento_id}' actualizado para usuario '{usuario}'.")
        return data
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logger.error(f"❌ Error al actualizar evento '{evento_id}' para usuario '{usuario}': {e}. Detalles: {error_details}. URL: {url}")
        raise Exception(f"Error al actualizar evento: {e}")


def eliminar_evento(
    evento_id: str,
    mailbox: Optional[str] = None # <--- Parámetro añadido
) -> dict:
    _actualizar_headers()
    usuario = mailbox or MAILBOX # <--- Lógica de selección
    url = f"{BASE_URL}/users/{usuario}/events/{evento_id}"
    # ... (resto del código igual, haciendo DELETE)
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logger.info(f"Evento '{evento_id}' eliminado para usuario '{usuario}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Error al eliminar el evento '{evento_id}' para usuario '{usuario}': {e}. URL: {url}")
        raise Exception(f"Error al eliminar el evento: {e}")

# ---- TEAMS (Chats - Generalmente no dependen de 'usuario' específico en la URL base, usan /chats o /me/chats) ----
# Las funciones de chat como listar_chats, obtener_chat, crear_chat, enviar_mensaje_chat, etc.,
# generalmente usan /me/chats (para el usuario autenticado) o /chats (para operaciones a nivel de tenant o con IDs).
# Por lo tanto, el parámetro 'mailbox' no suele ser aplicable aquí de la misma manera.
# Mantendremos estas funciones como estaban, asumiendo que operan en el contexto del usuario/app autenticado.

# ... (Funciones de Teams: listar_chats, obtener_chat, crear_chat, enviar_mensaje_chat, etc. se mantienen como antes) ...
# Ejemplo: listar_chats sigue usando /me/chats
def listar_chats(top: int = 20, skip: int = 0, filter_query: Optional[str] = None,
                 order_by: Optional[str] = None, expand: Optional[str] = None) -> dict:
    _actualizar_headers()
    # Usa /me/, no depende del parámetro 'mailbox'
    url = f"{BASE_URL}/me/chats"
    params = {'$top': top, '$skip': skip}
    if filter_query: params['$filter'] = filter_query
    if order_by: params['$orderby'] = order_by
    if expand: params['$expand'] = expand

    try:
        response = requests.get(url, headers=HEADERS, params=params)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} chats del usuario (me). Params: {params}")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Error al listar chats (/me/chats): {e}. URL: {url}, Params: {params}")
        raise Exception(f"Error al listar chats: {e}")


# ---- EQUIPOS Y CANALES (Adaptado donde aplica) ----
# Listar equipos unidos por el usuario actual usa /me/, no necesita 'mailbox'.
def listar_equipos(top: int = 20, skip: int = 0, filter_query: Optional[str] = None) -> dict:
    _actualizar_headers()
    # Usa /me/, no depende del parámetro 'mailbox'
    url = f"{BASE_URL}/me/joinedTeams"
    params = {'$top': top, '$skip': skip}
    if filter_query: params['$filter'] = filter_query

    try:
        response = requests.get(url, headers=HEADERS, params=params)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} equipos del usuario (me). Params: {params}")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Error al listar equipos (/me/joinedTeams): {e}. URL: {url}, Params: {params}")
        raise Exception(f"Error al listar equipos: {e}")

# Obtener equipo por ID no depende de un usuario específico en la URL.
def obtener_equipo(team_id: str, select: Optional[List[str]]=None) -> dict:
    _actualizar_headers()
    # No usa /users/ o /me/ en la URL base del recurso.
    url = f"{BASE_URL}/teams/{team_id}"
    params = {}
    if select: params['$select'] = ','.join(select)

    try:
        response = requests.get(url, headers=HEADERS, params=params or None)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Obtenido equipo con ID: {team_id}.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Error al obtener equipo {team_id}: {e}. URL: {url}, Params: {params}")
        raise Exception(f"Error al obtener equipo {team_id}: {e}")
