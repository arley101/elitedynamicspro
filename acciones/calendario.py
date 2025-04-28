import logging
import os
import requests
# Asume que auth.py está en la raíz
from auth import obtener_token
# CORRECCION: Añadir Any y json
from typing import Dict, List, Optional, Union, Any
from datetime import datetime, timezone
import json

# Configuración básica de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- INICIO: Configuración Redundante ---
# (Se mantiene por ahora, pero idealmente se eliminaría)
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
GRAPH_SCOPE = os.getenv('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')
if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE]):
    logging.error("❌ Faltan variables de entorno en calendario.py.")
# BASE_URL se define localmente como /me/events, podría causar confusión.
# Sería mejor usar un BASE_URL_GRAPH global y construir la URL completa aquí.
# Por ahora lo dejamos, pero es un punto de mejora.
BASE_URL = "https://graph.microsoft.com/v1.0/me" # Cambiado a /me, se añade /events o /calendarView después
HEADERS: Dict[str, Optional[str]] = {'Authorization': None, 'Content-Type': 'application/json'}
# MAILBOX no se usa en este archivo si BASE_URL es /me

def _actualizar_headers() -> None:
    """Obtiene un nuevo token de acceso y actualiza el diccionario HEADERS local."""
    try:
        token = obtener_token(); HEADERS['Authorization'] = f'Bearer {token}'; logging.info("Headers actualizados en calendario.py.")
    except Exception as e: logging.error(f"❌ Error token en calendario.py: {e}"); raise Exception(f"Error token en calendario.py: {e}")
# --- FIN: Configuración Redundante ---


# ---- FUNCIONES DE GESTIÓN DE CALENDARIO DE OUTLOOK ----
def listar_eventos(
    top: int = 10,
    start_date: Optional[datetime] = None,
    end_date: Optional[datetime] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None,
    select: Optional[List[str]] = None,
    use_calendar_view: bool = True, # Añadido para que coincida con HttpTrigger
    mailbox: Optional[str] = 'me' # Añadido para consistencia, default a 'me'
) -> Dict[str, Any]:
    """
    Lista eventos del calendario de Outlook... (Docstring como estaba)
    """
    _actualizar_headers()
    # Determinar el usuario/mailbox a usar
    usuario = mailbox if mailbox else 'me' # Default a 'me' si no se especifica
    base_endpoint = f"https://graph.microsoft.com/v1.0/users/{usuario}" # URL base correcta

    params: Dict[str, Any] = {}
    endpoint_suffix = ""

    # Lógica para determinar endpoint y parámetros (CalendarView o Events)
    # (Misma lógica que pusimos en HttpTrigger/__init__.py)
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
        response = requests.get(url, headers=HEADERS, params=clean_params)
        response.raise_for_status()
        data: Dict[str, Any] = response.json()
        logging.info(f"Listados eventos del calendario para '{usuario}'.")
        return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error listar eventos: {e}. Detalles: {error_details}"); raise Exception(f"Error al listar eventos: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (listar eventos): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar eventos): {e}")


def crear_evento(
    titulo: str,
    inicio: datetime,
    fin: datetime,
    asistentes: Optional[List[Dict[str, Any]]] = None, # Ajustado tipo Any
    cuerpo: Optional[str] = None,
    es_reunion_online: bool = False,
    proveedor_reunion_online: str = "teamsForBusiness",
    recordatorio_minutos: Optional[int] = None,
    ubicacion: Optional[str] = None, # Añadido para completar
    mostrar_como: str = "busy", # Añadido para completar
    mailbox: Optional[str] = 'me' # Añadido para consistencia
) -> Dict[str, Any]:
    """
    Crea un nuevo evento u reunión en el calendario de Outlook.
    """
    _actualizar_headers()
    usuario = mailbox if mailbox else 'me'
    url = f"https://graph.microsoft.com/v1.0/users/{usuario}/events" # URL correcta

    # Asegurar datetime y UTC
    if not isinstance(inicio, datetime) or not isinstance(fin, datetime): raise ValueError("'inicio' y 'fin' deben ser datetimes.")
    if inicio.tzinfo is None: inicio = inicio.replace(tzinfo=timezone.utc)
    if fin.tzinfo is None: fin = fin.replace(tzinfo=timezone.utc)

    # --- CORRECCIÓN PRINCIPAL: Tipado explícito de 'body' ---
    body: Dict[str, Any] = {
        "subject": titulo,
        "start": {"dateTime": inicio.isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": fin.isoformat(), "timeZone": "UTC"},
    }
    # --- FIN CORRECCIÓN PRINCIPAL ---

    # Ahora las asignaciones a 'body' deberían ser aceptadas por MyPy
    if asistentes is not None:
         if isinstance(asistentes, list) and all(isinstance(a, dict) for a in asistentes):
              # Línea ~123: Esta asignación ahora debería ser válida para MyPy
              body["attendees"] = [{"emailAddress": {"address": a.get('emailAddress')},"type": a.get('type', 'required')} for a in asistentes if a and a.get('emailAddress')]
         else: logging.warning(f"Tipo/Formato inesperado para 'asistentes': {type(asistentes)}")

    if cuerpo is not None and isinstance(cuerpo, str): body["body"] = {"contentType": "HTML", "content": cuerpo}
    if ubicacion is not None and isinstance(ubicacion, str): body["location"] = {"displayName": ubicacion}

    # Línea ~127: Esta asignación ahora debería ser válida
    if es_reunion_online: body["isOnlineMeeting"] = es_reunion_online # Asigna bool a Any
    if proveedor_reunion_online: body["onlineMeetingProvider"] = proveedor_reunion_online # Asigna str a Any

    # Línea ~130: Esta asignación ahora debería ser válida
    if recordatorio_minutos is not None and isinstance(recordatorio_minutos, int):
        body["isReminderOn"] = True
        body["reminderMinutesBeforeStart"] = recordatorio_minutos # Asigna int a Any
    else: body["isReminderOn"] = False # Asigna bool a Any

    if mostrar_como: body["showAs"] = mostrar_como # Asigna str a Any

    response: Optional[requests.Response] = None
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data: Dict[str, Any] = response.json()
        logging.info(f"Evento '{titulo}' creado para '{usuario}'.")
        return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error crear evento: {e}. Detalles: {error_details}"); raise Exception(f"Error al crear evento: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (crear evento): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (crear evento): {e}")


# (Función actualizar_evento - necesita corrección similar a la de HttpTrigger)
def actualizar_evento(evento_id: str, nuevos_valores: Dict[str, Any], mailbox: Optional[str] = 'me') -> Dict[str, Any]:
    """Actualiza un evento existente en el calendario de Outlook."""
    _actualizar_headers()
    usuario = mailbox if mailbox else 'me'
    url = f"https://graph.microsoft.com/v1.0/users/{usuario}/events/{evento_id}" # URL correcta

    payload = nuevos_valores.copy()
    # --- CORRECCIÓN: Procesar fechas en múltiples líneas ---
    if 'start' in payload and isinstance(payload.get('start'), datetime):
        start_dt = payload['start']
        if start_dt.tzinfo is None: start_dt = start_dt.replace(tzinfo=timezone.utc)
        payload['start'] = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}
    if 'end' in payload and isinstance(payload.get('end'), datetime):
        end_dt = payload['end']
        if end_dt.tzinfo is None: end_dt = end_dt.replace(tzinfo=timezone.utc)
        payload['end'] = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}
    # --- FIN CORRECCIÓN ---

    response: Optional[requests.Response] = None
    try:
        # Añadir ETag si es necesario (similar a actualizar_plan)
        etag = payload.pop('@odata.etag', None)
        current_headers = HEADERS.copy()
        if etag: current_headers['If-Match'] = etag; logging.info(f"Usando ETag para evento {evento_id}")

        response = requests.patch(url, headers=current_headers, json=payload)
        response.raise_for_status()
        logging.info(f"Evento '{evento_id}' actualizado para '{usuario}'.")
        if response.status_code == 204: return {"status": "Actualizado (No Content)", "id": evento_id} # No hay cuerpo
        else: return response.json()
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error actualizar evento '{evento_id}': {e}. Detalles: {error_details}"); raise Exception(f"Error actualizar evento: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (actualizar evento): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (actualizar evento): {e}")


def eliminar_evento(evento_id: str, mailbox: Optional[str] = 'me') -> Dict[str, Any]:
    """Elimina un evento del calendario de Outlook."""
    _actualizar_headers()
    usuario = mailbox if mailbox else 'me'
    url = f"https://graph.microsoft.com/v1.0/users/{usuario}/events/{evento_id}" # URL correcta
    response: Optional[requests.Response] = None
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status() # Espera 204
        logging.info(f"Evento '{evento_id}' eliminado para '{usuario}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error eliminar evento '{evento_id}': {e}. Detalles: {error_details}"); raise Exception(f"Error al eliminar evento: {e}")

# (Función crear_reunion_teams - sin cambios)
def crear_reunion_teams(
    titulo: str, inicio: datetime, fin: datetime,
    asistentes: Optional[List[Dict[str, Any]]] = None, # Ajustado tipo
    cuerpo: Optional[str] = None,
    mailbox: Optional[str] = 'me' # Añadido mailbox
) -> dict:
    """Crea una reunión de Teams (evento online en calendario)."""
    logging.info(f"Creando reunión Teams '{titulo}' para '{mailbox if mailbox else 'me'}'")
    # Pasa todos los argumentos relevantes a crear_evento
    return crear_evento(
        titulo=titulo, inicio=inicio, fin=fin, asistentes=asistentes, cuerpo=cuerpo,
        es_reunion_online=True, proveedor_reunion_online="teamsForBusiness",
        mailbox=mailbox
        # Pasar otros args si se añadieran a crear_reunion_teams
    )
