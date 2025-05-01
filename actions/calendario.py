# actions/calendario.py (Refactorizado)

import logging
import requests
import os # Necesario para leer variables específicas si se mantienen
from typing import Dict, List, Optional, Union, Any
from datetime import datetime, timezone
import json

# Usar logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar constantes globales desde __init__.py
try:
    from .. import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    # Fallback por si se ejecuta standalone
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_TIMEOUT = 45
    logger.warning("No se pudo importar BASE_URL/GRAPH_API_TIMEOUT desde el padre, usando defaults.")

# (Eliminada configuración redundante: CLIENT_ID, SECRET, SCOPE, HEADERS locales, _actualizar_headers)
# MAILBOX ya no se usa aquí, se usará /me implícito en el token delegado

# ---- FUNCIONES DE GESTIÓN DE CALENDARIO DE OUTLOOK ----
# Todas las funciones ahora aceptan 'headers' como primer argumento

def listar_eventos(
    headers: Dict[str, str],
    top: int = 10,
    start_date: Optional[datetime] = None,
    end_date: Optional[datetime] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None,
    select: Optional[List[str]] = None,
    use_calendar_view: bool = True,
    mailbox: Optional[str] = 'me' # 'me' es el default para usar con token delegado
) -> Dict[str, Any]:
    """Lista eventos del calendario de Outlook (/me o /users/{mailbox} si se especifica)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    usuario = mailbox if mailbox and mailbox != 'me' else 'me' # Default a /me
    base_endpoint = f"{BASE_URL}/users/{usuario}" if usuario != 'me' else f"{BASE_URL}/me"
    params: Dict[str, Any] = {}
    endpoint_suffix = ""

    def ensure_timezone(dt: Optional[datetime]) -> Optional[datetime]: return dt.replace(tzinfo=timezone.utc) if dt and isinstance(dt, datetime) and dt.tzinfo is None else dt
    start_date = ensure_timezone(start_date); end_date = ensure_timezone(end_date)

    if use_calendar_view and start_date and end_date:
        endpoint_suffix = "/calendarView"
        if isinstance(start_date, datetime): params['startDateTime'] = start_date.isoformat()
        if isinstance(end_date, datetime): params['endDateTime'] = end_date.isoformat()
        params['$top'] = int(top)
        if filter_query: params['$filter'] = filter_query
        if order_by: params['$orderby'] = order_by
        if select: params['$select'] = ','.join(select)
    else:
        endpoint_suffix = "/events"
        params['$top'] = int(top); filters = []
        if start_date: filters.append(f"start/dateTime ge '{start_date.isoformat()}'")
        if end_date: filters.append(f"end/dateTime le '{end_date.isoformat()}'")
        if filter_query: filters.append(f"({filter_query})")
        if filters: params['$filter'] = " and ".join(filters)
        if order_by: params['$orderby'] = order_by
        if select: params['$select'] = ','.join(select)

    url = f"{base_endpoint}{endpoint_suffix}"
    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} Params: {clean_params}")
        response = requests.get(url, headers=headers, params=clean_params, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data: Dict[str, Any] = response.json()
        logger.info(f"Listados eventos del calendario para '{usuario}'.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en listar_eventos: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_eventos: {e}", exc_info=True)
        raise


def crear_evento(
    headers: Dict[str, str],
    titulo: str,
    inicio: datetime,
    fin: datetime,
    asistentes: Optional[List[Dict[str, Any]]] = None,
    cuerpo: Optional[str] = None,
    es_reunion_online: bool = False,
    proveedor_reunion_online: str = "teamsForBusiness",
    recordatorio_minutos: Optional[int] = None,
    ubicacion: Optional[str] = None,
    mostrar_como: str = "busy",
    mailbox: Optional[str] = 'me'
) -> Dict[str, Any]:
    """Crea un nuevo evento en el calendario (/me o /users/{mailbox})."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    usuario = mailbox if mailbox and mailbox != 'me' else 'me'
    url = f"{BASE_URL}/users/{usuario}/events" if usuario != 'me' else f"{BASE_URL}/me/events"

    if not isinstance(inicio, datetime) or not isinstance(fin, datetime): raise ValueError("'inicio' y 'fin' deben ser datetimes.")
    if inicio.tzinfo is None: inicio = inicio.replace(tzinfo=timezone.utc)
    if fin.tzinfo is None: fin = fin.replace(tzinfo=timezone.utc)

    body: Dict[str, Any] = {
        "subject": titulo,
        "start": {"dateTime": inicio.isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": fin.isoformat(), "timeZone": "UTC"},
    }
    if asistentes is not None:
         if isinstance(asistentes, list) and all(isinstance(a, dict) for a in asistentes):
              body["attendees"] = [{"emailAddress": {"address": a.get('emailAddress')},"type": a.get('type', 'required')} for a in asistentes if a and a.get('emailAddress')]
         else: logger.warning(f"Tipo/Formato inesperado para 'asistentes': {type(asistentes)}")
    if cuerpo: body["body"] = {"contentType": "HTML", "content": cuerpo}
    if ubicacion: body["location"] = {"displayName": ubicacion}
    if es_reunion_online: body["isOnlineMeeting"] = es_reunion_online
    if es_reunion_online and proveedor_reunion_online: body["onlineMeetingProvider"] = proveedor_reunion_online
    if recordatorio_minutos is not None: body["isReminderOn"] = True; body["reminderMinutesBeforeStart"] = recordatorio_minutos
    else: body["isReminderOn"] = False
    if mostrar_como: body["showAs"] = mostrar_como

    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando evento '{titulo}' para '{usuario}')")
        response = requests.post(url, headers=headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data: Dict[str, Any] = response.json()
        logger.info(f"Evento '{titulo}' creado para '{usuario}'. ID: {data.get('id')}")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en crear_evento: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_evento: {e}", exc_info=True)
        raise

def actualizar_evento(
    headers: Dict[str, str],
    evento_id: str,
    nuevos_valores: Dict[str, Any],
    mailbox: Optional[str] = 'me'
) -> Dict[str, Any]:
    """Actualiza un evento existente (/me o /users/{mailbox})."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    usuario = mailbox if mailbox and mailbox != 'me' else 'me'
    url = f"{BASE_URL}/users/{usuario}/events/{evento_id}" if usuario != 'me' else f"{BASE_URL}/me/events/{evento_id}"

    payload = nuevos_valores.copy()
    if 'start' in payload and isinstance(payload.get('start'), datetime):
        start_dt = payload['start']
        if start_dt.tzinfo is None: start_dt = start_dt.replace(tzinfo=timezone.utc)
        payload['start'] = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}
    if 'end' in payload and isinstance(payload.get('end'), datetime):
        end_dt = payload['end']
        if end_dt.tzinfo is None: end_dt = end_dt.replace(tzinfo=timezone.utc)
        payload['end'] = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}

    response: Optional[requests.Response] = None
    try:
        etag = payload.pop('@odata.etag', None)
        current_headers = headers.copy()
        if etag: current_headers['If-Match'] = etag; logger.info(f"Usando ETag para evento {evento_id}")

        logger.info(f"API Call: PATCH {url} (Actualizando evento '{evento_id}' para '{usuario}')")
        response = requests.patch(url, headers=current_headers, json=payload, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        logger.info(f"Evento '{evento_id}' actualizado para '{usuario}'.")
        if response.status_code == 204: return {"status": "Actualizado (No Content)", "id": evento_id}
        else: return response.json()
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en actualizar_evento: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en actualizar_evento: {e}", exc_info=True)
        raise


def eliminar_evento(
    headers: Dict[str, str],
    evento_id: str,
    mailbox: Optional[str] = 'me'
) -> Dict[str, Any]:
    """Elimina un evento del calendario (/me o /users/{mailbox})."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    usuario = mailbox if mailbox and mailbox != 'me' else 'me'
    url = f"{BASE_URL}/users/{usuario}/events/{evento_id}" if usuario != 'me' else f"{BASE_URL}/me/events/{evento_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando evento '{evento_id}' para '{usuario}')")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 204
        logger.info(f"Evento '{evento_id}' eliminado para '{usuario}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en eliminar_evento: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_evento: {e}", exc_info=True)
        raise

def crear_reunion_teams(
    headers: Dict[str, str],
    titulo: str, inicio: datetime, fin: datetime,
    asistentes: Optional[List[Dict[str, Any]]] = None,
    cuerpo: Optional[str] = None,
    mailbox: Optional[str] = 'me'
) -> dict:
    """Crea una reunión de Teams (evento online en calendario)."""
    logger.info(f"Creando reunión Teams '{titulo}' para '{mailbox if mailbox else 'me'}'")
    # Llama a la función crear_evento refactorizada, pasando los headers
    return crear_evento(
        headers=headers, # Pasar los headers recibidos
        titulo=titulo, inicio=inicio, fin=fin, asistentes=asistentes, cuerpo=cuerpo,
        es_reunion_online=True, proveedor_reunion_online="teamsForBusiness",
        mailbox=mailbox
    )
