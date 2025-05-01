# actions/calendario.py (Refactorizado v2 con Helper)

import logging
import requests # Solo para tipos de excepción y paginación
import json
from typing import Dict, List, Optional, Union, Any
from datetime import datetime, timezone

# Usar el logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar helper y constantes
try:
    from helpers.http_client import hacer_llamada_api
    from shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    logger.error("Error importando helpers/constantes en Calendario.")
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs): raise NotImplementedError("Helper no importado")

# ---- FUNCIONES DE CALENDARIO ----
# Usan el helper hacer_llamada_api

def _ensure_timezone(dt: Optional[datetime]) -> Optional[datetime]:
    if dt and isinstance(dt, datetime) and dt.tzinfo is None: return dt.replace(tzinfo=timezone.utc)
    return dt

def listar_eventos(headers: Dict[str, str], top: int = 10, start_date: Optional[datetime] = None, end_date: Optional[datetime] = None, filter_query: Optional[str] = None, order_by: Optional[str] = None, select: Optional[List[str]] = None, use_calendar_view: bool = True, mailbox: str = 'me') -> Dict[str, Any]:
    """Lista eventos del calendario usando helper o requests directo para paginación."""
    base_endpoint = f"{BASE_URL}/users/{mailbox}"
    params: Dict[str, Any] = {}; endpoint_suffix: str = ""
    start_date_tz = _ensure_timezone(start_date); end_date_tz = _ensure_timezone(end_date)

    if use_calendar_view and start_date_tz and end_date_tz:
        endpoint_suffix = "/calendarView"; assert start_date_tz is not None; assert end_date_tz is not None
        params['startDateTime'] = start_date_tz.isoformat(); params['endDateTime'] = end_date_tz.isoformat()
        params['$top'] = int(top)
        if filter_query: params['$filter'] = filter_query
        if order_by: params['$orderby'] = order_by
        if select: params['$select'] = ','.join(select)
    else:
        endpoint_suffix = "/events"; params['$top'] = int(top); filters = []
        if start_date_tz: filters.append(f"start/dateTime ge '{start_date_tz.isoformat()}'")
        if end_date_tz: filters.append(f"end/dateTime le '{end_date_tz.isoformat()}'")
        if filter_query: filters.append(f"({filter_query})")
        if filters: params['$filter'] = " and ".join(filters)
        if order_by: params['$orderby'] = order_by
        if select: params['$select'] = ','.join(select)

    url_base = f"{base_endpoint}{endpoint_suffix}"
    clean_params = {k:v for k, v in params.items() if v is not None}

    # Si solo queremos el top N, usamos el helper. Si top es grande o queremos paginar, usamos requests directo.
    # Por simplicidad ahora, siempre usamos requests directo para manejar la paginación si existiera @odata.nextLink
    all_events: List[Dict[str, Any]] = []
    current_url: Optional[str] = url_base
    current_headers = headers.copy()
    response: Optional[requests.Response] = None
    try:
        page_count = 0
        while current_url:
             # Limitar páginas si es necesario (ej: si top es bajo y no se quiere paginar realmente)
             # if top <= 100 and page_count > 0: break # Ejemplo: no seguir si solo se pidieron pocos

             page_count += 1
             logger.info(f"Listando eventos para '{mailbox}', Página: {page_count}")
             current_params_page = clean_params if page_count == 1 else None
             assert current_url is not None # Para MyPy
             response = requests.get(current_url, headers=current_headers, params=current_params_page, timeout=GRAPH_API_TIMEOUT)
             response.raise_for_status()
             data = response.json()
             page_items = data.get('value', [])
             all_events.extend(page_items)
             current_url = data.get('@odata.nextLink')
             if current_url: logger.debug("Paginación: Siguiente link encontrado.")
             else: logger.debug("Paginación: Fin.")

        logger.info(f"Total eventos listados para '{mailbox}': {len(all_events)}")
        # Devolvemos un formato consistente aunque solo haya una página
        return {'value': all_events, '@odata.count': len(all_events)}

    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en listar_eventos: {req_ex}", exc_info=True); raise Exception(f"Error API listando eventos: {req_ex}")
    except Exception as e: logger.error(f"Error inesperado en listar_eventos: {e}", exc_info=True); raise

def crear_evento(headers: Dict[str, str], titulo: str, inicio: datetime, fin: datetime, asistentes: Optional[List[Dict[str, Any]]] = None, cuerpo: Optional[str] = None, es_reunion_online: bool = False, proveedor_reunion_online: str = "teamsForBusiness", recordatorio_minutos: Optional[int] = None, ubicacion: Optional[str] = None, mostrar_como: str = "busy", mailbox: str = 'me') -> Dict[str, Any]:
    url = f"{BASE_URL}/users/{mailbox}/events"
    if not isinstance(inicio, datetime) or not isinstance(fin, datetime): raise ValueError("'inicio' y 'fin' deben ser datetimes.")
    inicio_tz = _ensure_timezone(inicio); assert inicio_tz is not None
    fin_tz = _ensure_timezone(fin); assert fin_tz is not None

    body: Dict[str, Any] = {"subject": titulo, "start": {"dateTime": inicio_tz.isoformat(), "timeZone": "UTC"}, "end": {"dateTime": fin_tz.isoformat(), "timeZone": "UTC"}}
    if mostrar_como: body["showAs"] = mostrar_como
    if asistentes: body["attendees"] = asistentes
    if cuerpo: body["body"] = {"contentType": "HTML", "content": cuerpo}
    if ubicacion: body["location"] = {"displayName": ubicacion}
    if es_reunion_online: body["isOnlineMeeting"] = True; body["onlineMeetingProvider"] = proveedor_reunion_online
    if recordatorio_minutos is not None: body["isReminderOn"] = True; body["reminderMinutesBeforeStart"] = recordatorio_minutos
    else: body["isReminderOn"] = False
    logger.info(f"Creando evento '{titulo}' para '{mailbox}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def actualizar_evento(headers: Dict[str, str], evento_id: str, nuevos_valores: Dict[str, Any], mailbox: str = 'me') -> Dict[str, Any]:
    url = f"{BASE_URL}/users/{mailbox}/events/{evento_id}"
    payload = nuevos_valores.copy()
    if 'start' in payload and isinstance(payload.get('start'), datetime): start_dt = _ensure_timezone(payload['start']); assert start_dt is not None; payload['start'] = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}
    if 'end' in payload and isinstance(payload.get('end'), datetime): end_dt = _ensure_timezone(payload['end']); assert end_dt is not None; payload['end'] = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}
    etag = payload.pop('@odata.etag', None); current_headers = headers.copy()
    if etag: current_headers['If-Match'] = etag
    logger.info(f"Actualizando evento '{evento_id}' para '{mailbox}'")
    return hacer_llamada_api("PATCH", url, current_headers, json_data=payload)

def eliminar_evento(headers: Dict[str, str], evento_id: str, mailbox: str = 'me') -> Optional[Dict[str, Any]]:
    url = f"{BASE_URL}/users/{mailbox}/events/{evento_id}"
    logger.info(f"Eliminando evento '{evento_id}' para '{mailbox}'")
    # Añadir ETag si se pasa como parámetro opcional
    # etag = kwargs.get('etag') ... if etag: headers['If-Match']=etag
    hacer_llamada_api("DELETE", url, headers) # Devuelve None si 204
    return {"status": "Eliminado", "id": evento_id} # Devolver confirmación

def crear_reunion_teams(headers: Dict[str, str], titulo: str, inicio: datetime, fin: datetime, asistentes: Optional[List[Dict[str, Any]]] = None, cuerpo: Optional[str] = None, mailbox: str = 'me') -> dict:
    logger.info(f"Wrapper: Llamando a crear_evento para crear reunión Teams '{titulo}'")
    return crear_evento(headers=headers, titulo=titulo, inicio=inicio, fin=fin, asistentes=asistentes, cuerpo=cuerpo, es_reunion_online=True, proveedor_reunion_online="teamsForBusiness", mailbox=mailbox)
