import logging
import requests
import json
# Corregido: Añadir Any
from typing import Dict, List, Optional, Union, Any
from datetime import datetime, timezone

# Usar el logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar constantes globales desde shared/constants.py
try:
    from shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    # Fallback
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_TIMEOUT = 45
    logger.warning("No se pudo importar constantes desde shared (Calendario), usando defaults.")


# ---- FUNCIONES DE CALENDARIO (Refactorizadas) ----
# Aceptan 'headers' como parámetro

def _ensure_timezone(dt: Optional[datetime]) -> Optional[datetime]:
    """Asegura que el datetime tenga timezone (UTC)."""
    if dt and isinstance(dt, datetime) and dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt

def listar_eventos(
    headers: Dict[str, str],
    top: int = 10,
    start_date: Optional[datetime] = None,
    end_date: Optional[datetime] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None,
    select: Optional[List[str]] = None,
    use_calendar_view: bool = True,
    mailbox: str = 'me'
) -> Dict[str, Any]:
    """Lista eventos del calendario. Requiere headers autenticados."""
    base_endpoint = f"{BASE_URL}/users/{mailbox}"
    params: Dict[str, Any] = {}
    endpoint_suffix: str = ""

    # Aplicar corrección MyPy [assignment]
    start_date_tz = _ensure_timezone(start_date) # type: ignore[assignment]
    end_date_tz = _ensure_timezone(end_date) # type: ignore[assignment]

    if use_calendar_view and start_date_tz and end_date_tz:
        endpoint_suffix = "/calendarView"
        # Aplicar corrección MyPy [union-attr]
        assert start_date_tz is not None
        assert end_date_tz is not None
        params['startDateTime'] = start_date_tz.isoformat()
        params['endDateTime'] = end_date_tz.isoformat()
        params['$top'] = int(top)
        if filter_query: params['$filter'] = filter_query
        if order_by: params['$orderby'] = order_by
        if select: params['$select'] = ','.join(select)
    else:
        endpoint_suffix = "/events"
        params['$top'] = int(top)
        filters = []
        if start_date_tz: filters.append(f"start/dateTime ge '{start_date_tz.isoformat()}'")
        if end_date_tz: filters.append(f"end/dateTime le '{end_date_tz.isoformat()}'")
        if filter_query: filters.append(f"({filter_query})")
        if filters: params['$filter'] = " and ".join(filters)
        if order_by: params['$orderby'] = order_by
        if select: params['$select'] = ','.join(select)

    url = f"{base_endpoint}{endpoint_suffix}"
    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None

    try:
        logger.info(f"API Call: GET {url} Params: {clean_params} (Listando eventos para '{mailbox}')")
        response = requests.get(url, headers=headers, params=clean_params, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} eventos para '{mailbox}'.")
        return data
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en listar_eventos: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en listar_eventos: {e}", exc_info=True); raise

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
    mailbox: str = 'me'
) -> Dict[str, Any]:
    url = f"{BASE_URL}/users/{mailbox}/events"
    if not isinstance(inicio, datetime) or not isinstance(fin, datetime): raise ValueError("'inicio' y 'fin' deben ser datetimes.")
    inicio_tz = _ensure_timezone(inicio); assert inicio_tz is not None
    fin_tz = _ensure_timezone(fin); assert fin_tz is not None

    body: Dict[str, Any] = {"subject": titulo, "start": {"dateTime": inicio_tz.isoformat(), "timeZone": "UTC"}, "end": {"dateTime": fin_tz.isoformat(), "timeZone": "UTC"}}
    if mostrar_como: body["showAs"] = mostrar_como
    if asistentes: body["attendees"] = asistentes
    if cuerpo: body["body"] = {"contentType": "HTML", "content": cuerpo}
    if ubicacion: body["location"] = {"displayName": ubicacion}
    if es_reunion_online: body["isOnlineMeeting"] = True; body["onlineMeetingProvider"] = proveedor_reunion_online # Proveedor solo si es online
    if recordatorio_minutos is not None: body["isReminderOn"] = True; body["reminderMinutesBeforeStart"] = recordatorio_minutos
    else: body["isReminderOn"] = False

    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando evento '{titulo}' para '{mailbox}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Evento '{titulo}' creado para '{mailbox}'. ID: {data.get('id')}"); return data
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en crear_evento: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en crear_evento: {e}", exc_info=True); raise

def actualizar_evento(headers: Dict[str, str], evento_id: str, nuevos_valores: Dict[str, Any], mailbox: str = 'me') -> Dict[str, Any]:
    url = f"{BASE_URL}/users/{mailbox}/events/{evento_id}"
    payload = nuevos_valores.copy()
    if 'start' in payload and isinstance(payload.get('start'), datetime): start_dt = _ensure_timezone(payload['start']); assert start_dt is not None; payload['start'] = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}
    if 'end' in payload and isinstance(payload.get('end'), datetime): end_dt = _ensure_timezone(payload['end']); assert end_dt is not None; payload['end'] = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Actualizando evento '{evento_id}' para '{mailbox}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        etag = payload.pop('@odata.etag', None)
        if etag: current_headers['If-Match'] = etag
        response = requests.patch(url, headers=current_headers, json=payload, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Evento '{evento_id}' actualizado para '{mailbox}'."); return data
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en actualizar_evento {evento_id}: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en actualizar_evento {evento_id}: {e}", exc_info=True); raise

def eliminar_evento(headers: Dict[str, str], evento_id: str, mailbox: str = 'me') -> Dict[str, Any]:
    url = f"{BASE_URL}/users/{mailbox}/events/{evento_id}"; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando evento '{evento_id}' para '{mailbox}')")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); logger.info(f"Evento '{evento_id}' eliminado para '{mailbox}'."); return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en eliminar_evento {evento_id}: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en eliminar_evento {evento_id}: {e}", exc_info=True); raise

def crear_reunion_teams(headers: Dict[str, str], titulo: str, inicio: datetime, fin: datetime, asistentes: Optional[List[Dict[str, Any]]] = None, cuerpo: Optional[str] = None, mailbox: str = 'me') -> dict:
    logger.info(f"Llamando a crear_evento para crear reunión Teams '{titulo}' para '{mailbox}'")
    return crear_evento(headers=headers, titulo=titulo, inicio=inicio, fin=fin, asistentes=asistentes, cuerpo=cuerpo, es_reunion_online=True, proveedor_reunion_online="teamsForBusiness", mailbox=mailbox)
