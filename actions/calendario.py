# actions/calendario.py (Refactorizado)

import logging
import requests # Solo para tipos de excepción
import json
from typing import Dict, List, Optional, Union, Any
from datetime import datetime, timezone # Importar timezone

# Usar el logger estándar de Azure Functions
logger = logging.getLogger("azure.functions")

# Importar helper y constantes desde la estructura compartida
try:
    # Asume que shared está un nivel arriba de actions
    from ..shared.helpers.http_client import hacer_llamada_api
    from ..shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError as e:
    logging.critical(f"Error CRÍTICO importando helpers/constantes en Calendario: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs):
        raise NotImplementedError("Dependencia 'hacer_llamada_api' no importada correctamente.")

# ---- Helper Interno para Timezone ----
def _ensure_timezone(dt_input: Any) -> Optional[datetime]:
    """Asegura que un datetime tenga timezone UTC si es naive."""
    if isinstance(dt_input, datetime):
        if dt_input.tzinfo is None:
            return dt_input.replace(tzinfo=timezone.utc)
        return dt_input
    # Si no es datetime, intentar parsear si es string ISO
    if isinstance(dt_input, str):
        try:
            dt_parsed = datetime.fromisoformat(dt_input.replace('Z', '+00:00'))
            if dt_parsed.tzinfo is None:
                return dt_parsed.replace(tzinfo=timezone.utc)
            return dt_parsed
        except ValueError:
            logger.warning(f"No se pudo parsear '{dt_input}' como datetime ISO. Devolviendo None.")
            return None
    return None # Devolver None si no es datetime ni string parseable

# ---- FUNCIONES DE ACCIÓN PARA CALENDARIO ----
# Todas usan la firma (parametros: Dict[str, Any], headers: Dict[str, str])

def listar_eventos(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista eventos del calendario, manejando paginación y calendarView.

    Args:
        parametros (Dict[str, Any]): Opcional: 'mailbox' (default 'me'), 'top' (int, default 10),
                                     'start_date' (ISO string o datetime), 'end_date' (ISO string o datetime),
                                     'filter_query', 'order_by', 'select' (List[str]),
                                     'use_calendar_view' (bool, default True).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Un diccionario {'value': [lista_completa_de_eventos]}.
    """
    mailbox: str = parametros.get('mailbox', 'me')
    top: int = int(parametros.get('top', 10))
    start_date_in = parametros.get('start_date')
    end_date_in = parametros.get('end_date')
    filter_query: Optional[str] = parametros.get('filter_query')
    order_by: Optional[str] = parametros.get('order_by')
    select: Optional[List[str]] = parametros.get('select')
    use_calendar_view: bool = parametros.get('use_calendar_view', True)

    start_date_tz = _ensure_timezone(start_date_in)
    end_date_tz = _ensure_timezone(end_date_in)

    base_endpoint = f"{BASE_URL}/users/{mailbox}"
    params_query: Dict[str, Any] = {}
    endpoint_suffix: str = ""

    # Determinar endpoint y parámetros según use_calendar_view
    if use_calendar_view:
        if not start_date_tz or not end_date_tz:
            raise ValueError("Para 'use_calendar_view=True', se requieren 'start_date' y 'end_date'.")
        endpoint_suffix = "/calendarView"
        params_query['startDateTime'] = start_date_tz.isoformat()
        params_query['endDateTime'] = end_date_tz.isoformat()
        params_query['$top'] = min(top, 999) # Limitar top por llamada
        if filter_query: params_query['$filter'] = filter_query
        if order_by: params_query['$orderby'] = order_by
        if select: params_query['$select'] = ','.join(select)
    else:
        # Usar /events si no es calendarView o faltan fechas
        endpoint_suffix = "/events"
        params_query['$top'] = min(top, 999)
        filters = []
        if start_date_tz: filters.append(f"start/dateTime ge '{start_date_tz.isoformat()}'")
        if end_date_tz: filters.append(f"end/dateTime le '{end_date_tz.isoformat()}'")
        if filter_query: filters.append(f"({filter_query})") # Encerrar filtro original
        if filters: params_query['$filter'] = " and ".join(filters)
        if order_by: params_query['$orderby'] = order_by
        if select: params_query['$select'] = ','.join(select)

    url_base = f"{base_endpoint}{endpoint_suffix}"
    # Remover parámetros None antes de la llamada
    clean_params = {k: v for k, v in params_query.items() if v is not None}

    all_events: List[Dict[str, Any]] = []
    current_url: Optional[str] = url_base
    page_count = 0
    max_pages = 100 # Límite de seguridad para paginación

    try:
        while current_url and page_count < max_pages:
            page_count += 1
            logger.info(f"Listando eventos para '{mailbox}', Endpoint: {endpoint_suffix}, Página: {page_count}")

            # Usar el helper centralizado para cada página
            current_params_page = clean_params if page_count == 1 else None
            data = hacer_llamada_api("GET", current_url, headers, params=current_params_page)

            if data:
                page_items = data.get('value', [])
                all_events.extend(page_items)
                current_url = data.get('@odata.nextLink')
                if not current_url:
                    logger.debug("No hay '@odata.nextLink', fin de paginación.")
                    break
            else:
                 logger.warning(f"Llamada a {current_url} para listar eventos devolvió None/vacío.")
                 break

        if page_count >= max_pages:
             logger.warning(f"Se alcanzó límite de {max_pages} páginas listando eventos para '{mailbox}'.")

        logger.info(f"Total eventos listados para '{mailbox}': {len(all_events)}")
        return {'value': all_events} # Devolver siempre la estructura {'value': [...]}

    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en listar_eventos (página {page_count}): {req_ex}", exc_info=True)
        raise Exception(f"Error API listando eventos: {req_ex}") from req_ex
    except Exception as e:
        logger.error(f"Error inesperado en listar_eventos (página {page_count}): {e}", exc_info=True)
        raise


def crear_evento(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea un nuevo evento en el calendario.

    Args:
        parametros (Dict[str, Any]): Debe contener 'titulo', 'inicio', 'fin'.
                                     Opcional: 'mailbox' (default 'me'), 'asistentes' (List[Dict]),
                                     'cuerpo' (str HTML), 'es_reunion_online' (bool),
                                     'proveedor_reunion_online' (str), 'recordatorio_minutos' (int),
                                     'ubicacion' (str), 'mostrar_como' (str, ej. 'busy').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El evento creado devuelto por Graph API.
    """
    mailbox: str = parametros.get('mailbox', 'me')
    titulo: Optional[str] = parametros.get('titulo')
    inicio_in = parametros.get('inicio')
    fin_in = parametros.get('fin')
    asistentes: Optional[List[Dict[str, Any]]] = parametros.get('asistentes')
    cuerpo: Optional[str] = parametros.get('cuerpo')
    es_reunion_online: bool = parametros.get('es_reunion_online', False)
    proveedor_reunion_online: str = parametros.get('proveedor_reunion_online', "teamsForBusiness")
    recordatorio_minutos: Optional[int] = parametros.get('recordatorio_minutos')
    ubicacion: Optional[str] = parametros.get('ubicacion')
    mostrar_como: str = parametros.get('mostrar_como', "busy") # Default a 'busy'

    if not titulo: raise ValueError("Parámetro 'titulo' es requerido.")
    inicio_tz = _ensure_timezone(inicio_in)
    fin_tz = _ensure_timezone(fin_in)
    if not inicio_tz or not fin_tz:
        raise ValueError("Parámetros 'inicio' y 'fin' son requeridos y deben ser datetimes válidos.")
    if fin_tz <= inicio_tz:
        raise ValueError("La fecha/hora 'fin' debe ser posterior a 'inicio'.")

    url = f"{BASE_URL}/users/{mailbox}/events"
    body: Dict[str, Any] = {
        "subject": titulo,
        "start": {"dateTime": inicio_tz.isoformat(), "timeZone": "UTC"}, # Enviar siempre en UTC
        "end": {"dateTime": fin_tz.isoformat(), "timeZone": "UTC"}
    }

    if mostrar_como: body["showAs"] = mostrar_como
    if asistentes and isinstance(asistentes, list): body["attendees"] = asistentes
    if cuerpo: body["body"] = {"contentType": "HTML", "content": cuerpo} # Asumir HTML
    if ubicacion: body["location"] = {"displayName": ubicacion}
    if es_reunion_online:
        body["isOnlineMeeting"] = True
        body["onlineMeetingProvider"] = proveedor_reunion_online
    if recordatorio_minutos is not None:
        try:
            body["isReminderOn"] = True
            body["reminderMinutesBeforeStart"] = int(recordatorio_minutos)
        except ValueError:
             logger.warning(f"Valor inválido para 'recordatorio_minutos': {recordatorio_minutos}. Se ignorará.")
             body["isReminderOn"] = False
    else:
        # Si no se especifica, asegurar que el recordatorio esté apagado
        body["isReminderOn"] = False

    logger.info(f"Creando evento '{titulo}' para '{mailbox}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)


def actualizar_evento(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Actualiza un evento existente. Soporta ETag para concurrencia.

    Args:
        parametros (Dict[str, Any]): Debe contener 'evento_id', 'nuevos_valores' (dict).
                                     Opcional: 'mailbox' (default 'me'), '@odata.etag' dentro de nuevos_valores.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El evento actualizado.
    """
    mailbox: str = parametros.get('mailbox', 'me')
    evento_id: Optional[str] = parametros.get('evento_id')
    nuevos_valores: Optional[Dict[str, Any]] = parametros.get('nuevos_valores')

    if not evento_id: raise ValueError("Parámetro 'evento_id' es requerido.")
    if not nuevos_valores or not isinstance(nuevos_valores, dict):
        raise ValueError("Parámetro 'nuevos_valores' (diccionario) es requerido.")

    url = f"{BASE_URL}/users/{mailbox}/events/{evento_id}"
    payload = nuevos_valores.copy() # Copiar para modificar

    # Convertir fechas/horas a formato Graph si están presentes
    if 'start' in payload:
        start_dt = _ensure_timezone(payload['start'])
        if start_dt: payload['start'] = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}
        else: raise ValueError("Valor inválido para 'start' en nuevos_valores.")
    if 'end' in payload:
        end_dt = _ensure_timezone(payload['end'])
        if end_dt: payload['end'] = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}
        else: raise ValueError("Valor inválido para 'end' en nuevos_valores.")

    # Manejar ETag para concurrencia
    etag = payload.pop('@odata.etag', None)
    current_headers = headers.copy()
    if etag:
        current_headers['If-Match'] = etag
        logger.debug(f"Usando ETag '{etag}' para actualización de evento.")

    logger.info(f"Actualizando evento '{evento_id}' para '{mailbox}'")
    return hacer_llamada_api("PATCH", url, current_headers, json_data=payload)


def eliminar_evento(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina un evento. Soporta ETag para concurrencia.

    Args:
        parametros (Dict[str, Any]): Debe contener 'evento_id'.
                                     Opcional: 'mailbox' (default 'me'), 'etag'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de eliminación.
    """
    mailbox: str = parametros.get('mailbox', 'me')
    evento_id: Optional[str] = parametros.get('evento_id')
    etag: Optional[str] = parametros.get('etag') # ETag puede venir como param separado

    if not evento_id: raise ValueError("Parámetro 'evento_id' es requerido.")

    url = f"{BASE_URL}/users/{mailbox}/events/{evento_id}"
    current_headers = headers.copy()
    if etag:
        current_headers['If-Match'] = etag
        logger.debug(f"Usando ETag '{etag}' para eliminación de evento.")
    else:
        logger.warning(f"Eliminando evento {evento_id} sin ETag. Podría fallar si fue modificado.")

    logger.info(f"Eliminando evento '{evento_id}' para '{mailbox}'")
    # hacer_llamada_api devuelve None en éxito 204 (DELETE)
    hacer_llamada_api("DELETE", url, current_headers)
    return {"status": "Eliminado", "id": evento_id} # Devolver confirmación explícita


def crear_reunion_teams(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Wrapper para crear un evento que es una reunión de Teams.

    Args:
        parametros (Dict[str, Any]): Igual que `crear_evento`, pero 'es_reunion_online' se fuerza a True.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El evento/reunión creado.
    """
    logger.info(f"Wrapper: Llamando a crear_evento para crear reunión Teams.")
    # Forzar los parámetros necesarios para reunión de Teams
    parametros_reunion = parametros.copy()
    parametros_reunion['es_reunion_online'] = True
    parametros_reunion['proveedor_reunion_online'] = "teamsForBusiness"

    # Llamar a la función principal de crear evento con los parámetros ajustados
    return crear_evento(parametros_reunion, headers)

# --- FIN DEL MÓDULO actions/calendario.py ---
