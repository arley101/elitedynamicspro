import logging
import os
import requests
from auth import obtener_token  # Importante: Importar la función obtener_token
from typing import Dict, List, Optional, Union
from datetime import datetime, timezone

# Configuración básica de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Variables de entorno (¡CRUCIALES!)
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
GRAPH_SCOPE = os.getenv('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')  # Valor por defecto

# Verificar variables de entorno (¡CRUCIAL!)
if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE]):
    logging.error("❌ Faltan variables de entorno (CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE). La función no puede funcionar.")
    raise Exception("Faltan variables de entorno.")

BASE_URL = "https://graph.microsoft.com/v1.0/me/events"
HEADERS = {
    'Authorization': None,  # Inicialmente None, se actualiza con cada request
    'Content-Type': 'application/json'
}
MAILBOX = os.getenv('MAILBOX', 'me') #para permitir acceso a otros calendarios


# Función para obtener el token y actualizar los HEADERS
def _actualizar_headers() -> None:
    """Obtiene un nuevo token de acceso y actualiza el diccionario HEADERS."""
    try:
        HEADERS['Authorization'] = f'Bearer {obtener_token()}'
    except Exception as e:  # Captura la excepción de obtener_token
        logging.error(f"❌ Error al obtener el token: {e}")
        raise Exception(f"Error al obtener el token: {e}")



# ---- FUNCIONES DE GESTIÓN DE CALENDARIO DE OUTLOOK ----
def listar_eventos(
    top: int = 10,
    start_date: Optional[datetime] = None,
    end_date: Optional[datetime] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None,
    select: Optional[List[str]] = None
) -> dict:
    """
    Lista eventos del calendario de Outlook, con soporte para paginación, rango de fechas, filtrado, ordenamiento y selección de campos.

    Args:
        top: El número máximo de eventos a devolver.
        start_date: Fecha y hora de inicio para filtrar eventos.
        end_date: Fecha y hora de fin para filtrar eventos.
        filter_query: Una cadena de consulta para filtrar eventos (ej, "start/dateTime ge '2024-01-01T00:00:00Z'").
        order_by: Una cadena para especificar el orden de los resultados (ej, "start/dateTime desc").
        select: Lista de campos a seleccionar.

    Returns:
        Un diccionario con la respuesta de la API de Microsoft Graph.
    """
    _actualizar_headers()
    url = f"{BASE_URL}?$top={top}"

    if start_date:
        url += f"&startDateTime={start_date.isoformat()}"
    if end_date:
        url += f"&endDateTime={end_date.isoformat()}"
    if filter_query:
        url += f"&$filter={filter_query}"
    if order_by:
        url += f"&$orderby={order_by}"
    if select:
        url += f"&$select={','.join(select)}"

    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Listados eventos del calendario. Top: {top}, Start: {start_date}, End: {end_date}, Filter: {filter_query}, Order: {order_by}, Select: {select}")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar eventos: {e}")
        raise Exception(f"Error al listar eventos: {e}")



def crear_evento(
    titulo: str,
    inicio: datetime,
    fin: datetime,
    asistentes: Optional[List[Dict[str, Union[str, dict]]]] = None,
    cuerpo: Optional[str] = None,
    es_reunion_online: bool = False,
    proveedor_reunion_online: str = "teamsForBusiness",
    recordatorio_minutos: Optional[int] = None
) -> dict:
    """
    Crea un nuevo evento u reunión en el calendario de Outlook.

    Args:
        titulo: El título del evento.
        inicio: La fecha y hora de inicio del evento (datetime object).
        fin: La fecha y hora de fin del evento (datetime object).
        asistentes: Una lista de diccionarios con información sobre los asistentes (opcional).
        cuerpo: El cuerpo del evento (opcional).
        es_reunion_online: Indica si el evento es una reunión en línea (opcional).
        proveedor_reunion_online: El proveedor de la reunión en línea (por ejemplo, "teamsForBusiness") (opcional).
        recordatorio_minutos: Minutos antes del inicio del evento para enviar un recordatorio.
    Returns:
        La respuesta de la API de Microsoft Graph.
    """
    _actualizar_headers()
    url = BASE_URL
    body = {
        "subject": titulo,
        "start": {"dateTime": inicio.isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": fin.isoformat(), "timeZone": "UTC"},
    }
    if asistentes:
        body["attendees"] = [{"emailAddress": {"address": a['emailAddress']}, "type": a.get('type', 'required')} for a in asistentes]
    if cuerpo:
        body["body"] = {"contentType": "HTML", "content": cuerpo}  # o "Text"
    if es_reunion_online:
        body["isOnlineMeeting"] = es_reunion_online
        body["onlineMeetingProvider"] = proveedor_reunion_online
    if recordatorio_minutos:
        body["reminderMinutesBeforeStart"] = recordatorio_minutos

    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Evento '{titulo}' creado.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear evento: {e}")
        raise Exception(f"Error al crear evento: {e}")



def actualizar_evento(evento_id: str, nuevos_valores: dict) -> dict:
    """Actualiza un evento existente en el calendario de Outlook."""
    _actualizar_headers()
    url = f"{BASE_URL}/{evento_id}"
    try:
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status()
        logging.info(f"Evento '{evento_id}' actualizado. Nuevos valores: {nuevos_valores}")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al actualizar evento '{evento_id}': {e}")
        raise Exception(f"Error al actualizar evento '{evento_id}': {e}")



def eliminar_evento(evento_id: str) -> dict:
    """Elimina un evento del calendario de Outlook."""
    _actualizar_headers()
    url = f"{BASE_URL}/{evento_id}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Evento '{evento_id}' eliminado.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar el evento '{evento_id}': {e}")
        raise Exception(f"Error al eliminar el evento '{evento_id}': {e}")



def crear_reunion_teams(
    titulo: str,
    inicio: datetime,
    fin: datetime,
    asistentes: Optional[List[Dict[str, Union[str, dict]]]] = None,
    cuerpo: Optional[str] = None
) -> dict:
    """
    Crea una reunión de Teams y un evento de calendario asociado en Outlook.

    Args:
        titulo: El título de la reunión.
        inicio: La fecha y hora de inicio de la reunión (datetime object).
        fin: La fecha y hora de fin de la reunión (datetime object).
        asistentes: Una lista de diccionarios con información sobre los asistentes (opcional).
        cuerpo: El cuerpo de la reunión.

    Returns:
        La respuesta de la API de Microsoft Graph.
    """
    return crear_evento(titulo, inicio, fin, asistentes, cuerpo, es_reunion_online=True, proveedor_reunion_online="teamsForBusiness")
