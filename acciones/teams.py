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

BASE_URL = "https://graph.microsoft.com/v1.0"
HEADERS = {
    'Authorization': None,  # Inicialmente None, se actualiza con cada request
    'Content-Type': 'application/json'
}
MAILBOX = os.getenv('MAILBOX', 'me') #para permitir acceso a otros buzones

# Función para obtener el token y actualizar los HEADERS
def _actualizar_headers() -> None:
    """Obtiene un nuevo token de acceso y actualiza el diccionario HEADERS."""
    try:
        HEADERS['Authorization'] = f'Bearer {obtener_token()}'
    except Exception as e:  # Captura la excepción de obtener_token
        logging.error(f"❌ Error al obtener el token: {e}")
        raise Exception(f"Error al obtener el token: {e}")



# ---- CHAT ----

def listar_chats(top: int = 20, skip: int = 0, filter_query: Optional[str] = None, order_by: Optional[str] = None) -> dict:
    """Lista los chats del usuario actual."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/chats?$top={top}&$skip={skip}"
    if filter_query:
        url += f"&$filter={filter_query}"
    if order_by:
        url += f"&$orderby={order_by}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Listados chats del usuario. Top: {top}, Skip:{skip}, Filter: {filter_query}, Order: {order_by}")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar chats: {e}")
        raise Exception(f"Error al listar chats: {e}")



def obtener_chat(chat_id: str) -> dict:
    """Obtiene un chat específico."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/chats/{chat_id}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Obtenido chat: {chat_id}")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener chat {chat_id}: {e}")
        raise Exception(f"Error al obtener chat: {e}")


def crear_chat(participantes: List[Dict[str, dict]], tipo: str = "chat", tema: Optional[str] = None) -> dict:
    """Crea un nuevo chat."""
    _actualizar_headers()
    url = f"{BASE_URL}/chats"
    body = {
        "chatType": tipo,  # "chat" o "meeting"
        "members": [
            {
                "@odata.type": "microsoft.graph.chatMember#Microsoft.Graph.aadUserConversationMember",
                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{participant['user']['id']}"
            }
            for participant in participantes
        ]
    }
    if tema:
        body["topic"] = tema
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Chat creado. Tipo: {tipo}, Tema: {tema}")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear chat: {e}")
        raise Exception(f"Error al crear chat: {e}")



def enviar_mensaje_chat(chat_id: str, mensaje: str, tipo_contenido: str = "Text") -> dict:
    """Envía un mensaje a un chat."""
    _actualizar_headers()
    url = f"{BASE_URL}/chats/{chat_id}/messages"
    body = {
        "body": {
            "contentType": tipo_contenido,  # "Text" o "Html"
            "content": mensaje
        }
    }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Mensaje enviado al chat '{chat_id}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al enviar mensaje al chat '{chat_id}': {e}")
        raise Exception(f"Error al enviar mensaje al chat '{chat_id}': {e}")



def obtener_mensajes_chat(chat_id: str, top: int = 20, skip: int = 0) -> dict:
    """Obtiene los mensajes de un chat, con soporte para paginación."""
    _actualizar_headers()
    url = f"{BASE_URL}/chats/{chat_id}/messages?$top={top}&$skip={skip}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Obtenidos mensajes del chat '{chat_id}'. Top: {top}, Skip: {skip}")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener mensajes del chat '{chat_id}': {e}")
        raise Exception(f"Error al obtener mensajes del chat '{chat_id}': {e}")



def actualizar_mensaje_chat(chat_id: str, message_id: str, contenido: str, tipo_contenido: str = "Text") -> dict:
    """Actualiza un mensaje existente en un chat."""
    _actualizar_headers()
    url = f"{BASE_URL}/chats/{chat_id}/messages/{message_id}"
    body = {
        "body": {
            "contentType": tipo_contenido,
            "content": contenido
        }
    }
    try:
        response = requests.patch(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Mensaje '{message_id}' actualizado en el chat '{chat_id}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al actualizar el mensaje '{message_id}' en el chat '{chat_id}': {e}")
        raise Exception(f"Error al actualizar el mensaje '{message_id}' en el chat '{chat_id}': {e}")



def eliminar_mensaje_chat(chat_id: str, message_id: str) -> dict:
    """Elimina un mensaje de un chat."""
    _actualizar_headers()
    url = f"{BASE_URL}/chats/{chat_id}/messages/{message_id}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Mensaje '{message_id}' eliminado del chat '{chat_id}'.")
        return {"status": "Mensaje Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar el mensaje '{message_id}' del chat '{chat_id}': {e}")
        raise Exception(f"Error al eliminar el mensaje '{message_id}' del chat '{chat_id}': {e}")


# ---- REUNIONES (CALENDARIO) ----
def listar_reuniones(
    top: int = 10,
    start_date: Optional[datetime] = None,
    end_date: Optional[datetime] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None,
    select: Optional[List[str]] = None
) -> dict:
    """
    Lista reuniones de Teams (eventos online) del calendario de Outlook, con soporte para filtrado.

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
    url = f"{BASE_URL}?$filter=isOnlineMeeting eq true"

    if start_date:
        url += f" and start/dateTime ge '{start_date.isoformat()}'"
    if end_date:
        url += f" and end/dateTime le '{end_date.isoformat()}'"
    if filter_query:
        url = f"{url} and {filter_query}" if start_date or end_date else f"{url} and {filter_query}"
    url = f"{url}&$top={top}"
    if order_by:
        url += f"&$orderby={order_by}"
    if select:
        url += f"&$select={','.join(select)}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Listadas reuniones de Teams. Top: {top}, Start: {start_date}, End: {end_date}, Filter: {filter_query}, Order: {order_by}, Select: {select}")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar reuniones de Teams: {e}")
        raise Exception(f"Error al listar reuniones de Teams: {e}")



def crear_reunion_teams(
    titulo: str,
    inicio: datetime,
    fin: datetime,
    asistentes: Optional[List[Dict[str, Union[str, dict]]]] = None,
    cuerpo: Optional[str] = None,
    recordatorio_minutos: Optional[int] = None
) -> dict:
    """
    Crea una reunión de Teams y un evento de calendario asociado en Outlook.

    Args:
        titulo: El título de la reunión.
        inicio: La fecha y hora de inicio de la reunión (datetime object).
        fin: La fecha y hora de fin de la reunión (datetime object).
        asistentes: Una lista de diccionarios con información sobre los asistentes (opcional).
        cuerpo: El cuerpo de la reunión
        recordatorio_minutos: Minutos antes del inicio del evento para enviar un recordatorio.
    Returns:
        La respuesta de la API de Microsoft Graph.
    """
    _actualizar_headers()
    url = f"{BASE_URL}"
    body = {
        "subject": titulo,
        "start": {"dateTime": inicio.isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": fin.isoformat(), "timeZone": "UTC"},
        "isOnlineMeeting": True,
        "onlineMeetingProvider": "teamsForBusiness",
    }
    if asistentes:
        body["attendees"] = [{"emailAddress": {"address": a['emailAddress']}, "type": a.get('type', 'required')} for a in asistentes]
    if cuerpo:
        body["body"] = {"contentType": "HTML", "content": cuerpo}
    if recordatorio_minutos:
        body["reminderMinutesBeforeStart"] = recordatorio_minutos
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Reunion de teams creada: {data}")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear la reunión de Teams: {e}")
        raise Exception(f"Error al crear la reunión de Teams: {e}")


# ---- EQUIPOS Y CANALES ----

def listar_equipos(top: int = 20, skip: int = 0, filter_query: Optional[str] = None) -> dict:
    """Lista los equipos de Microsoft Teams a los que pertenece el usuario."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/joinedTeams?$top={top}&$skip={skip}"
    if filter_query:
        url += f"&$filter={filter_query}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Listados equipos del usuario. Top: {top}, Skip: {skip}, Filter: {filter_query}")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar equipos: {e}")
        raise Exception(f"Error al listar equipos: {e}")



def obtener_equipo(team_id: str) -> dict:
    """Obtiene información sobre un equipo de Microsoft Teams específico."""
    _actualizar_headers()
    url = f"{BASE_URL}/teams/{team_id}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Obtenido equipo con ID: {team_id}")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener equipo {team_id}: {e}")
        raise Exception(f"Error al obtener equipo {team_id}: {e}")



def crear_equipo(nombre: str, descripcion: str = "Equipo creado por Elite Dynamics Pro", tipo_plantilla: str = "standard") -> dict:
    """Crea un nuevo equipo de Microsoft Teams."""
    _actualizar_headers()
    url = f"{BASE_URL}/teams"
    body = {
        "template@odata.bind": f"https://graph.microsoft.com/v1.0/teamsTemplates('{tipo_plantilla}')",
        "displayName": nombre,
        "description": descripcion
    }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        team_id = data.get('id')
        logging.info(f"Equipo '{nombre}' creado exitosamente con ID: {team_id}.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear equipo '{nombre}': {e}")
        raise Exception(f"Error al crear equipo '{nombre}': {e}")



def archivar_equipo(team_id: str, set_frozen: bool = False) -> dict:
    """Archiva un equipo de Microsoft Teams."""
    _actualizar_headers()
    url = f"{BASE_URL}/teams/{team_id}/archive"
    body = {"shouldSetSpoSiteReadOnlyForUsers": set_frozen}
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Equipo '{team_id}' archivado.  Frozen: {set_frozen}")
        return {"status": "Archivado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al archivar equipo '{team_id}': {e}")
        raise Exception(f"Error al archivar equipo '{team_id}': {e}")



def unarchivar_equipo(team_id: str) -> dict:
    """Desarchiva un equipo de Microsoft Teams."""
    _actualizar_headers()
    url = f"{BASE_URL}/teams/{team_id}/unarchive"
    try:
        response = requests.post(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Equipo '{team_id}' desarchivado.")
        return {"status": "Desarchivado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al desarchivar equipo '{team_id}': {e}")
        raise Exception(f"Error al desarchivar equipo '{team_id}': {e}")



def eliminar_equipo(team_id: str) -> dict:
    """Elimina un equipo de Microsoft Teams."""
    _actualizar_headers()
    url = f"{BASE_URL}/teams/{team_id}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Equipo '{team_id}' eliminado.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar el equipo '{team_id}': {e}")
        raise Exception(f"Error al eliminar el equipo '{team_id}': {e}")



def listar_canales(team_id: str) -> dict:
    """Lista los canales de un equipo de Microsoft Teams."""
    _actualizar_headers()
    url = f"{BASE_URL}/teams/{team_id}/channels"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Listados canales del equipo '{team_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar canales del equipo '{team_id}': {e}")
        raise Exception(f"Error al listar canales del equipo '{team_id}': {e}")



def obtener_canal(team_id: str, channel_id: str) -> dict:
    """Obtiene información sobre un canal específico de un equipo de Microsoft Teams."""
    _actualizar_headers()
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Obtenido canal '{channel_id}' del equipo '{team_id}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener canal '{channel_id}' del equipo '{team_id}': {e}")
        raise Exception(f"Error al obtener canal '{channel_id}' del equipo '{team_id}': {e}")



def crear_canal(team_id: str, nombre_canal: str, descripcion: str = "Canal creado por Elite Dynamics Pro", tipo_canal: str = "standard") -> dict:
    """Crea un nuevo canal en un equipo de Microsoft Teams."""
    _actualizar_headers()
    url = f"{BASE_URL}/teams/{team_id}/channels"
    body = {
        "displayName": nombre_canal,
        "description": descripcion,
        "membershipType": tipo_canal # Standard, Private, Shared
    }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        channel_id = data.get('id')
        logging.info(f"Canal '{nombre_canal}' creado en el equipo '{team_id}' con ID: {channel_id}.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear canal '{nombre_canal}' en el equipo '{team_id}': {e}")
        raise Exception(f"Error al crear canal '{nombre_canal}' en el equipo '{team_id}': {e}")



def actualizar_canal(team_id: str, channel_id: str, nuevos_valores: dict) -> dict:
    """Actualiza la información de un canal de Microsoft Teams."""
    _actualizar_headers()
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"
    try:
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status()
        logging.info(f"Canal '{channel_id}' actualizado en el equipo '{team_id}'. Nuevos valores: {nuevos_valores}")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al actualizar el canal '{channel_id}' en el equipo '{team_id}': {e}")
        raise Exception(f"Error al actualizar el canal '{channel_id}' en el equipo '{team_id}': {e}")



def eliminar_canal(team_id: str, channel_id: str) -> dict:
    """Elimina un canal de un equipo de Microsoft Teams."""
    _actualizar_headers()
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Canal '{channel_id}' eliminado del equipo '{team_id}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar el canal '{channel_id}' del equipo '{team_id}': {e}")
        raise Exception(f"Error al eliminar el canal '{channel_id}' del equipo '{team_id}': {e}")



# ---- MENSAJES DE CANAL ----
def enviar_mensaje_canal(team_id: str, channel_id: str, mensaje: str, tipo_contenido: str = "Text") -> dict:
    """Envía un mensaje a un canal de Microsoft Teams."""
    _actualizar_headers()
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}/messages"
    body = {
        "body": {
            "contentType": tipo_contenido,  # "Text" o "Html"
            "content": mensaje
        }
    }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Mensaje enviado al canal '{channel_id}' del equipo '{team_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al enviar mensaje al canal '{channel_id}' del equipo '{team_id}': {e}")
        raise Exception(f"Error al enviar mensaje al canal '{channel_id}' del equipo '{team_id}': {e}")
