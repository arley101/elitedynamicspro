import logging
import requests
import json # Para manejo de errores
from typing import Dict, List, Optional, Union, Any
from datetime import datetime, timezone

# Usar el logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar constantes globales desde __init__.py
try:
    from .. import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    # Fallback si se ejecuta standalone
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_TIMEOUT = 45
    logger.warning("No se pudo importar BASE_URL/GRAPH_API_TIMEOUT desde el padre, usando defaults.")

# ---- CHAT ----
# Estas funciones usan /me/chats o /chats/{id} y requieren headers delegados

def listar_chats(headers: Dict[str, str], top: int = 20, skip: int = 0, filter_query: Optional[str] = None, order_by: Optional[str] = None, expand: Optional[str] = None) -> dict:
    """Lista los chats del usuario actual (/me). Requiere headers delegados."""
    url = f"{BASE_URL}/me/chats"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query: params['$filter'] = filter_query
    if order_by: params['$orderby'] = order_by
    if expand: params['$expand'] = expand # Ej: 'members'

    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} Params: {clean_params} (Listando chats /me)")
        response = requests.get(url, headers=headers, params=clean_params, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} chats para /me.")
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en listar_chats (Teams /me): {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_chats (Teams /me): {e}", exc_info=True)
        raise

def obtener_chat(headers: Dict[str, str], chat_id: str) -> dict:
    """Obtiene un chat específico por ID. Requiere headers (probablemente delegados)."""
    # Nota: Obtener un chat específico puede requerir permisos sobre ESE chat.
    url = f"{BASE_URL}/chats/{chat_id}" # Endpoint directo al chat ID
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo chat '{chat_id}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Obtenido chat: {chat_id}")
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en obtener_chat {chat_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_chat {chat_id}: {e}", exc_info=True)
        raise

def crear_chat(headers: Dict[str, str], miembros: List[Dict[str, Any]], tipo_chat: str = "oneOnOne", tema: Optional[str] = None) -> dict:
    """Crea un nuevo chat (oneOnOne, group). Requiere headers (probablemente delegados)."""
    # 'miembros' debe ser una lista de: {"@odata.type": "#microsoft.graph.aadUserConversationMember", "user@odata.bind": "https://graph.microsoft.com/v1.0/users('USER_ID_OR_UPN')", "roles": ["owner" | "guest"]}
    url = f"{BASE_URL}/chats"
    body: Dict[str, Any] = {
        "chatType": tipo_chat,
        "members": miembros
    }
    if tema:
        body["topic"] = tema

    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando chat tipo '{tipo_chat}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 201 Created
        data = response.json()
        logger.info(f"Chat creado. Tipo: {tipo_chat}, Tema: {tema}, ID: {data.get('id')}")
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en crear_chat: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_chat: {e}", exc_info=True)
        raise

def enviar_mensaje_chat(headers: Dict[str, str], chat_id: str, mensaje: str, tipo_contenido: str = "text") -> dict:
    """Envía un mensaje a un chat. Requiere headers (probablemente delegados)."""
    url = f"{BASE_URL}/chats/{chat_id}/messages"
    body = {
        "body": {
            "contentType": tipo_contenido.lower(), # 'text' o 'html'
            "content": mensaje
        }
    }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Enviando mensaje a chat '{chat_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 201 Created
        data = response.json()
        logger.info(f"Mensaje enviado al chat '{chat_id}'.")
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en enviar_mensaje_chat {chat_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en enviar_mensaje_chat {chat_id}: {e}", exc_info=True)
        raise

def obtener_mensajes_chat(headers: Dict[str, str], chat_id: str, top: int = 20, skip: int = 0) -> dict:
    """Obtiene los mensajes de un chat, con paginación. Requiere headers."""
    url = f"{BASE_URL}/chats/{chat_id}/messages"
    params = {'$top': int(top), '$skip': int(skip), '$orderby': 'createdDateTime desc'} # Ordenar por más reciente
    all_messages = []
    current_url: Optional[str] = url
    current_headers = headers.copy()
    response: Optional[requests.Response] = None

    try:
        page_count = 0
        # Implementar paginación similar a otras funciones de listado si se necesitan todos los mensajes
        # Por ahora, solo obtiene la primera página según top/skip
        logger.info(f"API Call: GET {current_url} Params: {params} (Obteniendo mensajes chat '{chat_id}')")
        response = requests.get(current_url, headers=current_headers, params=params, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Obtenidos {len(data.get('value',[]))} mensajes del chat '{chat_id}'.")
        return data # Devolver la respuesta directa con paginación si aplica

    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en obtener_mensajes_chat {chat_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_mensajes_chat {chat_id}: {e}", exc_info=True)
        raise

def actualizar_mensaje_chat(headers: Dict[str, str], chat_id: str, message_id: str, contenido: str, tipo_contenido: str = "text") -> dict:
    """Actualiza un mensaje existente en un chat. Requiere headers."""
    url = f"{BASE_URL}/chats/{chat_id}/messages/{message_id}"
    body = {
        "body": {
            "contentType": tipo_contenido.lower(),
            "content": contenido
        }
    }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Actualizando mensaje '{message_id}' en chat '{chat_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.patch(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 204 No Content
        logger.info(f"Mensaje '{message_id}' actualizado en chat '{chat_id}'.")
        # Devolver un estado ya que no hay cuerpo
        return {"status": "Mensaje Actualizado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en actualizar_mensaje_chat {message_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en actualizar_mensaje_chat {message_id}: {e}", exc_info=True)
        raise

def eliminar_mensaje_chat(headers: Dict[str, str], chat_id: str, message_id: str) -> dict:
    """Elimina un mensaje de un chat (soft delete). Requiere headers."""
    # Nota: Solo el usuario que envió el mensaje puede eliminarlo (permiso delegado ChatMessage.ReadWrite)
    url = f"{BASE_URL}/me/chats/{chat_id}/messages/{message_id}/softDelete" # Usar softDelete
    # O si se quiere eliminar realmente (requiere permisos de app Chat.ReadWrite.All):
    # url = f"{BASE_URL}/chats/{chat_id}/messages/{message_id}" # DELETE
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Soft deleting mensaje '{message_id}' en chat '{chat_id}')")
        # SoftDelete es POST sin cuerpo
        response = requests.post(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 204 No Content
        logger.info(f"Mensaje '{message_id}' eliminado (soft delete) del chat '{chat_id}'.")
        return {"status": "Mensaje Eliminado (Soft)", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en eliminar_mensaje_chat {message_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_mensaje_chat {message_id}: {e}", exc_info=True)
        raise


# ---- EQUIPOS Y CANALES ----
# Usan /me/joinedTeams o /teams/{id}

def listar_equipos(headers: Dict[str, str], top: int = 20, skip: int = 0, filter_query: Optional[str] = None) -> dict:
    """Lista los equipos a los que pertenece el usuario actual (/me). Requiere headers delegados."""
    url = f"{BASE_URL}/me/joinedTeams"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query: params['$filter'] = filter_query

    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} Params: {clean_params} (Listando equipos /me)")
        response = requests.get(url, headers=headers, params=clean_params, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} equipos para /me.")
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en listar_equipos (Teams /me): {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_equipos (Teams /me): {e}", exc_info=True)
        raise

def obtener_equipo(headers: Dict[str, str], team_id: str) -> dict:
    """Obtiene información sobre un equipo específico por ID. Requiere headers."""
    url = f"{BASE_URL}/teams/{team_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo equipo '{team_id}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Obtenido equipo con ID: {team_id}")
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en obtener_equipo {team_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_equipo {team_id}: {e}", exc_info=True)
        raise

def crear_equipo(headers: Dict[str, str], nombre: str, descripcion: str = "Equipo creado por API", tipo_plantilla: str = "standard") -> dict:
    """Crea un nuevo equipo de Microsoft Teams. Requiere headers."""
    # Requiere permisos elevados como Team.Create o Group.ReadWrite.All
    url = f"{BASE_URL}/teams"
    body = {
        "template@odata.bind": f"{BASE_URL}/teamsTemplates('{tipo_plantilla}')",
        "displayName": nombre,
        "description": descripcion
        # Se pueden añadir members aquí si se tienen los user IDs
    }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando equipo '{nombre}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT * 2) # Creación puede tardar
        response.raise_for_status() # Puede devolver 202 Accepted si es asíncrono
        if response.status_code == 202:
             monitor_url = response.headers.get('Location')
             logger.info(f"Creación de equipo '{nombre}' iniciada (asíncrona). Monitor: {monitor_url}")
             return {"status": "Creación Iniciada", "code": response.status_code, "monitorUrl": monitor_url}
        else:
            data = response.json()
            team_id = data.get('id')
            logger.info(f"Equipo '{nombre}' creado exitosamente con ID: {team_id}.")
            return data # Asume 201 Created
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en crear_equipo: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_equipo: {e}", exc_info=True)
        raise

def archivar_equipo(headers: Dict[str, str], team_id: str, set_frozen: bool = False) -> dict:
    """Archiva un equipo. Requiere headers."""
    # Requiere Group.ReadWrite.All o TeamSettings.ReadWrite.All
    url = f"{BASE_URL}/teams/{team_id}/archive"
    body = {"shouldSetSpoSiteReadOnlyForUsers": set_frozen}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Archivando equipo '{team_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 202 Accepted
        logger.info(f"Archivado de equipo '{team_id}' iniciado.")
        return {"status": "Archivado Iniciado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en archivar_equipo {team_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en archivar_equipo {team_id}: {e}", exc_info=True)
        raise

def unarchivar_equipo(headers: Dict[str, str], team_id: str) -> dict:
    """Desarchiva un equipo. Requiere headers."""
     # Requiere Group.ReadWrite.All o TeamSettings.ReadWrite.All
    url = f"{BASE_URL}/teams/{team_id}/unarchive"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Desarchivando equipo '{team_id}')")
        response = requests.post(url, headers=headers, timeout=GRAPH_API_TIMEOUT) # POST sin cuerpo
        response.raise_for_status() # 202 Accepted
        logger.info(f"Desarchivado de equipo '{team_id}' iniciado.")
        return {"status": "Desarchivado Iniciado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en unarchivar_equipo {team_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en unarchivar_equipo {team_id}: {e}", exc_info=True)
        raise

def eliminar_equipo(headers: Dict[str, str], team_id: str) -> dict:
    """Elimina un equipo (permanentemente). Requiere headers."""
    # Requiere Group.ReadWrite.All
    url = f"{BASE_URL}/groups/{team_id}" # La eliminación es sobre el grupo subyacente
    response: Optional[requests.Response] = None
    try:
        logger.warning(f"¡ACCIÓN PELIGROSA! Eliminando grupo/equipo '{team_id}'")
        logger.info(f"API Call: DELETE {url}")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT * 2)
        response.raise_for_status() # 204 No Content
        logger.info(f"Equipo/Grupo '{team_id}' eliminado.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en eliminar_equipo {team_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_equipo {team_id}: {e}", exc_info=True)
        raise

def listar_canales(headers: Dict[str, str], team_id: str) -> dict:
    """Lista los canales de un equipo. Requiere headers."""
    url = f"{BASE_URL}/teams/{team_id}/channels"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Listando canales equipo '{team_id}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} canales del equipo '{team_id}'.")
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en listar_canales {team_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_canales {team_id}: {e}", exc_info=True)
        raise

def obtener_canal(headers: Dict[str, str], team_id: str, channel_id: str) -> dict:
    """Obtiene información sobre un canal específico. Requiere headers."""
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo canal '{channel_id}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Obtenido canal '{channel_id}' del equipo '{team_id}'.")
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en obtener_canal {channel_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_canal {channel_id}: {e}", exc_info=True)
        raise

def crear_canal(headers: Dict[str, str], team_id: str, nombre_canal: str, descripcion: str = "", tipo_canal: str = "standard") -> dict:
    """Crea un nuevo canal en un equipo. Requiere headers."""
    # Requiere Teamwork.Migrate.All, Channel.Create, Group.ReadWrite.All? Revisar permisos
    url = f"{BASE_URL}/teams/{team_id}/channels"
    body = {
        "displayName": nombre_canal,
        "description": descripcion,
        # El valor debe ser standard, private, o shared (enum)
        "membershipType": tipo_canal.lower() if tipo_canal.lower() in ["standard", "private", "shared"] else "standard"
    }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando canal '{nombre_canal}' en equipo '{team_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 201 Created
        data = response.json()
        channel_id = data.get('id')
        logger.info(f"Canal '{nombre_canal}' creado en equipo '{team_id}' con ID: {channel_id}.")
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en crear_canal: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_canal: {e}", exc_info=True)
        raise

def actualizar_canal(headers: Dict[str, str], team_id: str, channel_id: str, nuevos_valores: dict) -> dict:
    """Actualiza la información de un canal. Requiere headers."""
    # Requiere ChannelSettings.ReadWrite.All, Group.ReadWrite.All?
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Actualizando canal '{channel_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        # Filtrar campos no actualizables? La API devuelve error si se intentan
        response = requests.patch(url, headers=current_headers, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 204 No Content
        logger.info(f"Canal '{channel_id}' actualizado en equipo '{team_id}'.")
        return {"status": "Canal Actualizado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en actualizar_canal {channel_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en actualizar_canal {channel_id}: {e}", exc_info=True)
        raise

def eliminar_canal(headers: Dict[str, str], team_id: str, channel_id: str) -> dict:
    """Elimina un canal. Requiere headers."""
    # Requiere Channel.Delete.All, Group.ReadWrite.All?
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"
    response: Optional[requests.Response] = None
    try:
        logger.warning(f"Eliminando canal '{channel_id}' del equipo '{team_id}'")
        logger.info(f"API Call: DELETE {url}")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 204 No Content
        logger.info(f"Canal '{channel_id}' eliminado del equipo '{team_id}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en eliminar_canal {channel_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_canal {channel_id}: {e}", exc_info=True)
        raise

# ---- MENSAJES DE CANAL ----

def enviar_mensaje_canal(headers: Dict[str, str], team_id: str, channel_id: str, mensaje: str, tipo_contenido: str = "text") -> dict:
    """Envía un mensaje a un canal de Teams. Requiere headers."""
    # Requiere ChannelMessage.Send, Group.ReadWrite.All?
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}/messages"
    body = {
        "body": {
            "contentType": tipo_contenido.lower(), # 'text' o 'html'
            "content": mensaje
        }
    }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Enviando mensaje a canal '{channel_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 201 Created
        data = response.json()
        logger.info(f"Mensaje enviado al canal '{channel_id}' del equipo '{team_id}'.")
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en enviar_mensaje_canal {channel_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en enviar_mensaje_canal {channel_id}: {e}", exc_info=True)
        raise
