# actions/teams.py (Refactorizado y Corregido - Final)

import logging
import requests
import json
# Corregido: Añadir Any
from typing import Dict, List, Optional, Union, Any
from datetime import datetime, timezone

# Usar el logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar helper y constantes
try:
    from helpers.http_client import hacer_llamada_api
    from shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    logger.error("Error importando helpers/constantes en Teams.")
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs): raise NotImplementedError("Helper no importado")

# ---- CHAT ----
# Usan /me/chats o /chats/{id}, requieren headers delegados

def listar_chats(headers: Dict[str, str], top: int = 20, skip: int = 0, filter_query: Optional[str] = None, order_by: Optional[str] = None, expand: Optional[str] = None) -> dict:
    """Lista los chats del usuario actual (/me). Requiere headers delegados."""
    url = f"{BASE_URL}/me/chats"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if filter_query: params['$filter'] = filter_query
    if order_by: params['$orderby'] = order_by
    if expand: params['$expand'] = expand
    clean_params = {k:v for k, v in params.items() if v is not None}
    logger.info(f"Listando chats /me con params: {clean_params}")
    # La paginación aquí es compleja (usa skiptoken), usar helper para la primera página
    return hacer_llamada_api("GET", url, headers, params=clean_params)

def obtener_chat(headers: Dict[str, str], chat_id: str) -> dict:
    """Obtiene un chat específico por ID. Requiere headers."""
    url = f"{BASE_URL}/chats/{chat_id}"; logger.info(f"Obteniendo chat '{chat_id}'")
    return hacer_llamada_api("GET", url, headers)

def crear_chat(headers: Dict[str, str], miembros: List[Dict[str, Any]], tipo_chat: str = "oneOnOne", tema: Optional[str] = None) -> dict:
    """Crea un nuevo chat. Requiere headers."""
    url = f"{BASE_URL}/chats"; body: Dict[str, Any] = {"chatType": tipo_chat, "members": miembros};
    if tema: body["topic"] = tema
    logger.info(f"Creando chat tipo '{tipo_chat}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def enviar_mensaje_chat(headers: Dict[str, str], chat_id: str, mensaje: str, tipo_contenido: str = "text") -> dict:
    """Envía un mensaje a un chat. Requiere headers."""
    url = f"{BASE_URL}/chats/{chat_id}/messages"; body = {"body": {"contentType": tipo_contenido.lower(), "content": mensaje}};
    logger.info(f"Enviando mensaje a chat '{chat_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def obtener_mensajes_chat(headers: Dict[str, str], chat_id: str, top: int = 20, skip: int = 0) -> dict:
    """Obtiene los mensajes de un chat (con paginación simple $skip). Requiere headers."""
    url_base = f"{BASE_URL}/chats/{chat_id}/messages"; params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip), '$orderby': 'createdDateTime desc'}
    # Corregido: Anotar all_messages
    all_messages: List[Dict[str, Any]] = []; current_url: Optional[str] = url_base; response: Optional[requests.Response] = None; current_headers = headers.copy()
    try: # Paginación simple con $skip (Graph usa a veces $skiptoken que es más complejo)
         # Limitaremos a obtener solo la primera página por simplicidad del helper
         logger.info(f"Obteniendo mensajes chat '{chat_id}' (top {top}, skip {skip})")
         data = hacer_llamada_api("GET", current_url, current_headers, params=params)
         return data # Devuelve la primera página obtenida
        # Para paginación real con $skiptoken se necesitaría bucle y requests directo
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en obtener_mensajes_chat {chat_id}: {e}", exc_info=True); raise Exception(f"Error API obteniendo mensajes chat: {e}")
    except Exception as e: logger.error(f"Error inesperado en obtener_mensajes_chat {chat_id}: {e}", exc_info=True); raise

def actualizar_mensaje_chat(headers: Dict[str, str], chat_id: str, message_id: str, contenido: str, tipo_contenido: str = "text") -> Optional[Dict[str, Any]]:
    """Actualiza un mensaje existente en un chat. Requiere headers."""
    url = f"{BASE_URL}/chats/{chat_id}/messages/{message_id}"; body = {"body": {"contentType": tipo_contenido.lower(), "content": contenido}};
    logger.info(f"Actualizando mensaje '{message_id}' en chat '{chat_id}'")
    # PATCH devuelve 204 No Content (None del helper)
    hacer_llamada_api("PATCH", url, headers, json_data=body)
    return {"status": "Mensaje Actualizado"}

def eliminar_mensaje_chat(headers: Dict[str, str], chat_id: str, message_id: str) -> Optional[Dict[str, Any]]:
    """Elimina un mensaje de un chat (soft delete). Requiere headers."""
    url = f"{BASE_URL}/me/chats/{chat_id}/messages/{message_id}/softDelete";
    logger.info(f"Soft deleting mensaje '{message_id}' en chat '{chat_id}'")
    # POST sin body, devuelve 204 (None del helper)
    hacer_llamada_api("POST", url, headers)
    return {"status": "Mensaje Eliminado (Soft)"}

# ---- EQUIPOS Y CANALES ----
def listar_equipos(headers: Dict[str, str], top: int = 20, skip: int = 0, filter_query: Optional[str] = None) -> dict:
    """Lista los equipos a los que pertenece el usuario actual (/me). Requiere headers."""
    url = f"{BASE_URL}/me/joinedTeams"; params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)};
    if filter_query: params['$filter'] = filter_query
    clean_params = {k:v for k, v in params.items() if v is not None}; logger.info(f"Listando equipos /me")
    # Paginación con $skiptoken, usar helper para primera página
    return hacer_llamada_api("GET", url, headers, params=clean_params)

def obtener_equipo(headers: Dict[str, str], team_id: str) -> dict:
    """Obtiene información sobre un equipo específico por ID. Requiere headers."""
    url = f"{BASE_URL}/teams/{team_id}"; logger.info(f"Obteniendo equipo '{team_id}'")
    return hacer_llamada_api("GET", url, headers)

def crear_equipo(headers: Dict[str, str], nombre: str, descripcion: str = "Equipo creado por API", tipo_plantilla: str = "standard") -> dict:
    """Crea un nuevo equipo de Microsoft Teams. Requiere headers."""
    url = f"{BASE_URL}/teams"; body = {"template@odata.bind": f"{BASE_URL}/teamsTemplates('{tipo_plantilla}')", "displayName": nombre, "description": descripcion};
    logger.info(f"Creando equipo '{nombre}'")
    # Creación puede ser asíncrona (202), el helper no maneja 'Location' header. Llamada directa:
    try:
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT * 2)
        response.raise_for_status()
        if response.status_code == 202:
             monitor_url = response.headers.get('Location'); logger.info(f"Creación de equipo '{nombre}' iniciada. Monitor: {monitor_url}"); return {"status": "Creación Iniciada", "code": response.status_code, "monitorUrl": monitor_url}
        else:
            data = response.json(); team_id = data.get('id'); logger.info(f"Equipo '{nombre}' creado ID: {team_id}."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en crear_equipo: {e}", exc_info=True); raise Exception(f"Error API creando equipo: {e}")
    except Exception as e: logger.error(f"Error inesperado en crear_equipo: {e}", exc_info=True); raise

def archivar_equipo(headers: Dict[str, str], team_id: str, set_frozen: bool = False) -> Optional[Dict[str, Any]]:
    """Archiva un equipo. Requiere headers."""
    url = f"{BASE_URL}/teams/{team_id}/archive"; body = {"shouldSetSpoSiteReadOnlyForUsers": set_frozen};
    logger.info(f"Archivando equipo '{team_id}'")
    # Devuelve 202 (None del helper)
    hacer_llamada_api("POST", url, headers, json_data=body)
    return {"status": "Archivado Iniciado"}

def unarchivar_equipo(headers: Dict[str, str], team_id: str) -> Optional[Dict[str, Any]]:
    """Desarchiva un equipo. Requiere headers."""
    url = f"{BASE_URL}/teams/{team_id}/unarchive"; logger.info(f"Desarchivando equipo '{team_id}'")
    # POST sin body, devuelve 202 (None del helper)
    hacer_llamada_api("POST", url, headers)
    return {"status": "Desarchivado Iniciado"}

def eliminar_equipo(headers: Dict[str, str], team_id: str) -> Optional[Dict[str, Any]]:
    """Elimina un equipo (permanentemente). Requiere headers."""
    url = f"{BASE_URL}/groups/{team_id}"; logger.warning(f"¡Eliminando grupo/equipo '{team_id}'!")
    hacer_llamada_api("DELETE", url, headers, timeout=GRAPH_API_TIMEOUT * 2) # Devuelve 204 (None)
    return {"status": "Eliminado"}

def listar_canales(headers: Dict[str, str], team_id: str) -> dict:
    """Lista los canales de un equipo. Requiere headers."""
    url = f"{BASE_URL}/teams/{team_id}/channels"; logger.info(f"Listando canales equipo '{team_id}'")
    return hacer_llamada_api("GET", url, headers)

def obtener_canal(headers: Dict[str, str], team_id: str, channel_id: str) -> dict:
    """Obtiene información sobre un canal específico. Requiere headers."""
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"; logger.info(f"Obteniendo canal '{channel_id}'")
    return hacer_llamada_api("GET", url, headers)

def crear_canal(headers: Dict[str, str], team_id: str, nombre_canal: str, descripcion: str = "", tipo_canal: str = "standard") -> dict:
    """Crea un nuevo canal en un equipo. Requiere headers."""
    url = f"{BASE_URL}/teams/{team_id}/channels"; body = {"displayName": nombre_canal, "description": descripcion, "membershipType": tipo_canal.lower() if tipo_canal.lower() in ["standard", "private", "shared"] else "standard"};
    logger.info(f"Creando canal '{nombre_canal}' en equipo '{team_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def actualizar_canal(headers: Dict[str, str], team_id: str, channel_id: str, nuevos_valores: dict) -> Optional[Dict[str, Any]]:
    """Actualiza la información de un canal. Requiere headers."""
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"; logger.info(f"Actualizando canal '{channel_id}'")
    hacer_llamada_api("PATCH", url, headers, json_data=nuevos_valores) # Devuelve 204 (None)
    return {"status": "Canal Actualizado"}

def eliminar_canal(headers: Dict[str, str], team_id: str, channel_id: str) -> Optional[Dict[str, Any]]:
    """Elimina un canal. Requiere headers."""
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"; logger.warning(f"Eliminando canal '{channel_id}' del equipo '{team_id}'")
    hacer_llamada_api("DELETE", url, headers) # Devuelve 204 (None)
    return {"status": "Eliminado"}

# ---- MENSAJES DE CANAL ----
def enviar_mensaje_canal(headers: Dict[str, str], team_id: str, channel_id: str, mensaje: str, tipo_contenido: str = "text") -> dict:
    """Envía un mensaje a un canal de Teams. Requiere headers."""
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}/messages"; body = {"body": {"contentType": tipo_contenido.lower(), "content": mensaje}};
    logger.info(f"Enviando mensaje a canal '{channel_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)
