# actions/teams.py (Refactorizado v2)

import logging
import requests # Solo para tipos de excepción y llamadas directas
import json
from typing import Dict, List, Optional, Union, Any
from datetime import datetime, timezone

# Usar el logger estándar de Azure Functions
logger = logging.getLogger("azure.functions")

# Importar helper y constantes desde la estructura compartida
try:
    # Asume que shared está un nivel arriba de actions
    from ..shared.helpers.http_client import hacer_llamada_api
    from ..shared.constants import BASE_URL, GRAPH_API_TIMEOUT
# CORRECCIÓN: No definir mock aquí, solo loguear/lanzar error si falla import
except ImportError as e:
    logging.critical(f"Error CRÍTICO importando helpers/constantes en Teams: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    # Lanzar excepción para detener la carga del módulo si las dependencias no están
    raise ImportError(f"No se pudieron importar dependencias necesarias para actions.teams: {e}") from e

# ============================================
# ==== FUNCIONES DE ACCIÓN PARA CHAT ====
# ============================================
def listar_chats(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Lista los chats del usuario actual (/me)."""
    top: int = int(parametros.get('top', 20))
    skip: int = int(parametros.get('skip', 0))
    filter_query: Optional[str] = parametros.get('filter_query')
    order_by: Optional[str] = parametros.get('order_by')
    expand: Optional[str] = parametros.get('expand')

    url = f"{BASE_URL}/me/chats"
    params_query: Dict[str, Any] = {'$top': min(top, 50)}
    if skip > 0: logger.warning("Usando '$skip' para paginación de chats."); params_query['$skip'] = skip
    if filter_query: params_query['$filter'] = filter_query
    if order_by: params_query['$orderby'] = order_by
    if expand: params_query['$expand'] = expand

    clean_params = {k: v for k, v in params_query.items() if v is not None}
    logger.info(f"Listando chats /me con params: {clean_params}")
    return hacer_llamada_api("GET", url, headers, params=clean_params)

def obtener_chat(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Obtiene un chat específico por ID."""
    chat_id: Optional[str] = parametros.get("chat_id")
    expand: Optional[str] = parametros.get("expand")
    if not chat_id: raise ValueError("Parámetro 'chat_id' es requerido.")

    url = f"{BASE_URL}/chats/{chat_id}"
    params_query: Dict[str, Any] = {}
    if expand: params_query['$expand'] = expand

    logger.info(f"Obteniendo chat '{chat_id}' (Expand: {expand})")
    return hacer_llamada_api("GET", url, headers, params=params_query or None)

def crear_chat(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Crea un nuevo chat (oneOnOne o group)."""
    miembros: Optional[List[Dict[str, Any]]] = parametros.get("miembros")
    tipo_chat: str = parametros.get("tipo_chat", "oneOnOne")
    tema: Optional[str] = parametros.get("tema")
    if not miembros or not isinstance(miembros, list): raise ValueError("Parámetro 'miembros' (lista) es requerido.")
    for i, m in enumerate(miembros):
        if not isinstance(m, dict) or "@odata.type" not in m or "user@odata.bind" not in m: raise ValueError(f"Formato inválido para miembro {i+1}.")
    if tipo_chat not in ["oneOnOne", "group"]: raise ValueError("Parámetro 'tipo_chat' debe ser 'oneOnOne' o 'group'.")

    url = f"{BASE_URL}/chats"; body: Dict[str, Any] = {"chatType": tipo_chat, "members": miembros};
    if tema and tipo_chat == "group": body["topic"] = tema
    logger.info(f"Creando chat tipo '{tipo_chat}' con {len(miembros)} miembros.")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def enviar_mensaje_chat(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Envía un mensaje a un chat existente."""
    chat_id: Optional[str] = parametros.get("chat_id"); mensaje: Optional[str] = parametros.get("mensaje")
    tipo_contenido: str = parametros.get("tipo_contenido", "text").lower()
    if not chat_id: raise ValueError("Parámetro 'chat_id' es requerido.")
    if not mensaje: raise ValueError("Parámetro 'mensaje' es requerido.")
    if tipo_contenido not in ["text", "html"]: raise ValueError("Parámetro 'tipo_contenido' debe ser 'text' o 'html'.")

    url = f"{BASE_URL}/chats/{chat_id}/messages"; body = {"body": {"contentType": tipo_contenido, "content": mensaje}};
    logger.info(f"Enviando mensaje ({tipo_contenido}) a chat '{chat_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def obtener_mensajes_chat(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Obtiene los mensajes de un chat, ordenados por fecha descendente."""
    chat_id: Optional[str] = parametros.get("chat_id"); top: int = int(parametros.get('top', 20)); skip: int = int(parametros.get('skip', 0))
    if not chat_id: raise ValueError("Parámetro 'chat_id' es requerido.")

    url_base = f"{BASE_URL}/chats/{chat_id}/messages"; params_query: Dict[str, Any] = {'$top': min(top, 50), '$orderby': 'createdDateTime desc'}
    if skip > 0: logger.warning("Usando '$skip' para paginación de mensajes."); params_query['$skip'] = skip

    # CORRECCIÓN: Añadir tipo explícito a all_messages
    all_messages: List[Dict[str, Any]] = []; current_url: Optional[str] = url_base; page_count = 0; max_pages = 100
    try:
        while current_url and page_count < max_pages:
            page_count += 1; logger.info(f"Obteniendo mensajes chat '{chat_id}', Página: {page_count}")
            current_params_page = params_query if page_count == 1 else None
            # CORRECCIÓN: Añadir assert para MyPy
            assert current_url is not None
            data = hacer_llamada_api("GET", current_url, headers, params=current_params_page)
            if data:
                page_items = data.get('value', []); all_messages.extend(page_items)
                current_url = data.get('@odata.nextLink');
                if not current_url: break
            else: break
        if page_count >= max_pages: logger.warning(f"Límite de {max_pages} páginas alcanzado.")
        logger.info(f"Total mensajes chat '{chat_id}': {len(all_messages)}")
        return {'value': all_messages}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en obtener_mensajes_chat {chat_id}: {e}", exc_info=True); raise Exception(f"Error API obteniendo mensajes chat: {e}") from e
    except Exception as e: logger.error(f"Error inesperado en obtener_mensajes_chat {chat_id}: {e}", exc_info=True); raise

def actualizar_mensaje_chat(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Actualiza el contenido de un mensaje existente en un chat."""
    chat_id: Optional[str] = parametros.get("chat_id"); message_id: Optional[str] = parametros.get("message_id"); contenido: Optional[str] = parametros.get("contenido")
    tipo_contenido: str = parametros.get("tipo_contenido", "text").lower()
    if not chat_id: raise ValueError("Parámetro 'chat_id' es requerido.")
    if not message_id: raise ValueError("Parámetro 'message_id' es requerido.")
    if contenido is None: raise ValueError("Parámetro 'contenido' es requerido.")
    if tipo_contenido not in ["text", "html"]: raise ValueError("Parámetro 'tipo_contenido' debe ser 'text' o 'html'.")

    url = f"{BASE_URL}/chats/{chat_id}/messages/{message_id}"; body = {"body": {"contentType": tipo_contenido, "content": contenido}};
    logger.info(f"Actualizando mensaje '{message_id}' en chat '{chat_id}'")
    hacer_llamada_api("PATCH", url, headers, json_data=body)
    return {"status": "Mensaje Actualizado"}

def eliminar_mensaje_chat(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Elimina un mensaje de un chat (soft delete)."""
    chat_id: Optional[str] = parametros.get("chat_id"); message_id: Optional[str] = parametros.get("message_id")
    if not chat_id: raise ValueError("Parámetro 'chat_id' es requerido.")
    if not message_id: raise ValueError("Parámetro 'message_id' es requerido.")

    url = f"{BASE_URL}/me/chats/{chat_id}/messages/{message_id}/softDelete";
    logger.info(f"Soft deleting mensaje '{message_id}' en chat '{chat_id}'")
    hacer_llamada_api("POST", url, headers)
    return {"status": "Mensaje Eliminado (Soft)"}

# ======================================================
# ==== FUNCIONES DE ACCIÓN PARA EQUIPOS Y CANALES ====
# ======================================================
def listar_equipos(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Lista los equipos a los que pertenece el usuario actual (/me/joinedTeams)."""
    top: int = int(parametros.get('top', 20)); skip: int = int(parametros.get('skip', 0)); filter_query: Optional[str] = parametros.get('filter_query')
    url = f"{BASE_URL}/me/joinedTeams"; params_query: Dict[str, Any] = {'$top': min(top, 999)};
    if skip > 0: logger.warning("Usando '$skip' para paginación de equipos."); params_query['$skip'] = skip
    if filter_query: params_query['$filter'] = filter_query
    clean_params = {k:v for k, v in params_query.items() if v is not None}; logger.info(f"Listando equipos /me")
    return hacer_llamada_api("GET", url, headers, params=clean_params)

def obtener_equipo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Obtiene información sobre un equipo específico por ID (Group ID)."""
    team_id: Optional[str] = parametros.get("team_id")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    url = f"{BASE_URL}/teams/{team_id}"; logger.info(f"Obteniendo equipo '{team_id}'")
    return hacer_llamada_api("GET", url, headers)

def crear_equipo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Crea un nuevo equipo de Microsoft Teams. Operación asíncrona."""
    nombre_equipo: Optional[str] = parametros.get("nombre_equipo"); descripcion: str = parametros.get("descripcion", f"Equipo {nombre_equipo}"); tipo_plantilla: str = parametros.get("tipo_plantilla", "standard"); miembros: Optional[List[Dict[str, Any]]] = parametros.get("miembros")
    if not nombre_equipo: raise ValueError("Parámetro 'nombre_equipo' es requerido.")

    url = f"{BASE_URL}/teams"; body: Dict[str, Any] = {"template@odata.bind": f"{BASE_URL}/teamsTemplates('{tipo_plantilla}')", "displayName": nombre_equipo, "description": descripcion};
    if miembros and isinstance(miembros, list):
        valid_members = [m for m in miembros if isinstance(m, dict) and m.get('@odata.type') == '#microsoft.graph.aadUserConversationMember' and 'user@odata.bind' in m and 'roles' in m]
        if valid_members: body["members"] = valid_members
        else: logger.warning("Formato de miembros inválido al crear equipo.")
    logger.info(f"Solicitando creación de equipo '{nombre_equipo}'")
    # Usar helper con expect_json=False para manejar 201 y 202
    response = hacer_llamada_api("POST", url, headers, json_data=body, timeout=GRAPH_API_TIMEOUT * 2, expect_json=False)
    if isinstance(response, requests.Response):
        if response.status_code == 201:
            try: data = response.json(); logger.info(f"Equipo '{nombre_equipo}' creado síncronamente."); return data
            except json.JSONDecodeError: return {"status": "Creado (Sin Cuerpo)", "status_code": 201}
        elif response.status_code == 202:
            monitor_url = response.headers.get('Location'); logger.info(f"Creación equipo '{nombre_equipo}' iniciada. Monitor: {monitor_url}"); return {"status": "Creación Iniciada", "status_code": response.status_code, "monitorUrl": monitor_url}
        else: raise Exception(f"Respuesta inesperada al crear equipo: {response.status_code}")
    else: raise Exception("Error interno al procesar creación de equipo.")

def archivar_equipo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Archiva un equipo. Operación asíncrona."""
    team_id: Optional[str] = parametros.get("team_id"); set_frozen: bool = parametros.get("set_frozen", False)
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    url = f"{BASE_URL}/teams/{team_id}/archive"; body = {"shouldSetSpoSiteReadOnlyForUsers": set_frozen} if set_frozen else None;
    logger.info(f"Solicitando archivado equipo '{team_id}'")
    hacer_llamada_api("POST", url, headers, json_data=body)
    return {"status": "Archivado Iniciado", "team_id": team_id}

def unarchivar_equipo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Desarchiva un equipo. Operación asíncrona."""
    team_id: Optional[str] = parametros.get("team_id")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    url = f"{BASE_URL}/teams/{team_id}/unarchive"; logger.info(f"Solicitando desarchivado equipo '{team_id}'")
    hacer_llamada_api("POST", url, headers)
    return {"status": "Desarchivado Iniciado", "team_id": team_id}

def eliminar_equipo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Elimina un equipo (elimina el Grupo M365 asociado). ¡PERMANENTE!"""
    team_id: Optional[str] = parametros.get("team_id")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    url = f"{BASE_URL}/groups/{team_id}"; logger.warning(f"¡ELIMINANDO GRUPO/EQUIPO '{team_id}'!")
    hacer_llamada_api("DELETE", url, headers, timeout=GRAPH_API_TIMEOUT * 2)
    return {"status": "Eliminado Permanentemente", "id": team_id}

def listar_canales(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Lista los canales de un equipo específico."""
    team_id: Optional[str] = parametros.get("team_id"); filter_query: Optional[str] = parametros.get("filter_query")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    url = f"{BASE_URL}/teams/{team_id}/channels"; params_query: Dict[str, Any] = {};
    if filter_query: params_query['$filter'] = filter_query
    logger.info(f"Listando canales equipo '{team_id}'")
    return hacer_llamada_api("GET", url, headers, params=params_query or None)

def obtener_canal(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Obtiene información detallada sobre un canal específico."""
    team_id: Optional[str] = parametros.get("team_id"); channel_id: Optional[str] = parametros.get("channel_id")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    if not channel_id: raise ValueError("Parámetro 'channel_id' es requerido.")
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"; logger.info(f"Obteniendo canal '{channel_id}'")
    return hacer_llamada_api("GET", url, headers)

def crear_canal(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Crea un nuevo canal (standard, private, shared) en un equipo."""
    team_id: Optional[str] = parametros.get("team_id"); nombre_canal: Optional[str] = parametros.get("nombre_canal"); descripcion: str = parametros.get("descripcion", ""); tipo_canal: str = parametros.get("tipo_canal", "standard").lower(); miembros: Optional[List[Dict[str, Any]]] = parametros.get("miembros")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    if not nombre_canal: raise ValueError("Parámetro 'nombre_canal' es requerido.")
    if tipo_canal not in ["standard", "private", "shared"]: raise ValueError("Tipo canal debe ser 'standard', 'private' o 'shared'.")
    if tipo_canal in ["private", "shared"] and (not miembros or not isinstance(miembros, list)): raise ValueError(f"Parámetro 'miembros' (lista) requerido para canal '{tipo_canal}'.")

    url = f"{BASE_URL}/teams/{team_id}/channels"; body: Dict[str, Any] = {"displayName": nombre_canal, "description": descripcion, "membershipType": tipo_canal};
    if tipo_canal in ["private", "shared"] and miembros:
        valid_members = [m for m in miembros if isinstance(m, dict) and m.get('@odata.type') == '#microsoft.graph.aadUserConversationMember' and 'user@odata.bind' in m and 'roles' in m]
        if not valid_members: raise ValueError(f"Miembros válidos requeridos para canal '{tipo_canal}'.")
        body["members"] = valid_members
    logger.info(f"Creando canal '{nombre_canal}' (tipo: {tipo_canal}) en equipo '{team_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def actualizar_canal(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Actualiza la información de un canal."""
    team_id: Optional[str] = parametros.get("team_id"); channel_id: Optional[str] = parametros.get("channel_id"); nuevos_valores: Optional[Dict[str, Any]] = parametros.get("nuevos_valores")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    if not channel_id: raise ValueError("Parámetro 'channel_id' es requerido.")
    if not nuevos_valores or not isinstance(nuevos_valores, dict): raise ValueError("Parámetro 'nuevos_valores' (diccionario) es requerido.")

    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"; logger.info(f"Actualizando canal '{channel_id}'")
    hacer_llamada_api("PATCH", url, headers, json_data=nuevos_valores)
    return {"status": "Canal Actualizado", "id": channel_id}

def eliminar_canal(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Elimina un canal de un equipo."""
    team_id: Optional[str] = parametros.get("team_id"); channel_id: Optional[str] = parametros.get("channel_id")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    if not channel_id: raise ValueError("Parámetro 'channel_id' es requerido.")
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"; logger.warning(f"Eliminando canal '{channel_id}' del equipo '{team_id}'")
    hacer_llamada_api("DELETE", url, headers)
    return {"status": "Eliminado", "id": channel_id}

# ============================================
# ==== FUNCIONES DE MENSAJES DE CANAL ====
# ============================================
def enviar_mensaje_canal(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Envía un mensaje a un canal específico de un equipo."""
    team_id: Optional[str] = parametros.get("team_id"); channel_id: Optional[str] = parametros.get("channel_id"); mensaje: Optional[str] = parametros.get("mensaje")
    tipo_contenido: str = parametros.get("tipo_contenido", "text").lower()
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    if not channel_id: raise ValueError("Parámetro 'channel_id' es requerido.")
    if not mensaje: raise ValueError("Parámetro 'mensaje' es requerido.")
    if tipo_contenido not in ["text", "html"]: raise ValueError("Tipo contenido debe ser 'text' o 'html'.")

    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}/messages"; body = {"body": {"contentType": tipo_contenido, "content": mensaje}};
    logger.info(f"Enviando mensaje ({tipo_contenido}) a canal '{channel_id}' en equipo '{team_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

# --- FIN DEL MÓDULO actions/teams.py ---
