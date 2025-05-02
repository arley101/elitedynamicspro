# actions/teams.py (Refactorizado v2)

import logging
import requests # Solo para tipos de excepción
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
except ImportError as e:
    # Log crítico y error si falta dependencia esencial
    logging.critical(f"Error CRÍTICO importando helpers/constantes en Teams: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    # No definir mock, dejar que falle si no se importa
    raise ImportError("No se pudo importar 'hacer_llamada_api' desde helpers.") from e

# ============================================
# ==== FUNCIONES DE ACCIÓN PARA CHAT ====
# ============================================
# Usan /me/chats o /chats/{id}, requieren headers delegados (token de usuario)
# Todas usan la firma (parametros: Dict[str, Any], headers: Dict[str, str])

def listar_chats(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista los chats del usuario actual (/me). Requiere token delegado.
    Maneja paginación básica con $top y $skip (Graph puede usar $skiptoken).

    Args:
        parametros (Dict[str, Any]): Opcional: 'top', 'skip', 'filter_query', 'order_by', 'expand'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Respuesta de Graph API.
    """
    top: int = int(parametros.get('top', 20))
    skip: int = int(parametros.get('skip', 0))
    filter_query: Optional[str] = parametros.get('filter_query')
    order_by: Optional[str] = parametros.get('order_by')
    expand: Optional[str] = parametros.get('expand')

    url = f"{BASE_URL}/me/chats"
    params_query: Dict[str, Any] = {'$top': min(top, 50)} # Limitar top
    if skip > 0:
        logger.warning("Usando '$skip' para paginación de chats, puede ser inconsistente.")
        params_query['$skip'] = skip
    if filter_query: params_query['$filter'] = filter_query
    if order_by: params_query['$orderby'] = order_by
    if expand: params_query['$expand'] = expand

    clean_params = {k: v for k, v in params_query.items() if v is not None}
    logger.info(f"Listando chats /me con params: {clean_params}")

    # TODO: Implementar paginación completa usando @odata.nextLink si es necesario.
    return hacer_llamada_api("GET", url, headers, params=clean_params)


def obtener_chat(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene detalles de un chat específico por ID. Requiere token delegado.

    Args:
        parametros (Dict[str, Any]): Debe contener 'chat_id'. Opcional: 'expand'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Objeto del chat.
    """
    chat_id: Optional[str] = parametros.get("chat_id")
    expand: Optional[str] = parametros.get("expand")
    if not chat_id: raise ValueError("Parámetro 'chat_id' es requerido.")

    url = f"{BASE_URL}/chats/{chat_id}"
    params_query: Dict[str, Any] = {}
    if expand: params_query['$expand'] = expand

    logger.info(f"Obteniendo chat '{chat_id}' (Expand: {expand})")
    return hacer_llamada_api("GET", url, headers, params=params_query or None)


def crear_chat(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea un nuevo chat (oneOnOne o group). Requiere token delegado.

    Args:
        parametros (Dict[str, Any]): Debe contener 'miembros' (List[Dict] formato Graph), 'tipo_chat'.
                                     Opcional: 'tema'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Objeto del chat creado.
    """
    miembros: Optional[List[Dict[str, Any]]] = parametros.get("miembros")
    tipo_chat: str = parametros.get("tipo_chat", "oneOnOne")
    tema: Optional[str] = parametros.get("tema")

    if not miembros or not isinstance(miembros, list):
        raise ValueError("Parámetro 'miembros' (lista de diccionarios) es requerido.")
    for i, m in enumerate(miembros):
        if not isinstance(m, dict) or "@odata.type" not in m or "user@odata.bind" not in m:
            raise ValueError(f"Formato inválido para miembro {i+1}.")

    if tipo_chat not in ["oneOnOne", "group"]:
        raise ValueError("Parámetro 'tipo_chat' debe ser 'oneOnOne' o 'group'.")
    # Warnings lógicos
    if tipo_chat == "oneOnOne" and len(miembros) != 2: logger.warning(f"Creando chat 'oneOnOne' con {len(miembros)} miembros.")
    if tipo_chat == "group" and len(miembros) < 3: logger.warning(f"Creando chat 'group' con solo {len(miembros)} miembros.")
    if tema and tipo_chat == "oneOnOne": logger.warning("El 'tema' se ignora para chats 'oneOnOne'.")

    url = f"{BASE_URL}/chats"
    body: Dict[str, Any] = {"chatType": tipo_chat, "members": miembros}
    if tema and tipo_chat == "group": body["topic"] = tema

    logger.info(f"Creando chat tipo '{tipo_chat}' con {len(miembros)} miembros.")
    return hacer_llamada_api("POST", url, headers, json_data=body)


def enviar_mensaje_chat(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Envía un mensaje a un chat existente. Requiere token delegado.

    Args:
        parametros (Dict[str, Any]): Debe contener 'chat_id', 'mensaje'. Opcional: 'tipo_contenido'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Objeto del mensaje enviado.
    """
    chat_id: Optional[str] = parametros.get("chat_id")
    mensaje: Optional[str] = parametros.get("mensaje")
    tipo_contenido: str = parametros.get("tipo_contenido", "text").lower()

    if not chat_id: raise ValueError("Parámetro 'chat_id' es requerido.")
    if not mensaje: raise ValueError("Parámetro 'mensaje' es requerido.")
    if tipo_contenido not in ["text", "html"]: raise ValueError("'tipo_contenido' debe ser 'text' o 'html'.")

    url = f"{BASE_URL}/chats/{chat_id}/messages"
    body = {"body": {"contentType": tipo_contenido, "content": mensaje}}

    logger.info(f"Enviando mensaje ({tipo_contenido}) a chat '{chat_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)


def obtener_mensajes_chat(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene mensajes de un chat, ordenados descendente. Paginación básica. Requiere token delegado.

    Args:
        parametros (Dict[str, Any]): Debe contener 'chat_id'. Opcional: 'top', 'skip'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Respuesta de Graph API.
    """
    chat_id: Optional[str] = parametros.get("chat_id")
    top: int = int(parametros.get('top', 20))
    skip: int = int(parametros.get('skip', 0))

    if not chat_id: raise ValueError("Parámetro 'chat_id' es requerido.")

    url_base = f"{BASE_URL}/chats/{chat_id}/messages"
    params_query: Dict[str, Any] = {'$top': min(top, 50), '$orderby': 'createdDateTime desc'}
    if skip > 0:
        logger.warning("Usando '$skip' para paginación de mensajes chat.")
        params_query['$skip'] = skip

    # TODO: Implementar paginación real con @odata.nextLink si es necesario.
    logger.info(f"Obteniendo mensajes chat '{chat_id}' (Top: {top}, Skip: {skip})")
    return hacer_llamada_api("GET", url_base, headers, params=params_query)


def actualizar_mensaje_chat(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Actualiza contenido de un mensaje en un chat. Requiere token delegado.

    Args:
        parametros (Dict[str, Any]): Debe contener 'chat_id', 'message_id', 'contenido'. Opcional: 'tipo_contenido'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Confirmación.
    """
    chat_id: Optional[str] = parametros.get("chat_id")
    message_id: Optional[str] = parametros.get("message_id")
    contenido: Optional[str] = parametros.get("contenido")
    tipo_contenido: str = parametros.get("tipo_contenido", "text").lower()

    if not chat_id: raise ValueError("Parámetro 'chat_id' es requerido.")
    if not message_id: raise ValueError("Parámetro 'message_id' es requerido.")
    if contenido is None: raise ValueError("Parámetro 'contenido' es requerido.")
    if tipo_contenido not in ["text", "html"]: raise ValueError("'tipo_contenido' debe ser 'text' o 'html'.")

    url = f"{BASE_URL}/chats/{chat_id}/messages/{message_id}"
    body = {"body": {"contentType": tipo_contenido, "content": contenido}}

    logger.info(f"Actualizando mensaje '{message_id}' en chat '{chat_id}'")
    # PATCH devuelve 204 No Content (None del helper).
    hacer_llamada_api("PATCH", url, headers, json_data=body)
    return {"status": "Mensaje Actualizado", "chat_id": chat_id, "message_id": message_id}


def eliminar_mensaje_chat(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina un mensaje de un chat (soft delete). Requiere token delegado y permisos.

    Args:
        parametros (Dict[str, Any]): Debe contener 'chat_id', 'message_id'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Confirmación.
    """
    chat_id: Optional[str] = parametros.get("chat_id")
    message_id: Optional[str] = parametros.get("message_id")

    if not chat_id: raise ValueError("Parámetro 'chat_id' es requerido.")
    if not message_id: raise ValueError("Parámetro 'message_id' es requerido.")

    url = f"{BASE_URL}/me/chats/{chat_id}/messages/{message_id}/softDelete"
    logger.info(f"Soft deleting mensaje '{message_id}' en chat '{chat_id}'")
    # POST sin body, devuelve 204 (None del helper).
    hacer_llamada_api("POST", url, headers)
    return {"status": "Mensaje Eliminado (Soft)", "chat_id": chat_id, "message_id": message_id}


# ======================================================
# ==== FUNCIONES DE ACCIÓN PARA EQUIPOS Y CANALES ====
# ======================================================
# Requieren token delegado con permisos apropiados

def listar_equipos(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista los equipos a los que pertenece el usuario actual (/me/joinedTeams).

    Args:
        parametros (Dict[str, Any]): Opcional: 'top', 'skip', 'filter_query'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Respuesta de Graph API.
    """
    top: int = int(parametros.get('top', 20))
    skip: int = int(parametros.get('skip', 0))
    filter_query: Optional[str] = parametros.get('filter_query')

    url = f"{BASE_URL}/me/joinedTeams"
    params_query: Dict[str, Any] = {'$top': min(top, 999)}
    if skip > 0:
        logger.warning("Usando '$skip' para paginación de equipos.")
        params_query['$skip'] = skip
    if filter_query: params_query['$filter'] = filter_query

    clean_params = {k: v for k, v in params_query.items() if v is not None}
    logger.info(f"Listando equipos unidos por /me con params: {clean_params}")
    # TODO: Implementar paginación con @odata.nextLink si es necesario.
    return hacer_llamada_api("GET", url, headers, params=clean_params)


def obtener_equipo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene información sobre un equipo específico por ID (Group ID).

    Args:
        parametros (Dict[str, Any]): Debe contener 'team_id'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Objeto del equipo.
    """
    team_id: Optional[str] = parametros.get("team_id")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")

    url = f"{BASE_URL}/teams/{team_id}"
    logger.info(f"Obteniendo detalles del equipo '{team_id}'")
    return hacer_llamada_api("GET", url, headers)


def crear_equipo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea un nuevo equipo de Microsoft Teams. Operación asíncrona o síncrona.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_equipo'. Opcional: 'descripcion', 'tipo_plantilla', 'miembros'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Objeto del equipo (si 201) o estado/URL monitor (si 202).
    """
    nombre_equipo: Optional[str] = parametros.get("nombre_equipo")
    descripcion: str = parametros.get("descripcion", f"Equipo {nombre_equipo}")
    tipo_plantilla: str = parametros.get("tipo_plantilla", "standard")
    miembros: Optional[List[Dict[str, Any]]] = parametros.get("miembros")

    if not nombre_equipo: raise ValueError("Parámetro 'nombre_equipo' es requerido.")

    url = f"{BASE_URL}/teams"
    body: Dict[str, Any] = {
        "template@odata.bind": f"{BASE_URL}/teamsTemplates('{tipo_plantilla}')",
        "displayName": nombre_equipo,
        "description": descripcion
    }
    if miembros and isinstance(miembros, list):
        valid_members = [m for m in miembros if isinstance(m, dict) and m.get('@odata.type') == '#microsoft.graph.aadUserConversationMember']
        if valid_members: body["members"] = valid_members
        else: logger.warning("Formato de miembros inválido al crear equipo, se ignorarán.")

    logger.info(f"Solicitando creación de equipo '{nombre_equipo}'")
    # Usar helper con expect_json=False para manejar 201 o 202
    response = hacer_llamada_api("POST", url, headers, json_data=body, timeout=GRAPH_API_TIMEOUT * 2, expect_json=False)

    if isinstance(response, requests.Response):
        if response.status_code == 201: # Creado síncronamente
            try:
                data = response.json(); logger.info(f"Equipo '{nombre_equipo}' creado síncronamente. ID: {data.get('id')}."); return data
            except json.JSONDecodeError: return {"status": "Creado (Sin Cuerpo)", "status_code": 201}
        elif response.status_code == 202: # Creación asíncrona
            monitor_url = response.headers.get('Location'); logger.info(f"Creación de equipo '{nombre_equipo}' iniciada. Monitor: {monitor_url}");
            return {"status": "Creación Iniciada", "status_code": 202, "monitorUrl": monitor_url}
        else:
            logger.error(f"Respuesta inesperada al crear equipo: {response.status_code}. Cuerpo: {response.text[:200]}");
            raise Exception(f"Respuesta inesperada al crear equipo: {response.status_code}")
    else:
        logger.error(f"Respuesta inesperada del helper al crear equipo: {type(response)}");
        raise Exception("Error interno al procesar creación de equipo.")


def archivar_equipo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Archiva un equipo. Operación asíncrona.

    Args:
        parametros (Dict[str, Any]): Debe contener 'team_id'. Opcional: 'set_frozen'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Confirmación.
    """
    team_id: Optional[str] = parametros.get("team_id")
    set_frozen: bool = parametros.get("set_frozen", False)
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")

    url = f"{BASE_URL}/teams/{team_id}/archive"
    body = {"shouldSetSpoSiteReadOnlyForUsers": set_frozen} if set_frozen else None

    logger.info(f"Solicitando archivado del equipo '{team_id}' (Congelar sitio SP: {set_frozen})")
    # Devuelve 202 (None del helper).
    hacer_llamada_api("POST", url, headers, json_data=body)
    return {"status": "Archivado Iniciado", "team_id": team_id}


def unarchivar_equipo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Desarchiva un equipo. Operación asíncrona.

    Args:
        parametros (Dict[str, Any]): Debe contener 'team_id'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Confirmación.
    """
    team_id: Optional[str] = parametros.get("team_id")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")

    url = f"{BASE_URL}/teams/{team_id}/unarchive"
    logger.info(f"Solicitando desarchivado del equipo '{team_id}'")
    # Devuelve 202 (None del helper).
    hacer_llamada_api("POST", url, headers)
    return {"status": "Desarchivado Iniciado", "team_id": team_id}


def eliminar_equipo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina un equipo (elimina el Grupo M365 asociado). ¡PERMANENTE!

    Args:
        parametros (Dict[str, Any]): Debe contener 'team_id'.
        headers (Dict[str, str]): Cabeceras con token delegado y permisos suficientes.

    Returns:
        Dict[str, Any]: Confirmación.
    """
    team_id: Optional[str] = parametros.get("team_id")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")

    url = f"{BASE_URL}/groups/{team_id}"
    logger.warning(f"¡¡¡SOLICITANDO ELIMINACIÓN PERMANENTE DEL GRUPO/EQUIPO '{team_id}'!!!")
    delete_timeout = max(GRAPH_API_TIMEOUT, 90)
    # Devuelve 204 (None del helper).
    hacer_llamada_api("DELETE", url, headers, timeout=delete_timeout)
    return {"status": "Eliminado Permanentemente", "id": team_id}


def listar_canales(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista los canales de un equipo específico.

    Args:
        parametros (Dict[str, Any]): Debe contener 'team_id'. Opcional: 'filter_query'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Respuesta de Graph API.
    """
    team_id: Optional[str] = parametros.get("team_id")
    filter_query: Optional[str] = parametros.get("filter_query")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")

    url = f"{BASE_URL}/teams/{team_id}/channels"
    params_query: Dict[str, Any] = {}
    if filter_query: params_query['$filter'] = filter_query

    logger.info(f"Listando canales para equipo '{team_id}' (Filtro: {filter_query})")
    return hacer_llamada_api("GET", url, headers, params=params_query or None)


def obtener_canal(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene información detallada sobre un canal específico.

    Args:
        parametros (Dict[str, Any]): Debe contener 'team_id', 'channel_id'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Objeto del canal.
    """
    team_id: Optional[str] = parametros.get("team_id")
    channel_id: Optional[str] = parametros.get("channel_id")
    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    if not channel_id: raise ValueError("Parámetro 'channel_id' es requerido.")

    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"
    logger.info(f"Obteniendo detalles del canal '{channel_id}' en equipo '{team_id}'")
    return hacer_llamada_api("GET", url, headers)


def crear_canal(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea un nuevo canal (standard, private, shared) en un equipo.

    Args:
        parametros (Dict[str, Any]): Debe contener 'team_id', 'nombre_canal'. Opcional: 'descripcion', 'tipo_canal', 'miembros'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Objeto del canal creado.
    """
    team_id: Optional[str] = parametros.get("team_id")
    nombre_canal: Optional[str] = parametros.get("nombre_canal")
    descripcion: str = parametros.get("descripcion", "")
    tipo_canal: str = parametros.get("tipo_canal", "standard").lower()
    miembros: Optional[List[Dict[str, Any]]] = parametros.get("miembros")

    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    if not nombre_canal: raise ValueError("Parámetro 'nombre_canal' es requerido.")
    if tipo_canal not in ["standard", "private", "shared"]: raise ValueError("'tipo_canal' debe ser 'standard', 'private' o 'shared'.")
    if tipo_canal in ["private", "shared"] and not (miembros and isinstance(miembros, list)):
         raise ValueError(f"Parámetro 'miembros' (lista) es requerido para crear canal '{tipo_canal}'.")

    url = f"{BASE_URL}/teams/{team_id}/channels"
    body: Dict[str, Any] = {"displayName": nombre_canal, "description": descripcion, "membershipType": tipo_canal}
    if tipo_canal in ["private", "shared"] and miembros:
        valid_members = [m for m in miembros if isinstance(m, dict) and m.get('@odata.type') == '#microsoft.graph.aadUserConversationMember']
        if not valid_members: raise ValueError(f"Se requieren miembros válidos para crear canal '{tipo_canal}'.")
        body["members"] = valid_members

    logger.info(f"Creando canal '{nombre_canal}' (tipo: {tipo_canal}) en equipo '{team_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)


def actualizar_canal(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Actualiza la información de un canal.

    Args:
        parametros (Dict[str, Any]): Debe contener 'team_id', 'channel_id', 'nuevos_valores' (dict).
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Confirmación.
    """
    team_id: Optional[str] = parametros.get("team_id")
    channel_id: Optional[str] = parametros.get("channel_id")
    nuevos_valores: Optional[Dict[str, Any]] = parametros.get("nuevos_valores")

    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    if not channel_id: raise ValueError("Parámetro 'channel_id' es requerido.")
    if not nuevos_valores or not isinstance(nuevos_valores, dict): raise ValueError("'nuevos_valores' (diccionario) es requerido.")

    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"
    logger.info(f"Actualizando canal '{channel_id}' en equipo '{team_id}'")
    # PATCH devuelve 204 No Content (None del helper).
    hacer_llamada_api("PATCH", url, headers, json_data=nuevos_valores)
    return {"status": "Canal Actualizado", "id": channel_id}


def eliminar_canal(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina un canal de un equipo.

    Args:
        parametros (Dict[str, Any]): Debe contener 'team_id', 'channel_id'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Confirmación.
    """
    team_id: Optional[str] = parametros.get("team_id")
    channel_id: Optional[str] = parametros.get("channel_id")

    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    if not channel_id: raise ValueError("Parámetro 'channel_id' es requerido.")

    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}"
    logger.warning(f"Eliminando canal '{channel_id}' del equipo '{team_id}'")
    # DELETE devuelve 204 No Content (None del helper).
    hacer_llamada_api("DELETE", url, headers)
    return {"status": "Eliminado", "id": channel_id}


def enviar_mensaje_canal(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Envía un mensaje a un canal específico de un equipo.

    Args:
        parametros (Dict[str, Any]): Debe contener 'team_id', 'channel_id', 'mensaje'. Opcional: 'tipo_contenido'.
        headers (Dict[str, str]): Cabeceras con token delegado.

    Returns:
        Dict[str, Any]: Objeto del mensaje enviado.
    """
    team_id: Optional[str] = parametros.get("team_id")
    channel_id: Optional[str] = parametros.get("channel_id")
    mensaje: Optional[str] = parametros.get("mensaje")
    tipo_contenido: str = parametros.get("tipo_contenido", "text").lower()

    if not team_id: raise ValueError("Parámetro 'team_id' es requerido.")
    if not channel_id: raise ValueError("Parámetro 'channel_id' es requerido.")
    if not mensaje: raise ValueError("Parámetro 'mensaje' es requerido.")
    if tipo_contenido not in ["text", "html"]: raise ValueError("'tipo_contenido' debe ser 'text' o 'html'.")

    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}/messages"
    body = {"body": {"contentType": tipo_contenido, "content": mensaje}}

    logger.info(f"Enviando mensaje ({tipo_contenido}) a canal '{channel_id}' en equipo '{team_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

# --- FIN DEL MÓDULO actions/teams.py ---
