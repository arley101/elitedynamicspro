# actions/correo.py (Refactorizado)

import logging
import requests # Solo para tipos de excepción
import json
from typing import Dict, List, Optional, Union, Any

# Usar logger estándar de Azure Functions
logger = logging.getLogger("azure.functions")

# Importar helper y constantes desde la estructura compartida
try:
    # Asume que shared está un nivel arriba de actions
    from ..shared.helpers.http_client import hacer_llamada_api
    from ..shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError as e:
    logging.critical(f"Error CRÍTICO importando helpers/constantes en Correo: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs):
        raise NotImplementedError("Dependencia 'hacer_llamada_api' no importada correctamente.")

# ---- Helper Interno para Normalizar Destinatarios ----
def _normalize_recipients(rec_input: Optional[Union[str, List[str], List[Dict[str, Any]]]], type_name: str) -> List[Dict[str, Any]]:
    """Normaliza diferentes formatos de entrada de destinatarios a la estructura de Graph API."""
    recipients_list: List[Dict[str, Any]] = []
    if not rec_input:
        return recipients_list # Lista vacía si la entrada es None o vacía

    if isinstance(rec_input, str):
        # Asumir que es una sola dirección de correo
        if rec_input.strip():
            recipients_list.append({"emailAddress": {"address": rec_input.strip()}})
    elif isinstance(rec_input, list):
        for item in rec_input:
            if isinstance(item, str) and item.strip():
                recipients_list.append({"emailAddress": {"address": item.strip()}})
            elif isinstance(item, dict) and isinstance(item.get("emailAddress"), dict) and isinstance(item["emailAddress"].get("address"), str):
                # Ya está en el formato correcto
                recipients_list.append(item)
            else:
                logger.warning(f"Item inválido en lista de {type_name}: {item}. Se ignorará.")
    else:
        raise TypeError(f"Formato inválido para {type_name}: Se esperaba str, List[str] o List[Dict]. Se recibió {type(rec_input)}.")

    return recipients_list

# ---- FUNCIONES DE ACCIÓN PARA CORREO ----
# Todas usan la firma (parametros: Dict[str, Any], headers: Dict[str, str])

def listar_correos(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista correos de una carpeta específica.

    Args:
        parametros (Dict[str, Any]): Opcional: 'mailbox' (default 'me'), 'folder' (default 'Inbox'),
                                     'top' (int, default 10), 'skip' (int, default 0),
                                     'select' (List[str]), 'filter_query', 'order_by'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Respuesta de Graph API, usualmente {'value': [...]}.
    """
    mailbox: str = parametros.get('mailbox', 'me')
    folder: str = parametros.get('folder', 'Inbox') # Default a Inbox
    top: int = int(parametros.get('top', 10))
    skip: int = int(parametros.get('skip', 0))
    select: Optional[List[str]] = parametros.get('select')
    filter_query: Optional[str] = parametros.get('filter_query')
    order_by: Optional[str] = parametros.get('order_by')

    # Construir URL y parámetros de query
    url = f"{BASE_URL}/users/{mailbox}/mailFolders/{folder}/messages"
    params_query: Dict[str, Any] = {'$top': top, '$skip': skip}
    if select: params_query['$select'] = ','.join(select)
    if filter_query: params_query['$filter'] = filter_query
    if order_by: params_query['$orderby'] = order_by

    # Remover parámetros None
    clean_params = {k: v for k, v in params_query.items() if v is not None}

    logger.info(f"Listando correos para '{mailbox}' carpeta '{folder}' (Top: {top}, Skip: {skip})")
    # La paginación real requeriría manejar @odata.nextLink, similar a listar_eventos/listar_elementos_lista.
    # Por ahora, solo obtiene la página solicitada por top/skip.
    return hacer_llamada_api("GET", url, headers, params=clean_params)


def leer_correo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lee un correo específico por su ID.

    Args:
        parametros (Dict[str, Any]): Debe contener 'message_id'.
                                     Opcional: 'mailbox' (default 'me'), 'select' (List[str]).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El objeto del mensaje de Graph API.
    """
    mailbox: str = parametros.get('mailbox', 'me')
    message_id: Optional[str] = parametros.get('message_id')
    select: Optional[List[str]] = parametros.get('select')

    if not message_id: raise ValueError("Parámetro 'message_id' es requerido.")

    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}"
    params_query = {'$select': ','.join(select)} if select else None

    logger.info(f"Leyendo correo '{message_id}' para '{mailbox}'")
    return hacer_llamada_api("GET", url, headers, params=params_query)


def enviar_correo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Envía un correo electrónico.

    Args:
        parametros (Dict[str, Any]): Debe contener 'destinatario' (str, List[str], o List[Dict]),
                                     'asunto' (str), 'mensaje' (str HTML).
                                     Opcional: 'mailbox' (default 'me'), 'cc', 'bcc' (mismo formato que destinatario),
                                     'attachments' (List[Dict] formato Graph), 'save_to_sent' (bool, default True).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de envío.
    """
    mailbox: str = parametros.get('mailbox', 'me')
    destinatario_in = parametros.get('destinatario')
    asunto: Optional[str] = parametros.get('asunto')
    mensaje: Optional[str] = parametros.get('mensaje')
    cc_in = parametros.get('cc')
    bcc_in = parametros.get('bcc')
    attachments: Optional[List[dict]] = parametros.get('attachments')
    save_to_sent: bool = parametros.get('save_to_sent', True)

    if not destinatario_in: raise ValueError("Parámetro 'destinatario' es requerido.")
    if asunto is None : raise ValueError("Parámetro 'asunto' es requerido.") # Asunto vacío es permitido
    if not mensaje: raise ValueError("Parámetro 'mensaje' (cuerpo del correo) es requerido.")

    # Normalizar destinatarios
    try:
        to_recipients = _normalize_recipients(destinatario_in, "destinatario")
        cc_recipients = _normalize_recipients(cc_in, "cc")
        bcc_recipients = _normalize_recipients(bcc_in, "bcc")
    except TypeError as e:
        raise ValueError(f"Error en formato de destinatarios: {e}") from e

    if not to_recipients: raise ValueError("Al menos un destinatario válido es requerido en 'destinatario'.")

    # Construir payload para sendMail
    message_payload: Dict[str, Any] = {
        "subject": asunto,
        "body": {"contentType": "HTML", "content": mensaje}, # Asumir HTML
        "toRecipients": to_recipients
    }
    if cc_recipients: message_payload["ccRecipients"] = cc_recipients
    if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
    if attachments and isinstance(attachments, list): message_payload["attachments"] = attachments

    final_payload = {
        "message": message_payload,
        "saveToSentItems": str(save_to_sent).lower() # API espera string 'true' o 'false'
    }

    url = f"{BASE_URL}/users/{mailbox}/sendMail"
    logger.info(f"Enviando correo para '{mailbox}'. Asunto: '{asunto}'")

    # sendMail devuelve 202 Accepted (sin cuerpo). El helper devuelve None para 2xx sin cuerpo.
    hacer_llamada_api("POST", url, headers, json_data=final_payload)

    # Devolver una confirmación ya que la API no devuelve contenido en éxito
    return {"status": "Correo enviado/encolado exitosamente"}


def guardar_borrador(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Guarda un correo como borrador.

    Args:
        parametros (Dict[str, Any]): Debe contener 'asunto', 'mensaje'.
                                     Opcional: 'mailbox' (default 'me'), 'destinatario', 'cc', 'bcc', 'attachments'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El objeto del mensaje borrador creado por Graph API.
    """
    mailbox: str = parametros.get('mailbox', 'me')
    asunto: Optional[str] = parametros.get('asunto')
    mensaje: Optional[str] = parametros.get('mensaje')
    destinatario_in = parametros.get('destinatario')
    cc_in = parametros.get('cc')
    bcc_in = parametros.get('bcc')
    attachments: Optional[List[dict]] = parametros.get('attachments')

    if asunto is None : raise ValueError("Parámetro 'asunto' es requerido.")
    if not mensaje: raise ValueError("Parámetro 'mensaje' (cuerpo del correo) es requerido.")

    # Normalizar destinatarios (son opcionales para borrador)
    try:
        to_recipients = _normalize_recipients(destinatario_in, "destinatario")
        cc_recipients = _normalize_recipients(cc_in, "cc")
        bcc_recipients = _normalize_recipients(bcc_in, "bcc")
    except TypeError as e:
        raise ValueError(f"Error en formato de destinatarios: {e}") from e

    # Construir payload para crear mensaje (borrador)
    message_payload: Dict[str, Any] = {
        "subject": asunto,
        "body": {"contentType": "HTML", "content": mensaje} # Asumir HTML
    }
    if to_recipients: message_payload["toRecipients"] = to_recipients
    if cc_recipients: message_payload["ccRecipients"] = cc_recipients
    if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
    if attachments and isinstance(attachments, list): message_payload["attachments"] = attachments

    url = f"{BASE_URL}/users/{mailbox}/messages" # POST a /messages crea un borrador
    logger.info(f"Guardando borrador para '{mailbox}'. Asunto: '{asunto}'")

    # POST a /messages devuelve el objeto del mensaje creado (201 Created)
    return hacer_llamada_api("POST", url, headers, json_data=message_payload)


def enviar_borrador(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Envía un correo que previamente fue guardado como borrador.

    Args:
        parametros (Dict[str, Any]): Debe contener 'message_id'. Opcional: 'mailbox' (default 'me').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de envío.
    """
    mailbox: str = parametros.get('mailbox', 'me')
    message_id: Optional[str] = parametros.get('message_id')

    if not message_id: raise ValueError("Parámetro 'message_id' del borrador es requerido.")

    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}/send"
    logger.info(f"Enviando borrador '{message_id}' para '{mailbox}'")

    # POST a /send no requiere body y devuelve 202 Accepted (None del helper).
    hacer_llamada_api("POST", url, headers)

    return {"status": "Borrador enviado exitosamente"}


def responder_correo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Responde a un correo existente (reply o replyAll).

    Args:
        parametros (Dict[str, Any]): Debe contener 'message_id', 'mensaje_respuesta'.
                                     Opcional: 'mailbox' (default 'me'), 'reply_all' (bool, default False),
                                     'to_recipients' (List[Dict], para modificar destinatarios al responder).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de respuesta.
    """
    mailbox: str = parametros.get('mailbox', 'me')
    message_id: Optional[str] = parametros.get('message_id')
    mensaje_respuesta: Optional[str] = parametros.get('mensaje_respuesta')
    reply_all: bool = parametros.get('reply_all', False)
    # Permite opcionalmente sobreescribir/añadir destinatarios
    to_recipients_in = parametros.get('to_recipients')

    if not message_id: raise ValueError("Parámetro 'message_id' es requerido.")
    if not mensaje_respuesta: raise ValueError("Parámetro 'mensaje_respuesta' es requerido.")

    action = "replyAll" if reply_all else "reply"
    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}/{action}"

    # El cuerpo principal va en 'comment'. Opcionalmente se puede modificar el 'message'.
    payload: Dict[str, Any] = {"comment": mensaje_respuesta}

    # Si se proporcionan 'to_recipients', se añaden al objeto 'message'
    if to_recipients_in:
        try:
            normalized_to = _normalize_recipients(to_recipients_in, "to_recipients (respuesta)")
            if normalized_to:
                payload["message"] = {"toRecipients": normalized_to}
        except TypeError as e:
             raise ValueError(f"Error en formato 'to_recipients': {e}") from e

    logger.info(f"Respondiendo{' a todos' if reply_all else ''} correo '{message_id}' para '{mailbox}'")

    # POST a /reply o /replyAll devuelve 202 Accepted (None del helper).
    hacer_llamada_api("POST", url, headers, json_data=payload)

    return {"status": "Respuesta enviada exitosamente"}


def reenviar_correo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Reenvía un correo existente.

    Args:
        parametros (Dict[str, Any]): Debe contener 'message_id', 'destinatarios' (str, List[str], o List[Dict]).
                                     Opcional: 'mailbox' (default 'me'), 'mensaje_reenvio' (str, comentario adicional).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de reenvío.
    """
    mailbox: str = parametros.get('mailbox', 'me')
    message_id: Optional[str] = parametros.get('message_id')
    destinatarios_in = parametros.get('destinatarios')
    mensaje_reenvio: str = parametros.get('mensaje_reenvio', "") # Comentario opcional

    if not message_id: raise ValueError("Parámetro 'message_id' es requerido.")
    if not destinatarios_in: raise ValueError("Parámetro 'destinatarios' es requerido.")

    # Normalizar destinatarios
    try:
        to_recipients = _normalize_recipients(destinatarios_in, "destinatarios (reenvío)")
    except TypeError as e:
        raise ValueError(f"Error en formato de destinatarios: {e}") from e

    if not to_recipients: raise ValueError("Al menos un destinatario válido es requerido en 'destinatarios'.")

    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}/forward"
    payload = {
        "toRecipients": to_recipients,
        "comment": mensaje_reenvio # Comentario que se añade al cuerpo del mensaje reenviado
        # Se podría añadir un 'message' aquí si se quiere modificar más el correo reenviado
    }

    logger.info(f"Reenviando correo '{message_id}' para '{mailbox}' a {len(to_recipients)} destinatario(s)")

    # POST a /forward devuelve 202 Accepted (None del helper).
    hacer_llamada_api("POST", url, headers, json_data=payload)

    return {"status": "Correo reenviado exitosamente"}


def eliminar_correo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina un correo (lo mueve a Elementos Eliminados).

    Args:
        parametros (Dict[str, Any]): Debe contener 'message_id'. Opcional: 'mailbox' (default 'me').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de eliminación.
    """
    mailbox: str = parametros.get('mailbox', 'me')
    message_id: Optional[str] = parametros.get('message_id')

    if not message_id: raise ValueError("Parámetro 'message_id' es requerido.")

    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}"
    logger.info(f"Eliminando correo '{message_id}' para '{mailbox}'")

    # DELETE devuelve 204 No Content (None del helper).
    hacer_llamada_api("DELETE", url, headers)

    return {"status": "Correo eliminado exitosamente"}

# --- FIN DEL MÓDULO actions/correo.py ---

