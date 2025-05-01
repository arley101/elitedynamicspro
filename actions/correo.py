# actions/correo.py (Refactorizado)

import logging
import requests
# Eliminado import auth
from typing import Dict, List, Optional, Union, Any
import json

# Usar logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar constantes globales desde __init__.py
try:
    from .. import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    # Fallback por si se ejecuta standalone
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_TIMEOUT = 45
    logger.warning("No se pudo importar BASE_URL/GRAPH_API_TIMEOUT desde el padre, usando defaults.")

# (Eliminada configuración redundante: CLIENT_ID, SECRET, SCOPE, HEADERS locales, MAILBOX, _actualizar_headers)

# ---- FUNCIONES DE GESTIÓN DE CORREO DE OUTLOOK ----
# Usan /me implícito en el token delegado pasado en 'headers'

def listar_correos(
    headers: Dict[str, str],
    top: int = 10,
    skip: int = 0,
    folder: str = 'Inbox',
    select: Optional[List[str]] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None
    # Quitamos mailbox, siempre será /me con token delegado
) -> Dict[str, Any]:
    """Lista correos electrónicos de una carpeta de Outlook (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    url = f"{BASE_URL}/me/mailFolders/{folder}/messages"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if select: params['$select'] = ','.join(select)
    if filter_query: params['$filter'] = filter_query
    if order_by: params['$orderby'] = order_by

    response: Optional[requests.Response] = None
    try:
        clean_params = {k:v for k, v in params.items() if v is not None}
        logger.info(f"API Call: GET {url} con params: {clean_params}")
        response = requests.get(url, headers=headers, params=clean_params, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data: Dict[str, Any] = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} correos de /me/{folder}.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en listar_correos: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_correos: {e}", exc_info=True)
        raise

def leer_correo(
    headers: Dict[str, str],
    message_id: str,
    select: Optional[List[str]] = None
) -> Dict[str, Any]:
    """Lee un correo electrónico específico de Outlook (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    url = f"{BASE_URL}/me/messages/{message_id}"
    params = {}
    if select: params['$select'] = ','.join(select)
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} con params: {params}")
        response = requests.get(url, headers=headers, params=params or None, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data: Dict[str, Any] = response.json(); logger.info(f"Correo '{message_id}' leído (/me)."); return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en leer_correo: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en leer_correo: {e}", exc_info=True)
        raise

def enviar_correo(
    headers: Dict[str, str],
    destinatario: Union[str, List[str]],
    asunto: str,
    mensaje: str,
    cc: Optional[Union[str, List[str]]] = None,
    bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[dict]] = None,
    # from_email: Optional[str] = None, # Requiere permisos SendAs/SendOnBehalf
    is_draft: bool = False
) -> Dict[str, Any]:
    """Envía un correo electrónico (/me) o guarda un borrador."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    if is_draft: url = f"{BASE_URL}/me/messages"; log_action = "Guardando borrador"
    else: url = f"{BASE_URL}/me/sendMail"; log_action = "Enviando correo"

    def normalize_recipients(rec_input: Union[str, List[str]], type_name: str) -> List[Dict[str, Any]]:
        if isinstance(rec_input, str): rec_list = [rec_input]
        elif isinstance(rec_input, list): rec_list = rec_input
        else: raise TypeError(f"{type_name} debe ser str o List[str]")
        return [{"emailAddress": {"address": r}} for r in rec_list if r and isinstance(r, str)]

    try: to_recipients = normalize_recipients(destinatario, "Destinatario"); cc_recipients = normalize_recipients(cc, "CC") if cc else []; bcc_recipients = normalize_recipients(bcc, "BCC") if bcc else []
    except TypeError as e: logger.error(f"Error formato destinatarios: {e}"); raise ValueError(f"Formato destinatario inválido: {e}")
    if not to_recipients: raise ValueError("Destinatario válido requerido.")

    message_payload: Dict[str, Any] = {
        "subject": asunto, "body": {"contentType": "HTML", "content": mensaje}, "toRecipients": to_recipients,
    }
    if cc_recipients: message_payload["ccRecipients"] = cc_recipients
    if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
    if attachments: message_payload["attachments"] = attachments
    # if from_email: message_payload["from"] = {"emailAddress": {"address": from_email}} # Quitado por ahora

    final_payload = {"message": message_payload, "saveToSentItems": "true"} if not is_draft else message_payload
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} ({log_action} /me)")
        # Usar headers recibidos, Content-Type ya está bien para JSON
        response = requests.post(url, headers=headers, json=final_payload, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        if not is_draft: logger.info(f"Correo enviado asunto '{asunto}'."); return {"status": "Enviado", "code": response.status_code}
        else: data = response.json(); message_id = data.get('id'); logger.info(f"Borrador guardado ID: {message_id}."); return {"status": "Borrador Guardado", "code": response.status_code, "id": message_id, "data": data}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en enviar_correo ({log_action}): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en enviar_correo ({log_action}): {e}", exc_info=True)
        raise

def guardar_borrador(
    headers: Dict[str, str],
    destinatario: Union[str, List[str]], asunto: str, mensaje: str,
    cc: Optional[Union[str, List[str]]] = None, bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[dict]] = None # , from_email: Optional[str] = None
) -> dict:
    """Guarda un correo como borrador (/me)."""
    logger.info(f"Llamando a guardar_borrador (/me). Asunto: '{asunto}'")
    return enviar_correo(headers, destinatario, asunto, mensaje, cc, bcc, attachments, # from_email,
                         is_draft=True)

def enviar_borrador(headers: Dict[str, str], message_id: str) -> dict:
    """Envía un borrador previamente guardado (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    url = f"{BASE_URL}/me/messages/{message_id}/send"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Enviando borrador /me {message_id})")
        response = requests.post(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); logger.info(f"Borrador '{message_id}' enviado (/me)."); return {"status": "Borrador Enviado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en enviar_borrador: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en enviar_borrador: {e}", exc_info=True)
        raise

def responder_correo(
    headers: Dict[str, str],
    message_id: str,
    mensaje_respuesta: str,
    to_recipients: Optional[List[dict]] = None, # Para sobreescribir destinatarios
    reply_all: bool = False
) -> dict:
    """Responde a un correo (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    action = "replyAll" if reply_all else "reply"
    url = f"{BASE_URL}/me/messages/{message_id}/{action}"
    payload: Dict[str, Any] = {"comment": mensaje_respuesta}
    if to_recipients: payload["message"] = { "toRecipients": to_recipients }; logger.info(f"Respondiendo a {message_id} con destinatarios custom.")
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Respondiendo{' a todos' if reply_all else ''} correo /me {message_id})")
        response = requests.post(url, headers=headers, json=payload, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); logger.info(f"Respuesta enviada correo '{message_id}' (/me)."); return {"status": "Respondido", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en responder_correo: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en responder_correo: {e}", exc_info=True)
        raise

def reenviar_correo(
    headers: Dict[str, str],
    message_id: str,
    destinatarios: Union[str, List[str]],
    mensaje_reenvio: str = "FYI"
) -> dict:
    """Reenvía un correo (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    url = f"{BASE_URL}/me/messages/{message_id}/forward"
    if isinstance(destinatarios, str): destinatarios = [destinatarios]
    to_recipients_list = [{"emailAddress": {"address": r}} for r in destinatarios if r and isinstance(r, str)]
    if not to_recipients_list: raise ValueError("Destinatario válido requerido para reenviar.")
    payload = {"toRecipients": to_recipients_list, "comment": mensaje_reenvio}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Reenviando correo /me {message_id})")
        response = requests.post(url, headers=headers, json=payload, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); logger.info(f"Correo '{message_id}' reenviado (/me)."); return {"status": "Reenviado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en reenviar_correo: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en reenviar_correo: {e}", exc_info=True)
        raise

def eliminar_correo(headers: Dict[str, str], message_id: str) -> dict:
    """Elimina un correo (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    url = f"{BASE_URL}/me/messages/{message_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando correo /me {message_id})")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); logger.info(f"Correo '{message_id}' eliminado (/me)."); return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en eliminar_correo: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_correo: {e}", exc_info=True)
        raise
