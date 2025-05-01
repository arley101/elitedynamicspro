import logging
import requests
import json # Para manejo de errores
from typing import Dict, List, Optional, Union, Any

# Usar el logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar constantes globales desde __init__.py
try:
    from .. import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    # Fallback
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_TIMEOUT = 45
    logger.warning("No se pudo importar BASE_URL/GRAPH_API_TIMEOUT desde el padre, usando defaults.")


# ---- FUNCIONES DE CORREO (Refactorizadas) ----
# Aceptan 'headers', usan mailbox='me' por defecto

def listar_correos(
    headers: Dict[str, str],
    top: int = 10,
    skip: int = 0,
    folder: str = 'Inbox',
    select: Optional[List[str]] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None,
    mailbox: str = 'me' # Usa 'me' por defecto con token delegado
) -> Dict[str, Any]:
    """Lista correos de una carpeta. Requiere headers autenticados."""
    url = f"{BASE_URL}/users/{mailbox}/mailFolders/{folder}/messages"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if select: params['$select'] = ','.join(select)
    if filter_query: params['$filter'] = filter_query
    if order_by: params['$orderby'] = order_by

    clean_params = {k:v for k, v in params.items() if v is not None}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} Params: {clean_params} (Listando correos para '{mailbox}')")
        response = requests.get(url, headers=headers, params=clean_params, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados {len(data.get('value',[]))} correos para '{mailbox}'.")
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en listar_correos: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_correos: {e}", exc_info=True)
        raise

def leer_correo(headers: Dict[str, str], message_id: str, select: Optional[List[str]] = None, mailbox: str = 'me') -> Dict[str, Any]:
    """Lee un correo específico. Requiere headers autenticados."""
    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}"
    params = {}
    if select: params['$select'] = ','.join(select)
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} Params: {params} (Leyendo correo '{message_id}' para '{mailbox}')")
        response = requests.get(url, headers=headers, params=params or None, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Correo '{message_id}' leído para '{mailbox}'.")
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en leer_correo {message_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en leer_correo {message_id}: {e}", exc_info=True)
        raise

def enviar_correo(
    headers: Dict[str, str],
    destinatario: Union[str, List[str]],
    asunto: str,
    mensaje: str,
    cc: Optional[Union[str, List[str]]] = None,
    bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[dict]] = None, # Formato Graph API esperado
    # from_email: Optional[str] = None, # 'From' usualmente no permitido con Graph SendMail básico
    save_to_sent: bool = True, # Parámetro para controlar si se guarda
    mailbox: str = 'me'
) -> Dict[str, Any]:
    """Envía un correo electrónico. Requiere headers autenticados."""
    # NOTA: Esta función envía directamente, no crea borrador. Usar guardar_borrador para eso.
    url = f"{BASE_URL}/users/{mailbox}/sendMail"
    log_action = "Enviando correo"

    def normalize_recipients(rec_input: Union[str, List[str]], type_name: str) -> List[Dict[str, Any]]:
        if isinstance(rec_input, str): rec_list = [rec_input]
        elif isinstance(rec_input, list): rec_list = rec_input
        else: raise TypeError(f"{type_name} debe ser str o List[str]")
        return [{"emailAddress": {"address": r}} for r in rec_list if r and isinstance(r, str)]

    try:
        to_recipients = normalize_recipients(destinatario, "Destinatario")
        cc_recipients = normalize_recipients(cc, "CC") if cc else []
        bcc_recipients = normalize_recipients(bcc, "BCC") if bcc else []
    except TypeError as e:
        logger.error(f"Error en formato de destinatarios: {e}")
        raise ValueError(f"Formato de destinatario inválido: {e}")

    if not to_recipients:
        raise ValueError("Se requiere al menos un destinatario válido.")

    message_payload: Dict[str, Any] = {
        "subject": asunto,
        "body": {"contentType": "HTML", "content": mensaje}, # Asume HTML
        "toRecipients": to_recipients,
    }
    if cc_recipients: message_payload["ccRecipients"] = cc_recipients
    if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
    if attachments: message_payload["attachments"] = attachments
    # if from_email: message_payload["from"] = {"emailAddress": {"address": from_email}} # Necesita permisos especiales

    final_payload = {"message": message_payload, "saveToSentItems": str(save_to_sent).lower()}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} ({log_action} para '{mailbox}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=final_payload, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 202 Accepted
        logger.info(f"Correo enviado para '{mailbox}'. Asunto: '{asunto}'")
        return {"status": "Enviado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en {log_action}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en {log_action}: {e}", exc_info=True)
        raise

def guardar_borrador(
    headers: Dict[str, str],
    destinatario: Union[str, List[str]],
    asunto: str,
    mensaje: str,
    cc: Optional[Union[str, List[str]]] = None,
    bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[dict]] = None,
    mailbox: str = 'me'
) -> Dict[str, Any]:
    """Guarda un correo como borrador. Requiere headers autenticados."""
    url = f"{BASE_URL}/users/{mailbox}/messages" # POST a /messages
    log_action = "Guardando borrador"

    # Reutilizar normalización de destinatarios
    def normalize_recipients(rec_input: Union[str, List[str]], type_name: str) -> List[Dict[str, Any]]:
        if isinstance(rec_input, str): rec_list = [rec_input]
        elif isinstance(rec_input, list): rec_list = rec_input
        else: raise TypeError(f"{type_name} debe ser str o List[str]")
        return [{"emailAddress": {"address": r}} for r in rec_list if r and isinstance(r, str)]
    try:
        to_recipients = normalize_recipients(destinatario, "Destinatario")
        cc_recipients = normalize_recipients(cc, "CC") if cc else []
        bcc_recipients = normalize_recipients(bcc, "BCC") if bcc else []
    except TypeError as e: raise ValueError(f"Formato destinatario inválido: {e}")
    # Permitir guardar borrador sin destinatario si es necesario
    # if not to_recipients: raise ValueError("Destinatario válido requerido.")

    # Construir el objeto message directamente
    message_payload: Dict[str, Any] = {
        "subject": asunto,
        "body": {"contentType": "HTML", "content": mensaje},
    }
    # Añadir solo si hay destinatarios válidos
    if to_recipients: message_payload["toRecipients"] = to_recipients
    if cc_recipients: message_payload["ccRecipients"] = cc_recipients
    if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
    if attachments: message_payload["attachments"] = attachments

    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} ({log_action} para '{mailbox}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=message_payload, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 201 Created
        data = response.json()
        message_id = data.get('id')
        logger.info(f"Borrador guardado para '{mailbox}'. ID: {message_id}.")
        # Devolver el objeto completo del borrador creado
        return data
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en {log_action}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en {log_action}: {e}", exc_info=True)
        raise

def enviar_borrador(headers: Dict[str, str], message_id: str, mailbox: str = 'me') -> Dict[str, Any]:
    """Envía un borrador guardado. Requiere headers autenticados."""
    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}/send"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Enviando borrador '{message_id}' para '{mailbox}')")
        # Send es POST sin cuerpo
        response = requests.post(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 202 Accepted
        logger.info(f"Borrador '{message_id}' enviado para '{mailbox}'.")
        return {"status": "Borrador Enviado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en enviar_borrador {message_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en enviar_borrador {message_id}: {e}", exc_info=True)
        raise

def responder_correo(headers: Dict[str, str], message_id: str, mensaje_respuesta: str, to_recipients: Optional[List[dict]] = None, reply_all: bool = False, mailbox: str = 'me') -> dict:
    """Responde a un correo. Requiere headers autenticados."""
    action = "replyAll" if reply_all else "reply"
    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}/{action}"
    payload: Dict[str, Any] = {"comment": mensaje_respuesta}
    if to_recipients: payload["message"] = { "toRecipients": to_recipients }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Respondiendo{' a todos' if reply_all else ''} correo '{message_id}' para '{mailbox}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=payload, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 202 Accepted
        logger.info(f"Respuesta {'a todos ' if reply_all else ''}enviada correo '{message_id}'.")
        return {"status": "Respondido", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en responder_correo {message_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en responder_correo {message_id}: {e}", exc_info=True)
        raise

def reenviar_correo(headers: Dict[str, str], message_id: str, destinatarios: Union[str, List[str]], mensaje_reenvio: str = "FYI", mailbox: str = 'me') -> dict:
    """Reenvía un correo. Requiere headers autenticados."""
    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}/forward"

    if isinstance(destinatarios, str): destinatarios_list = [destinatarios]
    elif isinstance(destinatarios, list): destinatarios_list = destinatarios
    else: raise TypeError("Destinatarios debe ser str o List[str]")
    to_recipients_list = [{"emailAddress": {"address": r}} for r in destinatarios_list if r and isinstance(r, str)]
    if not to_recipients_list: raise ValueError("Destinatario válido requerido.")

    payload = {"toRecipients": to_recipients_list, "comment": mensaje_reenvio}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Reenviando correo '{message_id}' para '{mailbox}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=payload, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 202 Accepted
        logger.info(f"Correo '{message_id}' reenviado para '{mailbox}'.")
        return {"status": "Reenviado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en reenviar_correo {message_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en reenviar_correo {message_id}: {e}", exc_info=True)
        raise

def eliminar_correo(headers: Dict[str, str], message_id: str, mailbox: str = 'me') -> dict:
    """Elimina un correo (mueve a Elementos Eliminados). Requiere headers autenticados."""
    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando correo '{message_id}' para '{mailbox}')")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 204 No Content
        logger.info(f"Correo '{message_id}' eliminado para '{mailbox}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request en eliminar_correo {message_id}: {req_ex}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_correo {message_id}: {e}", exc_info=True)
        raise
