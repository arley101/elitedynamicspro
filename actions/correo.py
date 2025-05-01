# actions/correo.py (Refactorizado v2 con Helper)

import logging
import requests # Para tipos de excepción y paginación
import json
from typing import Dict, List, Optional, Union, Any

# Usar logger principal
logger = logging.getLogger("azure.functions")

# Importar helper y constantes
try:
    from helpers.http_client import hacer_llamada_api
    from shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    logger.error("Error importando helpers/constantes en Correo.")
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs): raise NotImplementedError("Helper no importado")

# ---- FUNCIONES DE CORREO ----
# Usan el helper hacer_llamada_api

def listar_correos(headers: Dict[str, str], top: int = 10, skip: int = 0, folder: str = 'Inbox', select: Optional[List[str]] = None, filter_query: Optional[str] = None, order_by: Optional[str] = None, mailbox: str = 'me') -> Dict[str, Any]:
    url = f"{BASE_URL}/users/{mailbox}/mailFolders/{folder}/messages"
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if select: params['$select'] = ','.join(select)
    if filter_query: params['$filter'] = filter_query
    if order_by: params['$orderby'] = order_by
    clean_params = {k:v for k, v in params.items() if v is not None}
    logger.info(f"Listando correos para '{mailbox}' carpeta '{folder}'")
    # Paginación podría necesitar llamada directa a requests si se quieren > top
    return hacer_llamada_api("GET", url, headers, params=clean_params)

def leer_correo(headers: Dict[str, str], message_id: str, select: Optional[List[str]] = None, mailbox: str = 'me') -> Dict[str, Any]:
    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}"
    params = {'$select': ','.join(select)} if select else None
    logger.info(f"Leyendo correo '{message_id}' para '{mailbox}'")
    return hacer_llamada_api("GET", url, headers, params=params)

def enviar_correo(headers: Dict[str, str], destinatario: Union[str, List[str]], asunto: str, mensaje: str, cc: Optional[Union[str, List[str]]] = None, bcc: Optional[Union[str, List[str]]] = None, attachments: Optional[List[dict]] = None, save_to_sent: bool = True, mailbox: str = 'me') -> Optional[Dict[str, Any]]:
    url = f"{BASE_URL}/users/{mailbox}/sendMail"
    def normalize_recipients(rec_input: Union[str, List[str]], type_name: str) -> List[Dict[str, Any]]:
        if isinstance(rec_input, str): rec_list = [rec_input]
        elif isinstance(rec_input, list): rec_list = rec_input
        else: raise TypeError(f"{type_name} debe ser str o List[str]")
        return [{"emailAddress": {"address": r}} for r in rec_list if r and isinstance(r, str)]
    try: to_recipients = normalize_recipients(destinatario, "Destinatario"); cc_recipients = normalize_recipients(cc, "CC") if cc else []; bcc_recipients = normalize_recipients(bcc, "BCC") if bcc else []
    except TypeError as e: raise ValueError(f"Formato destinatario inválido: {e}")
    if not to_recipients: raise ValueError("Destinatario válido requerido.")
    message_payload: Dict[str, Any] = {"subject": asunto, "body": {"contentType": "HTML", "content": mensaje}, "toRecipients": to_recipients}
    if cc_recipients: message_payload["ccRecipients"] = cc_recipients
    if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
    if attachments: message_payload["attachments"] = attachments
    final_payload = {"message": message_payload, "saveToSentItems": str(save_to_sent).lower()}
    logger.info(f"Enviando correo para '{mailbox}'. Asunto: '{asunto}'")
    # sendMail devuelve 202 Accepted sin cuerpo, el helper devuelve None en ese caso
    hacer_llamada_api("POST", url, headers, json_data=final_payload)
    return {"status": "Enviado"} # Devolver confirmación

def guardar_borrador(headers: Dict[str, str], destinatario: Union[str, List[str]], asunto: str, mensaje: str, cc: Optional[Union[str, List[str]]] = None, bcc: Optional[Union[str, List[str]]] = None, attachments: Optional[List[dict]] = None, mailbox: str = 'me') -> Dict[str, Any]:
    url = f"{BASE_URL}/users/{mailbox}/messages"
    def normalize_recipients(rec_input: Union[str, List[str]], type_name: str) -> List[Dict[str, Any]]:
        if isinstance(rec_input, str): rec_list = [rec_input]
        elif isinstance(rec_input, list): rec_list = rec_input
        else: raise TypeError(f"{type_name} debe ser str o List[str]")
        return [{"emailAddress": {"address": r}} for r in rec_list if r and isinstance(r, str)]
    try: to_recipients = normalize_recipients(destinatario, "Destinatario"); cc_recipients = normalize_recipients(cc, "CC") if cc else []; bcc_recipients = normalize_recipients(bcc, "BCC") if bcc else []
    except TypeError as e: raise ValueError(f"Formato destinatario inválido: {e}")
    message_payload: Dict[str, Any] = {"subject": asunto, "body": {"contentType": "HTML", "content": mensaje}}
    if to_recipients: message_payload["toRecipients"] = to_recipients
    if cc_recipients: message_payload["ccRecipients"] = cc_recipients
    if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
    if attachments: message_payload["attachments"] = attachments
    logger.info(f"Guardando borrador para '{mailbox}'. Asunto: '{asunto}'")
    return hacer_llamada_api("POST", url, headers, json_data=message_payload) # Devuelve el objeto del mensaje creado

def enviar_borrador(headers: Dict[str, str], message_id: str, mailbox: str = 'me') -> Optional[Dict[str, Any]]:
    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}/send"
    logger.info(f"Enviando borrador '{message_id}' para '{mailbox}'")
    hacer_llamada_api("POST", url, headers) # POST sin body, devuelve 202 (None del helper)
    return {"status": "Borrador Enviado"}

def responder_correo(headers: Dict[str, str], message_id: str, mensaje_respuesta: str, to_recipients: Optional[List[dict]] = None, reply_all: bool = False, mailbox: str = 'me') -> Optional[Dict[str, Any]]:
    action = "replyAll" if reply_all else "reply"; url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}/{action}"
    payload: Dict[str, Any] = {"comment": mensaje_respuesta}
    if to_recipients: payload["message"] = { "toRecipients": to_recipients }
    logger.info(f"Respondiendo{' a todos' if reply_all else ''} correo '{message_id}' para '{mailbox}'")
    hacer_llamada_api("POST", url, headers, json_data=payload) # Devuelve 202 (None del helper)
    return {"status": "Respondido"}

def reenviar_correo(headers: Dict[str, str], message_id: str, destinatarios: Union[str, List[str]], mensaje_reenvio: str = "FYI", mailbox: str = 'me') -> Optional[Dict[str, Any]]:
    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}/forward"
    if isinstance(destinatarios, str): destinatarios_list = [destinatarios]
    elif isinstance(destinatarios, list): destinatarios_list = destinatarios
    else: raise TypeError("Destinatarios debe ser str o List[str]")
    to_recipients_list = [{"emailAddress": {"address": r}} for r in destinatarios_list if r and isinstance(r, str)]
    if not to_recipients_list: raise ValueError("Destinatario válido requerido.")
    payload = {"toRecipients": to_recipients_list, "comment": mensaje_reenvio}
    logger.info(f"Reenviando correo '{message_id}' para '{mailbox}'")
    hacer_llamada_api("POST", url, headers, json_data=payload) # Devuelve 202 (None del helper)
    return {"status": "Reenviado"}

def eliminar_correo(headers: Dict[str, str], message_id: str, mailbox: str = 'me') -> Optional[Dict[str, Any]]:
    url = f"{BASE_URL}/users/{mailbox}/messages/{message_id}"
    logger.info(f"Eliminando correo '{message_id}' para '{mailbox}'")
    hacer_llamada_api("DELETE", url, headers) # Devuelve 204 (None del helper)
    return {"status": "Eliminado"}
