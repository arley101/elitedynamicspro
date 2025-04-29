import logging
import os
import requests
# Asegúrate que auth.py esté en la raíz o ajusta la importación
from auth import obtener_token
from typing import Dict, List, Optional, Union, Any # Añadido Any
import json # Añadido para los except blocks

# Configuración básica de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- INICIO: Configuración Redundante ---
# Considera eliminar esta sección si ya está centralizada
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
GRAPH_SCOPE = os.getenv('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')
if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE]):
    logging.error("❌ Faltan variables de entorno en correo.py.")
    # No lanzar excepción aquí para permitir importación
BASE_URL = "https://graph.microsoft.com/v1.0"
HEADERS: Dict[str, Optional[str]] = { # Tipado explícito
    'Authorization': None,
    'Content-Type': 'application/json'
}
MAILBOX = os.getenv('MAILBOX', 'me')

def _actualizar_headers() -> None:
    """Obtiene un nuevo token de acceso y actualiza el diccionario HEADERS local."""
    try:
        # Llama a la función importada de auth.py
        token = obtener_token() # Asume flujo 'aplicacion' por defecto
        HEADERS['Authorization'] = f'Bearer {token}'
        logging.info("Headers actualizados en correo.py.")
    except Exception as e:
        logging.error(f"❌ Error al obtener el token en correo.py: {e}")
        raise Exception(f"Error al obtener el token en correo.py: {e}")
# --- FIN: Configuración Redundante ---


# ---- FUNCIONES DE GESTIÓN DE CORREO DE OUTLOOK ----
def listar_correos(
    top: int = 10,
    skip: int = 0,
    folder: str = 'Inbox',
    select: Optional[List[str]] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None
    # Quitamos mailbox si usamos el global y _actualizar_headers local
) -> Dict[str, Any]: # Mejor tipo retorno
    """
    Lista correos electrónicos de una carpeta de Outlook... (Docstring como estaba)
    """
    _actualizar_headers()
    # Usar MAILBOX global definido en este archivo
    usuario = MAILBOX
    url = f"{BASE_URL}/users/{usuario}/mailFolders/{folder}/messages" # Ajustado para usar /users/ en vez de /me/ si MAILBOX no es 'me'
    params: Dict[str, Any] = {'$top': int(top), '$skip': int(skip)}
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    if filter_query is not None and isinstance(filter_query, str): params['$filter'] = filter_query
    if order_by is not None and isinstance(order_by, str): params['$orderby'] = order_by

    response: Optional[requests.Response] = None
    try:
        clean_params = {k:v for k, v in params.items() if v is not None}
        logging.info(f"Llamando a Graph API (correo.py): GET {url} con params: {clean_params}")
        response = requests.get(url, headers=HEADERS, params=clean_params)
        response.raise_for_status()
        data: Dict[str, Any] = response.json()
        logging.info(f"Listados {len(data.get('value',[]))} correos desde correo.py.")
        return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error listar correos (correo.py): {e}. URL: {url}. Detalles: {error_details}"); raise Exception(f"Error al listar correos (correo.py): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (listar correos - correo.py): {e}. Respuesta: {response_text}"); raise Exception(f"Error al decodificar JSON (listar correos): {e}")


def leer_correo(message_id: str, select: Optional[List[str]] = None) -> Dict[str, Any]: # Mejor tipo retorno
    """Lee un correo electrónico específico de Outlook."""
    _actualizar_headers()
    usuario = MAILBOX
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}" # Ajustado
    params = {}
    response: Optional[requests.Response] = None
    if select and isinstance(select, list): params['$select'] = ','.join(select)
    try:
        logging.info(f"Llamando a Graph API (correo.py): GET {url} con params: {params}")
        response = requests.get(url, headers=HEADERS, params=params or None)
        response.raise_for_status()
        data: Dict[str, Any] = response.json(); logging.info(f"Correo '{message_id}' leído (correo.py)."); return data
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error leer correo (correo.py): {e}. URL: {url}. Detalles: {error_details}"); raise Exception(f"Error leer correo (correo.py): {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (leer correo - correo.py): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (leer correo): {e}")


# !!!!! INICIO: VERSIÓN CORREGIDA DE enviar_correo !!!!!
def enviar_correo(
    destinatario: Union[str, List[str]],
    asunto: str,
    mensaje: str,
    cc: Optional[Union[str, List[str]]] = None,
    bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[dict]] = None,
    from_email: Optional[str] = None, # Requiere permisos especiales
    is_draft: bool = False
) -> Dict[str, Any]: # Mejor tipo retorno
    """Envía un correo electrónico desde Outlook o guarda un borrador."""
    _actualizar_headers()
    usuario = MAILBOX # Usar la variable global de este módulo

    if is_draft:
        url = f"{BASE_URL}/users/{usuario}/messages" # POST a /messages para crear borrador
    else:
        url = f"{BASE_URL}/users/{usuario}/sendMail" # POST a /sendMail para enviar directo

    # --- CORRECCIÓN: Filtrar Nones/vacíos/no-strings en destinatarios ---
    if isinstance(destinatario, str): destinatario_list = [destinatario]
    elif isinstance(destinatario, list): destinatario_list = destinatario
    else: raise TypeError("Destinatario debe ser str o List[str]")
    # Filtra None, strings vacíos, y asegura que sean strings
    to_recipients = [{"emailAddress": {"address": r}} for r in destinatario_list if r and isinstance(r, str)]

    cc_recipients = []
    if cc:
        if isinstance(cc, str): cc_list = [cc]
        elif isinstance(cc, list): cc_list = cc
        else: raise TypeError("CC debe ser str o List[str]")
        # Filtra None, strings vacíos, y asegura que sean strings
        cc_recipients = [{"emailAddress": {"address": r}} for r in cc_list if r and isinstance(r, str)]

    bcc_recipients = []
    if bcc:
        if isinstance(bcc, str): bcc_list = [bcc]
        elif isinstance(bcc, list): bcc_list = bcc
        else: raise TypeError("BCC debe ser str o List[str]")
        # Filtra None, strings vacíos, y asegura que sean strings
        bcc_recipients = [{"emailAddress": {"address": r}} for r in bcc_list if r and isinstance(r, str)]
    # --- FIN CORRECCIÓN ---

    if not to_recipients:
        logging.error("No se proporcionaron destinatarios válidos para enviar_correo.")
        raise ValueError("Se requiere al menos un destinatario válido.")

    # Construir el objeto 'message'
    message_payload: Dict[str, Any] = {
        "subject": asunto,
        "body": {"contentType": "HTML", "content": mensaje},
        "toRecipients": to_recipients,
    }
    if cc_recipients: message_payload["ccRecipients"] = cc_recipients
    if bcc_recipients: message_payload["bccRecipients"] = bcc_recipients
    if attachments: message_payload["attachments"] = attachments
    if from_email: message_payload["from"] = {"emailAddress": {"address": from_email}}

    # El payload final depende si es borrador o envío directo
    if is_draft:
        final_payload = message_payload
    else:
        final_payload = {"message": message_payload, "saveToSentItems": "true"}

    response: Optional[requests.Response] = None
    try:
        logging.info(f"Llamando a Graph API (correo.py): POST {url}")
        response = requests.post(url, headers=HEADERS, json=final_payload)
        response.raise_for_status()

        if not is_draft:
            logging.info(f"Correo enviado con asunto '{asunto}'.")
            return {"status": "Enviado", "code": response.status_code}
        else:
            data = response.json()
            message_id = data.get('id')
            logging.info(f"Correo guardado como borrador ID: {message_id}.")
            return {"status": "Borrador Guardado", "code": response.status_code, "id": message_id, "data": data}

    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logging.error(f"❌ Error al {'enviar' if not is_draft else 'guardar borrador'} correo (correo.py): {e}. Detalles: {error_details}. URL: {url}")
        raise Exception(f"Error al {'enviar' if not is_draft else 'guardar borrador'} correo: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logging.error(f"❌ Error al decodificar JSON (guardar borrador - correo.py): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (guardar borrador): {e}")
# !!!!! FIN: VERSIÓN CORREGIDA DE enviar_correo !!!!!


def guardar_borrador(
    destinatario: Union[str, List[str]],
    asunto: str,
    mensaje: str,
    cc: Optional[Union[str, List[str]]] = None,
    bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[dict]] = None,
    from_email: Optional[str] = None
) -> dict:
    """Guarda un correo electrónico como borrador en Outlook."""
    # Llama a la función enviar_correo (corregida) de este mismo módulo
    return enviar_correo(destinatario, asunto, mensaje, cc, bcc, attachments, from_email, is_draft=True)


def enviar_borrador(message_id: str) -> dict:
    """Envía un correo electrónico que ha sido guardado como borrador en Outlook."""
    _actualizar_headers()
    usuario = MAILBOX
    # CORRECCION: Usar la URL correcta con el usuario/mailbox
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/send"
    response: Optional[requests.Response] = None
    try:
        response = requests.post(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Borrador de correo '{message_id}' enviado.")
        return {"status": "Borrador Enviado", "code": response.status_code}
    except requests.exceptions.RequestException as e: logging.error(f"❌ Error enviar borrador '{message_id}': {e}"); raise Exception(f"Error enviar borrador '{message_id}': {e}")


def responder_correo(message_id: str, mensaje_respuesta: str) -> dict:
    """Responde a un correo electrónico de Outlook."""
    _actualizar_headers()
    usuario = MAILBOX
    # CORRECCION: Usar la URL correcta con el usuario/mailbox
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/reply"
    payload = {"comment": mensaje_respuesta}
    response: Optional[requests.Response] = None
    try:
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status()
        logging.info(f"Respondido al correo '{message_id}'.")
        return {"status": "Respondido", "code": response.status_code}
    except requests.exceptions.RequestException as e: logging.error(f"❌ Error responder correo '{message_id}': {e}"); raise Exception(f"Error responder correo '{message_id}': {e}")


def reenviar_correo(message_id: str, destinatarios: List[str], mensaje_reenvio: str = "Reenviado desde Elite Dynamics Pro") -> dict:
    """Reenvía un correo electrónico de Outlook."""
    _actualizar_headers()
    usuario = MAILBOX
    # CORRECCION: Usar la URL correcta con el usuario/mailbox
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}/forward"
    # CORRECCION: Filtrar destinatarios inválidos
    to_recipients = [{"emailAddress": {"address": r}} for r in destinatarios if r and isinstance(r, str)]
    if not to_recipients: raise ValueError("Se requiere al menos un destinatario válido para reenviar.")
    payload = {"toRecipients": to_recipients, "comment": mensaje_reenvio}
    response: Optional[requests.Response] = None
    try:
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status()
        logging.info(f"Reenviado correo '{message_id}' a: {destinatarios}.")
        return {"status": "Reenviado", "code": response.status_code}
    except requests.exceptions.RequestException as e: logging.error(f"❌ Error reenviar correo '{message_id}': {e}"); raise Exception(f"Error reenviar correo '{message_id}': {e}")


def eliminar_correo(message_id: str) -> dict:
    """Elimina un correo electrónico de Outlook."""
    _actualizar_headers()
    usuario = MAILBOX
    # CORRECCION: Usar la URL correcta con el usuario/mailbox
    url = f"{BASE_URL}/users/{usuario}/messages/{message_id}"
    response: Optional[requests.Response] = None
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Correo '{message_id}' eliminado.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e: logging.error(f"❌ Error eliminar correo '{message_id}': {e}"); raise Exception(f"Error eliminar correo '{message_id}': {e}")
