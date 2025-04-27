import logging
import os
import requests
from auth import obtener_token  # Importante: Importar la función obtener_token
from typing import Dict, List, Optional, Union

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
MAILBOX = os.getenv('MAILBOX', 'me')  # Para permitir acceso a otras bandejas

# Función para obtener el token y actualizar los HEADERS
def _actualizar_headers() -> None:
    """Obtiene un nuevo token de acceso y actualiza el diccionario HEADERS."""
    try:
        HEADERS['Authorization'] = f'Bearer {obtener_token()}'
    except Exception as e:  # Captura la excepción de obtener_token
        logging.error(f"❌ Error al obtener el token: {e}")
        raise Exception(f"Error al obtener el token: {e}")



# ---- FUNCIONES DE GESTIÓN DE CORREO DE OUTLOOK ----
def listar_correos(
    top: int = 10,
    skip: int = 0,
    folder: str = 'Inbox',
    select: Optional[List[str]] = None,
    filter_query: Optional[str] = None,
    order_by: Optional[str] = None
) -> dict:
    """
    Lista correos electrónicos de una carpeta de Outlook, con soporte para paginación, selección de campos, filtrado y ordenamiento.

    Args:
        top: El número máximo de correos electrónicos a devolver.
        skip: El número de correos electrónicos a omitir.
        folder: La carpeta de correo de la que se van a obtener los correos electrónicos (por ejemplo, 'Inbox', 'SentItems').
        select: Una lista de campos a seleccionar (por ejemplo, ['id', 'subject', 'from']). Si no se especifica, se devuelven todos los campos.
        filter_query: Una cadena de consulta para filtrar los correos electrónicos (por ejemplo, "receivedDateTime ge 2024-01-01").
        order_by: Una cadena para especificar el orden de los resultados (por ejemplo, "receivedDateTime desc").

    Returns:
        Un diccionario con la respuesta de la API de Microsoft Graph.
    """
    _actualizar_headers()
    url = f"{BASE_URL}/{MAILBOX}/mailFolders/{folder}/messages?$top={top}&$skip={skip}"

    if select:
        url += f"&$select={','.join(select)}"
    if filter_query:
        url += f"&$filter={filter_query}"
    if order_by:
        url += f"&$orderby={order_by}"

    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Listados {top} correos de la carpeta '{folder}', omitiendo los primeros {skip}.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar correos de la carpeta '{folder}': {e}")
        raise Exception(f"Error al listar correos de la carpeta '{folder}': {e}")



def leer_correo(message_id: str, select: Optional[List[str]] = None) -> dict:
    """Lee un correo electrónico específico de Outlook."""
    _actualizar_headers()
    url = f"{BASE_URL}/{MAILBOX}/messages/{message_id}"
    if select:
        url += f"?$select={','.join(select)}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Correo '{message_id}' leído.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al leer el correo '{message_id}': {e}")
        raise Exception(f"Error al leer el correo '{message_id}': {e}")



def enviar_correo(
    destinatario: Union[str, List[str]],
    asunto: str,
    mensaje: str,
    cc: Optional[Union[str, List[str]]] = None,
    bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[dict]] = None,
    from_email: Optional[str] = None,
    is_draft: bool = False
) -> dict:
    """Envía un correo electrónico desde Outlook o guarda un borrador."""
    _actualizar_headers()
    url = f"https://graph.microsoft.com/v1.0/{MAILBOX}/{'sendMail' if not is_draft else ''}" #cambio de url

    # Manejar destinatarios como lista o cadena
    to_recipients = [{"emailAddress": {"address": recipient}} for recipient in (destinatario if isinstance(destinatario, list) else [destinatario])]
    cc_recipients = [{"emailAddress": {"address": recipient}} for recipient in (cc if isinstance(cc, list) else [cc]) if cc]
    bcc_recipients = [{"emailAddress": {"address": recipient}} for recipient in (bcc if isinstance(bcc, list) else [bcc]) if bcc]

    payload = {
        "message": {
            "subject": asunto,
            "body": {
                "contentType": "Text",  # O "HTML"
                "content": mensaje
            },
            "toRecipients": to_recipients,
        }
    }
    if cc_recipients:
        payload["message"]["ccRecipients"] = cc_recipients
    if bcc_recipients:
        payload["message"]["bccRecipients"] = bcc_recipients
    if attachments:
        payload["message"]["attachments"] = attachments
    if from_email:
        payload["message"]["from"] = {"emailAddress": {"address": from_email}}

    try:
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status()
        if not is_draft:
            logging.info(f"Correo enviado con asunto '{asunto}' a: {destinatario}.")
            return {"status": "Enviado", "code": response.status_code}
        else:
            message_id = response.json().get('id')
            logging.info(f"Correo guardado como borrador con ID: {message_id} y asunto: '{asunto}'")
            return {"status": "Borrador Guardado", "code": response.status_code, "id": message_id}

    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al enviar/guardar correo: {e}")
        raise Exception(f"Error al enviar/guardar correo: {e}")



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
    return enviar_correo(destinatario, asunto, mensaje, cc, bcc, attachments, from_email, is_draft=True)



def enviar_borrador(message_id: str) -> dict:
    """Envía un correo electrónico que ha sido guardado como borrador en Outlook."""
    _actualizar_headers()
    url = f"{BASE_URL}/{message_id}/send"
    try:
        response = requests.post(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Borrador de correo '{message_id}' enviado.")
        return {"status": "Borrador Enviado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al enviar borrador '{message_id}': {e}")
        raise Exception(f"Error al enviar borrador '{message_id}': {e}")



def responder_correo(message_id: str, mensaje_respuesta: str) -> dict:
    """Responde a un correo electrónico de Outlook."""
    _actualizar_headers()
    url = f"{BASE_URL}/{message_id}/reply"
    payload = {"comment": mensaje_respuesta}
    try:
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status()
        logging.info(f"Respondido al correo '{message_id}'.")
        return {"status": "Respondido", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al responder al correo '{message_id}': {e}")
        raise Exception(f"Error al responder al correo '{message_id}': {e}")



def reenviar_correo(message_id: str, destinatarios: List[str], mensaje_reenvio: str = "Reenviado desde Elite Dynamics Pro") -> dict:
    """Reenvía un correo electrónico de Outlook."""
    _actualizar_headers()
    url = f"{BASE_URL}/{message_id}/forward"
    to_recipients = [{"emailAddress": {"address": recipient}} for recipient in destinatarios]
    payload = {
        "toRecipients": to_recipients,
        "comment": mensaje_reenvio
    }
    try:
        response = requests.post(url, headers=HEADERS, json=payload)
        response.raise_for_status()
        logging.info(f"Reenviado el correo '{message_id}' a: {destinatarios}.")
        return {"status": "Reenviado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al reenviar el correo '{message_id}': {e}")
        raise Exception(f"Error al reenviar el correo '{message_id}': {e}")



def eliminar_correo(message_id: str) -> dict:
    """Elimina un correo electrónico de Outlook."""
    _actualizar_headers()
    url = f"{BASE_URL}/{message_id}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Correo '{message_id}' eliminado.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar el correo '{message_id}': {e}")
        raise Exception(f"Error al eliminar el correo '{message_id}': {e}")
