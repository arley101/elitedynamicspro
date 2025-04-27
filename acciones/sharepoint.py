import json
import logging
import requests
from auth import obtener_token  # Importante: Importar desde auth.py
import os
import azure.functions as func  # Si este archivo es el __init__.py de la función
from typing import Dict, List, Optional, Union

# Configuración básica de logging (si este archivo es el __init__.py)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Variables de entorno (¡CRUCIALES!)
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
GRAPH_SCOPE = os.getenv('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')  # Valor por defecto

# Verificar variables de entorno (¡CRUCIAL!)
if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE]):
    logging.error("❌ Faltan variables de entorno (CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE). La función no puede funcionar.")
    # Si este archivo es el __init__.py, puedes lanzar una excepción para detener la ejecución
    # raise Exception("Faltan variables de entorno.")
    # O, si estás dentro de una función de Azure, devolver un error HTTP
    def main(req: func.HttpRequest) -> func.HttpResponse:
        return func.HttpResponse(json.dumps({"error": "Faltan variables de entorno."}), status_code=500)
    # Y luego salir del resto del script
    import sys
    sys.exit(1)


BASE_URL = "https://graph.microsoft.com/v1.0"
HEADERS = {
    'Authorization': None,  # Inicialmente None, se actualizará en cada llamada
    'Content-Type': 'application/json'
}
MAX_PAGINATION_SIZE = 999  # Define el tamaño máximo de página permitido por Graph (y por nosotros)


# Función para obtener el SITE_ID dinámicamente
def obtener_site_root() -> str:
    """Obtiene el ID del sitio raíz de SharePoint."""
    _actualizar_headers()  # Obtener el token antes de la llamada
    url = f"{BASE_URL}/sites/root"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        site_id = response.json().get('id')
        logging.info(f"SITE_ID obtenido: {site_id}")
        return site_id
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener el SITE_ID: {e}")
        raise Exception(f"Error al obtener el SITE_ID: {e}")



# Función para obtener el token y actualizar los HEADERS (¡CRUCIAL!)
def _actualizar_headers() -> None:
    """Obtiene un nuevo token de acceso y actualiza el diccionario HEADERS."""
    try:
        HEADERS['Authorization'] = f'Bearer {obtener_token()}'
    except Exception as e:  # Captura excepciones al obtener el token
        logging.error(f"❌ Error al obtener el token: {e}")
        raise Exception(f"Error al obtener el token: {e}")


# ---- FUNCIONES DE MEMORIA Y GESTIÓN (LISTAS) ----

def crear_lista(nombre_lista: str) -> dict:
    """Crea una nueva lista de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/lists"
    body = {
        "displayName": nombre_lista,
        "columns": [
            {"name": "Clave", "text": {}},
            {"name": "Valor", "text": {}}
        ],
        "list": {"template": "genericList"}
    }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Lista '{nombre_lista}' creada exitosamente.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear la lista '{nombre_lista}': {e}")
        raise Exception(f"Error al crear la lista '{nombre_lista}': {e}")



def listar_listas() -> dict:
    """Lista todas las listas en el sitio de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/lists"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        logging.info("Listando listas en el sitio.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar listas: {e}")
        raise Exception(f"Error al listar listas: {e}")



def agregar_elemento(nombre_lista: str, clave: str, valor: str) -> dict:
    """Agrega un elemento a una lista de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/lists/{nombre_lista}/items"
    body = {"fields": {"Clave": clave, "Valor": valor}}
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Elemento agregado a la lista '{nombre_lista}': Clave='{clave}', Valor='{valor}'")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al agregar elemento a la lista '{nombre_lista}': {e}")
        raise Exception(f"Error al agregar elemento a la lista '{nombre_lista}': {e}")



def listar_elementos(nombre_lista: str, expand_fields: bool = True) -> dict:
    """
    Lista elementos de una lista de SharePoint, manejando la paginación.

    :param nombre_lista: El nombre de la lista.
    :param expand_fields: Indica si se deben expandir los campos de los elementos.
    :return: Un diccionario con todos los elementos de la lista.
    """
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/lists/{nombre_lista}/items"
    if expand_fields:
        url += "?expand=fields"

    try:
        all_items = []
        while url:  # Sigue la paginación hasta que no haya más resultados
            response = requests.get(url, headers=HEADERS)
            response.raise_for_status()
            data = response.json()
            all_items.extend(data.get('value', []))  # Agrega los elementos de la página actual

            url = data.get('@odata.nextLink')  # Obtiene la URL de la siguiente página, si existe
            if url:
                _actualizar_headers()

        logging.info(f"Obtenidos todos los elementos de la lista '{nombre_lista}'. Total: {len(all_items)}")
        return {'value': all_items}  # Devuelve los elementos bajo la clave 'value', consistente con la API
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar elementos de la lista '{nombre_lista}': {e}")
        raise Exception(f"Error al listar elementos de la lista '{nombre_lista}': {e}")



def actualizar_elemento(nombre_lista: str, item_id: str, nuevos_valores: dict) -> dict:
    """Actualiza un elemento de una lista de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/lists/{nombre_lista}/items/{item_id}/fields"
    try:
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status()
        logging.info(f"Elemento '{item_id}' actualizado en la lista '{nombre_lista}'. Nuevos valores: {nuevos_valores}")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al actualizar elemento '{item_id}' en la lista '{nombre_lista}': {e}")
        raise Exception(f"Error al actualizar elemento '{item_id}' en la lista '{nombre_lista}': {e}")



def eliminar_elemento(nombre_lista: str, item_id: str) -> dict:
    """Elimina un elemento de una lista de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/lists/{nombre_lista}/items/{item_id}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Elemento '{item_id}' eliminado de la lista '{nombre_lista}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar elemento '{item_id}' de la lista '{nombre_lista}': {e}")
        raise Exception(f"Error al eliminar elemento '{item_id}' de la lista '{nombre_lista}': {e}")



# ---- GESTIÓN DE DOCUMENTOS EN SHAREPOINT ----

def listar_documentos_biblioteca(biblioteca: str = "Documents") -> dict:
    """Lista documentos en una biblioteca de documentos de SharePoint, manejando paginación."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/drives/{biblioteca}/root/children"
    try:
        all_files = []
        while url:
            response = requests.get(url, headers=HEADERS)
            response.raise_for_status()
            data = response.json()
            all_files.extend(data.get('value', []))
            url = data.get('@odata.nextLink')
            if url:
                _actualizar_headers()
        logging.info(f"Listados documentos de la biblioteca '{biblioteca}'. Total: {len(all_files)}")
        return {'value': all_files}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar documentos de la biblioteca '{biblioteca}': {e}")
        raise Exception(f"Error al listar documentos de la biblioteca '{biblioteca}': {e}")



def subir_documento(nombre_archivo: str, contenido_base64: Union[str, bytes], biblioteca: str = "Documents") -> dict:
    """Sube un documento a una biblioteca de documentos de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/drives/{biblioteca}/root:/{nombre_archivo}:/content"
    # Asegurarse de que el contenido sea bytes
    if isinstance(contenido_base64, str):
        contenido_bytes = contenido_base64.encode('utf-8')  # Codificar la cadena a bytes
    else:
        contenido_bytes = contenido_base64

    try:
        response = requests.put(url, headers=HEADERS, data=contenido_bytes)
        response.raise_for_status()
        logging.info(f"Documento '{nombre_archivo}' subido a la biblioteca '{biblioteca}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al subir el documento '{nombre_archivo}' a la biblioteca '{biblioteca}': {e}")
        raise Exception(f"Error al subir el documento '{nombre_archivo}' a la biblioteca '{biblioteca}': {e}")



def eliminar_documento(nombre_archivo: str, biblioteca: str = "Documents") -> dict:
    """Elimina un documento de una biblioteca de documentos de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/drives/{biblioteca}/root:/{nombre_archivo}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Documento '{nombre_archivo}' eliminado de la biblioteca '{biblioteca}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar el documento '{nombre_archivo}' de la biblioteca '{biblioteca}': {e}")
        raise Exception(f"Error al eliminar el documento '{nombre_archivo}' de la biblioteca '{biblioteca}': {e}")



# ---- FUNCIONES AVANZADAS DE ARCHIVOS (POSIBLES EXTENSIONES) ----

def crear_carpeta_biblioteca(biblioteca: str, nombre_carpeta: str) -> dict:
    """Crea una nueva carpeta dentro de una biblioteca de documentos de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/drives/{biblioteca}/root/children"
    body = {
        "name": nombre_carpeta,
        "folder": {}
    }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Carpeta '{nombre_carpeta}' creada en la biblioteca '{biblioteca}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear la carpeta '{nombre_carpeta}' en la biblioteca '{biblioteca}': {e}")
        raise Exception(f"Error al crear la carpeta '{nombre_carpeta}' en la biblioteca '{biblioteca}': {e}")



def mover_archivo(biblioteca: str, nombre_archivo: str, nueva_ubicacion: str) -> dict:
    """Mueve un archivo dentro de una biblioteca de documentos de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/drives/{biblioteca}/root:/{nombre_archivo}"
    body = {
        "parentReference": {
            "path": f"/root:{nueva_ubicacion}"  # La ruta debe ser relativa a la raíz
        }
    }
    try:
        response = requests.patch(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Archivo '{nombre_archivo}' movido en la biblioteca '{biblioteca}' a '{nueva_ubicacion}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al mover el archivo '{nombre_archivo}' en la biblioteca '{biblioteca}' a '{nueva_ubicacion}': {e}")
        raise Exception(f"Error al mover el archivo '{nombre_archivo}' en la biblioteca '{biblioteca}': {e}")



def copiar_archivo(biblioteca: str, nombre_archivo: str, nueva_ubicacion: str) -> dict:
    """Copia un archivo dentro de una biblioteca de documentos de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/drives/{biblioteca}/root:/{nombre_archivo}/copy"
    body = {
        "parentReference": {
            "path": f"/root:{nueva_ubicacion}"  # La ruta debe ser relativa a la raíz
        },
        "name": nombre_archivo  # Puedes cambiar el nombre si lo deseas
    }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Archivo '{nombre_archivo}' copiado en la biblioteca '{biblioteca}' a '{nueva_ubicacion}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al copiar el archivo '{nombre_archivo}' en la biblioteca '{biblioteca}' a '{nueva_ubicacion}': {e}")
        raise Exception(f"Error al copiar el archivo '{nombre_archivo}' en la biblioteca '{biblioteca}': {e}")



def obtener_metadatos_archivo(biblioteca: str, nombre_archivo: str) -> dict:
    """Obtiene los metadatos de un archivo dentro de una biblioteca de documentos de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/drives/{biblioteca}/root:/{nombre_archivo}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Obteniendo metadatos del archivo '{nombre_archivo}' en la biblioteca '{biblioteca}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener los metadatos del archivo '{nombre_archivo}' en la biblioteca '{biblioteca}': {e}")
        raise Exception(f"Error al obtener los metadatos del archivo '{nombre_archivo}' en la biblioteca '{biblioteca}': {e}")



def actualizar_metadatos_archivo(biblioteca: str, nombre_archivo: str, nuevos_valores: dict) -> dict:
    """Actualiza los metadatos de un archivo dentro de una biblioteca de documentos de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/drives/{biblioteca}/root:/{nombre_archivo}"
    try:
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status()
        logging.info(f"Metadatos del archivo '{nombre_archivo}' actualizados en la biblioteca '{biblioteca}'. Nuevos valores: {nuevos_valores}")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al actualizar los metadatos del archivo '{nombre_archivo}' en la biblioteca '{biblioteca}': {e}")
        raise Exception(f"Error al actualizar los metadatos del archivo '{nombre_archivo}' en la biblioteca '{biblioteca}': {e}")



def obtener_contenido_archivo(biblioteca: str, nombre_archivo: str) -> bytes:
    """Obtiene el contenido binario de un archivo dentro de una biblioteca de documentos de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/drives/{biblioteca}/root:/{nombre_archivo}/content"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Obteniendo contenido del archivo '{nombre_archivo}' en la biblioteca '{biblioteca}'.")
        return response.content
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener el contenido del archivo '{nombre_archivo}' en la biblioteca '{biblioteca}': {e}")
        raise Exception(f"Error al obtener el contenido del archivo '{nombre_archivo}' en la biblioteca '{biblioteca}': {e}")



def actualizar_contenido_archivo(biblioteca: str, nombre_archivo: str, nuevo_contenido: bytes) -> dict:
    """Actualiza el contenido binario de un archivo dentro de una biblioteca de documentos de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/drives/{biblioteca}/root:/{nombre_archivo}/content"
    try:
        response = requests.put(url, headers=HEADERS, data=nuevo_contenido)
        response.raise_for_status()
        logging.info(f"Contenido del archivo '{nombre_archivo}' actualizado en la biblioteca '{biblioteca}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al actualizar el contenido del archivo '{nombre_archivo}' en la biblioteca '{biblioteca}': {e}")
        raise Exception(f"Error al actualizar el contenido del archivo '{nombre_archivo}' en la biblioteca '{biblioteca}': {e}")



def crear_enlace_compartido_archivo(biblioteca: str, nombre_archivo: str, tipo_enlace: str = "view", alcance: str = "anonymous") -> dict:
    """Crea un enlace compartido para un archivo dentro de una biblioteca de documentos de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/drives/{biblioteca}/root:/{nombre_archivo}/createLink"
    body = {
        "type": tipo_enlace,
        "scope": alcance
    }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Enlace compartido creado para el archivo '{nombre_archivo}' en la biblioteca '{biblioteca}'. Tipo: {tipo_enlace}, Alcance: {alcance}")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear el enlace compartido para el archivo '{nombre_archivo}' en la biblioteca '{biblioteca}': {e}")
        raise Exception(f"Error al crear el enlace compartido para el archivo '{nombre_archivo}' en la biblioteca '{biblioteca}': {e}")



def eliminar_archivo(biblioteca: str, nombre_archivo: str) -> dict:
    """Elimina un archivo de una biblioteca de documentos de SharePoint."""
    _actualizar_headers()
    url = f"{BASE_URL}/sites/{obtener_site_root()}/drives/{biblioteca}/root:/{nombre_archivo}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Archivo '{nombre_archivo}' eliminado de la biblioteca '{biblioteca}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar el archivo '{nombre_archivo}' de la biblioteca '{biblioteca}': {e}")
        raise Exception(f"Error al eliminar el archivo '{nombre_archivo}' de la biblioteca '{biblioteca}': {e}")
