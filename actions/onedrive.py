import logging
import os
import requests
from auth import obtener_token  # Importante: Importar la función obtener_token
from typing import Dict, Optional, Union

# Configuración básica de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Variables de entorno (¡CRUCIALES!)
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
GRAPH_SCOPE = os.getenv('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')  # Valor por defecto

# Verificar variables de entorno
if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE]):
    logging.error("❌ Faltan variables de entorno (CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE). La función no puede funcionar.")
    raise Exception("Faltan variables de entorno.")

BASE_URL = "https://graph.microsoft.com/v1.0/me/drive"
HEADERS = {
    'Authorization': None,  # Inicialmente None, se actualiza con cada request
    'Content-Type': 'application/json'
}

# Función para obtener el token y actualizar los HEADERS
def _actualizar_headers() -> None:
    """Obtiene un nuevo token de acceso y actualiza el diccionario HEADERS."""
    try:
        HEADERS['Authorization'] = f'Bearer {obtener_token()}'
    except Exception as e:  # Captura la excepción de obtener_token
        logging.error(f"❌ Error al obtener el token: {e}")
        raise Exception(f"Error al obtener el token: {e}")

# ---- FUNCIONES DE GESTIÓN DE ARCHIVOS Y CARPETAS EN ONEDRIVE ----

def listar_archivos(ruta: str = "/") -> dict:
    """Lista archivos y carpetas en una ruta de OneDrive, manejando paginación."""
    _actualizar_headers()
    url = f"{BASE_URL}/root:{ruta}:/children"
    try:
        all_items = []
        while url:
            response = requests.get(url, headers=HEADERS)
            response.raise_for_status()
            data = response.json()
            all_items.extend(data.get('value', []))
            url = data.get('@odata.nextLink')  # URL para la siguiente página
            if url:
                _actualizar_headers() # Actualiza el token si hay paginacion
        logging.info(f"Listados archivos y carpetas en ruta '{ruta}'. Total: {len(all_items)}")
        return {'value': all_items}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar archivos en ruta '{ruta}': {e}")
        raise Exception(f"Error al listar archivos en ruta '{ruta}': {e}")



def subir_archivo(nombre_archivo: str, contenido: Union[str, bytes], ruta: str = "/") -> dict:
    """Sube un archivo a OneDrive."""
    _actualizar_headers()
    url = f"{BASE_URL}/root:{ruta}/{nombre_archivo}:/content"

    # Asegurarse de que el contenido sea bytes
    if isinstance(contenido, str):
        contenido_bytes = contenido.encode('utf-8')
    else:
        contenido_bytes = contenido

    try:
        response = requests.put(url, headers={'Authorization': HEADERS['Authorization']}, data=contenido_bytes)
        response.raise_for_status()
        logging.info(f"Archivo '{nombre_archivo}' subido a ruta '{ruta}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al subir el archivo '{nombre_archivo}' a ruta '{ruta}': {e}")
        raise Exception(f"Error al subir el archivo '{nombre_archivo}' a ruta '{ruta}': {e}")



def descargar_archivo(nombre_archivo: str, ruta: str = "/") -> bytes:
    """Descarga un archivo de OneDrive y devuelve el contenido como bytes."""
    _actualizar_headers()
    url = f"{BASE_URL}/root:{ruta}/{nombre_archivo}:/content"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()  # Lanza excepción para errores
        logging.info(f"Archivo '{nombre_archivo}' descargado de ruta '{ruta}'.")
        return response.content  # Devuelve el contenido como bytes
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al descargar el archivo '{nombre_archivo}' de ruta '{ruta}': {e}")
        raise Exception(f"Error al descargar el archivo '{nombre_archivo}' de ruta '{ruta}': {e}")



def eliminar_archivo(nombre_archivo: str, ruta: str = "/") -> dict:
    """Elimina un archivo de OneDrive."""
    _actualizar_headers()
    url = f"{BASE_URL}/root:{ruta}/{nombre_archivo}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Archivo '{nombre_archivo}' eliminado de ruta '{ruta}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar el archivo '{nombre_archivo}' de ruta '{ruta}': {e}")
        raise Exception(f"Error al eliminar el archivo '{nombre_archivo}' de ruta '{ruta}': {e}")



def crear_carpeta(nombre_carpeta: str, ruta: str = "/") -> dict:
    """Crea una carpeta en OneDrive."""
    _actualizar_headers()
    url = f"{BASE_URL}/root:{ruta}:/children"
    body = {
        "name": nombre_carpeta,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename"  # Para renombrar si ya existe
    }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Carpeta '{nombre_carpeta}' creada en ruta '{ruta}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear la carpeta '{nombre_carpeta}' en ruta '{ruta}': {e}")
        raise Exception(f"Error al crear la carpeta '{nombre_carpeta}' en ruta '{ruta}': {e}")



def mover_archivo(nombre_archivo: str, ruta_origen: str = "/", ruta_destino: str = "/NuevaCarpeta") -> dict:
    """Mueve un archivo o carpeta de OneDrive."""
    _actualizar_headers()
    url = f"{BASE_URL}/root:{ruta_origen}/{nombre_archivo}"
    body = {
        "parentReference": {
            "path": f"/drive/root:{ruta_destino}"  # La ruta debe ser relativa a la raíz de la unidad
        },
        "name": nombre_archivo #Conserva el nombre original
    }
    try:
        response = requests.patch(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Archivo/Carpeta '{nombre_archivo}' movido de '{ruta_origen}' a '{ruta_destino}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al mover archivo/carpeta '{nombre_archivo}' de '{ruta_origen}' a '{ruta_destino}': {e}")
        raise Exception(f"Error al mover archivo/carpeta '{nombre_archivo}' de '{ruta_origen}' a '{ruta_destino}': {e}")



def copiar_archivo(nombre_archivo: str, ruta_origen: str = "/", ruta_destino: str = "/Copias") -> dict:
    """Copia un archivo o carpeta de OneDrive."""
    _actualizar_headers()
    url = f"{BASE_URL}/root:{ruta_origen}/{nombre_archivo}/copy"
    body = {
        "parentReference": {
            "path": f"/drive/root:{ruta_destino}"  # La ruta debe ser relativa a la raíz de la unidad
        },
        "name": f"Copia_{nombre_archivo}"  # Puedes cambiar el nombre si lo deseas
    }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Copia de '{nombre_archivo}' iniciada en ruta '{ruta_destino}'.")
        return response.json()  # La respuesta de la copia es asíncrona
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al copiar archivo/carpeta '{nombre_archivo}' de '{ruta_origen}' a '{ruta_destino}': {e}")
        raise Exception(f"Error al copiar archivo/carpeta '{nombre_archivo}' de '{ruta_origen}' a '{ruta_destino}': {e}")


def obtener_metadatos_archivo(nombre_archivo: str, ruta: str = "/") -> dict:
    """Obtiene los metadatos de un archivo o carpeta en OneDrive."""
    _actualizar_headers()
    url = f"{BASE_URL}/root:{ruta}/{nombre_archivo}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Obtenidos metadatos de archivo/carpeta '{nombre_archivo}' en ruta '{ruta}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener metadatos de archivo/carpeta '{nombre_archivo}' en ruta '{ruta}': {e}")
        raise Exception(f"Error al obtener metadatos de archivo/carpeta '{nombre_archivo}' en ruta '{ruta}': {e}")

def actualizar_metadatos_archivo(nombre_archivo: str, nuevos_valores: dict, ruta:str = "/")->dict:
    """Actualiza los metadatos de un archivo o carpeta"""
    _actualizar_headers()
    url = f"{BASE_URL}/root:{ruta}/{nombre_archivo}"
    try:
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status()
        logging.info(f"Metadatos de archivo/carpeta '{nombre_archivo}' actualizados en ruta '{ruta}': Nuevos valores: {nuevos_valores}")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al actualizar metadatos de archivo/carpeta '{nombre_archivo}' en ruta '{ruta}': {e}")
        raise Exception(f"Error al actualizar metadatos de archivo/carpeta '{nombre_archivo}' en ruta '{ruta}': {e}")
