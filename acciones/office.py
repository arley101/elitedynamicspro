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

# Verificar variables de entorno (¡CRUCIALES!)
if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE]):
    logging.error("❌ Faltan variables de entorno (CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE). La función no puede funcionar.")
    raise Exception("Faltan variables de entorno.")

BASE_URL = "https://graph.microsoft.com/v1.0"
HEADERS = {
    'Authorization': None,  # Inicialmente None, se actualiza con cada request
    'Content-Type': 'application/json'
}


# Función para obtener el token y actualizar los HEADERS
def _actualizar_headers() -> None:
    """Obtiene un nuevo token de acceso y actualiza el diccionario HEADERS."""
    try:
        HEADERS['Authorization'] = f'Bearer {obtener_token()}'
    except Exception as e:
        logging.error(f"❌ Error al obtener el token: {e}")
        raise Exception(f"Error al obtener el token: {e}")



# ---- WORD ONLINE ----
def crear_documento_word(nombre_archivo: str, ruta: str = "/") -> dict:
    """Crea un nuevo documento de Word en OneDrive."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/drive/root:{ruta}/{nombre_archivo}.docx"
    body = {
        "@odata.type": "microsoft.graph.file"  # Especifica el tipo de archivo
    }
    try:
        response = requests.put(url, headers=HEADERS, json=body)  # Usa PUT para crear archivos con contenido vacío
        response.raise_for_status()
        data = response.json()
        logging.info(f"Documento de Word '{nombre_archivo}.docx' creado en ruta '{ruta}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear documento de Word '{nombre_archivo}.docx' en ruta '{ruta}': {e}")
        raise Exception(f"Error al crear documento de Word '{nombre_archivo}.docx' en ruta '{ruta}': {e}")



def insertar_texto_word(item_id: str, texto: str) -> dict:
    """Inserta texto al final de un documento de Word existente en OneDrive."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/drive/items/{item_id}/content"  # Inserta al final del archivo
    headers = {
        'Authorization': HEADERS['Authorization'],
        'Content-Type': 'text/plain'  # Content-Type para texto plano
    }
    try:
        response = requests.put(url, headers=headers, data=texto.encode('utf-8'))
        response.raise_for_status()
        logging.info(f"Texto insertado en el documento Word con ID '{item_id}'.")
        return {"status": "Texto insertado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al insertar texto en el documento Word con ID '{item_id}': {e}")
        raise Exception(f"Error al insertar texto en el documento Word con ID '{item_id}': {e}")



def obtener_documento_word(item_id: str) -> bytes:
    """Obtiene el contenido de un documento de Word."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/drive/items/{item_id}/content"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Obtenido contenido del documento Word con ID '{item_id}'.")
        return response.content
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener el contenido del documento Word con ID '{item_id}': {e}")
        raise Exception(f"Error al obtener el contenido del documento Word con ID '{item_id}': {e}")


# ---- EXCEL ONLINE ----

def crear_excel(nombre_archivo: str, ruta: str = "/") -> dict:
    """Crea un nuevo libro de Excel en OneDrive."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/drive/root:{ruta}/{nombre_archivo}.xlsx"
    body = {
        "@odata.type": "microsoft.graph.file"
    }
    try:
        response = requests.put(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Libro de Excel '{nombre_archivo}.xlsx' creado en ruta '{ruta}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear el libro de Excel '{nombre_archivo}.xlsx' en ruta '{ruta}': {e}")
        raise Exception(f"Error al crear el libro de Excel '{nombre_archivo}.xlsx' en ruta '{ruta}': {e}")



def escribir_celda_excel(item_id: str, hoja: str, celda: str, valor: str) -> dict:
    """Escribe un valor en una celda de una hoja de cálculo de Excel."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/range(address='{celda}')"
    body = {"values": [[valor]]}
    try:
        response = requests.patch(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Escrito '{valor}' en la celda '{celda}' de la hoja '{hoja}' del libro de Excel con ID '{item_id}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al escribir en la celda '{celda}' de la hoja '{hoja}' del libro de Excel con ID '{item_id}': {e}")
        raise Exception(f"Error al escribir en la celda '{celda}' de la hoja '{hoja}' del libro de Excel con ID '{item_id}': {e}")



def leer_celda_excel(item_id: str, hoja: str, celda: str) -> dict:
    """Lee el valor de una celda de una hoja de cálculo de Excel."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/range(address='{celda}')"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Leído el valor de la celda '{celda}' de la hoja '{hoja}' del libro de Excel con ID '{item_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al leer la celda '{celda}' de la hoja '{hoja}' del libro de Excel con ID '{item_id}': {e}")
        raise Exception(f"Error al leer la celda '{celda}' de la hoja '{hoja}' del libro de Excel con ID '{item_id}': {e}")



def crear_tabla_excel(item_id: str, hoja: str, rango: str) -> dict:
    """Crea una tabla en una hoja de cálculo de Excel."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/tables"
    body = {
        "address": rango,
        "hasHeaders": False  # Puedes cambiar esto según tus necesidades
        }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        table_id = data.get('id')
        logging.info(f"Tabla con ID '{table_id}' creada en la hoja '{hoja}' del libro de Excel con ID '{item_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear tabla en la hoja '{hoja}' del libro de Excel con ID '{item_id}': {e}")
        raise Exception(f"Error al crear tabla en la hoja '{hoja}' del libro de Excel con ID '{item_id}': {e}")



def agregar_datos_tabla_excel(item_id: str, tabla_id: str, valores: List[List[str]]) -> dict:
    """Agrega datos a una tabla existente en una hoja de cálculo de Excel."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/tables/{tabla_id}/rows/add"
    body = {
        "values": valores
    }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        logging.info(f"Datos agregados a la tabla con ID '{tabla_id}' del libro de Excel con ID '{item_id}'.")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al agregar datos a la tabla con ID '{tabla_id}' del libro de Excel con ID '{item_id}': {e}")
        raise Exception(f"Error al agregar datos a la tabla con ID '{tabla_id}': {e}")
