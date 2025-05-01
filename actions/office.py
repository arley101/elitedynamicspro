import logging
import requests
import json # Para manejo de errores
import os # No se usaba, pero podría ser útil para paths
from typing import Dict, List, Optional, Union, Any

# Usar el logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar constantes globales desde __init__.py
try:
    from .. import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    # Fallback si se ejecuta standalone
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_TIMEOUT = 45
    logger.warning("No se pudo importar BASE_URL/GRAPH_API_TIMEOUT desde el padre, usando defaults.")


# ---- WORD ONLINE (via OneDrive /me/drive) ----
# Requieren headers delegados

def crear_documento_word(headers: Dict[str, str], nombre_archivo: str, ruta: str = "/") -> dict:
    """Crea un nuevo documento de Word (.docx) en OneDrive. Requiere headers."""
    # Asegurar extensión .docx
    if not nombre_archivo.lower().endswith(".docx"):
        nombre_archivo += ".docx"

    # Construir la ruta completa en OneDrive
    drive_path = f"/me/drive/root:{ruta.strip('/')}/{nombre_archivo}:" if ruta != "/" else f"/me/drive/root:/{nombre_archivo}:"
    url = f"{BASE_URL}{drive_path}/content"

    # PUT en /content con cuerpo vacío crea el archivo
    # Alternativa: POST a /children con {"name": "...", "file": {}}
    # Usaremos PUT como en el original

    # No se necesita body para PUT en /content para crear archivo vacío
    # body = { "@odata.type": "microsoft.graph.file" } # Esto es para POST a /children

    # Headers para la creación (PUT en /content asume binario por defecto)
    create_headers = headers.copy()
    # Content-Type no es estrictamente necesario si el cuerpo es vacío,
    # pero podríamos poner 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    create_headers.setdefault('Content-Type', 'application/octet-stream') # O el tipo MIME correcto

    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PUT {BASE_URL}{drive_path}/content (Creando Word '{nombre_archivo}' en ruta '{ruta}')")
        # Enviar PUT sin cuerpo (data=None o data='')
        response = requests.put(url, headers=create_headers, data=b'', timeout=GRAPH_API_TIMEOUT) # Enviar bytes vacíos
        response.raise_for_status() # Espera 201 Created
        data = response.json()
        logger.info(f"Documento Word '{nombre_archivo}' creado en ruta '{ruta}'. ID: {data.get('id')}")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en crear_documento_word: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_documento_word: {e}", exc_info=True)
        raise

def insertar_texto_word(headers: Dict[str, str], item_id: str, texto: str) -> dict:
    """
    Actualiza el contenido de un documento Word con el texto proporcionado.
    ¡PRECAUCIÓN: PUT a /content REEMPLAZA todo el contenido existente!
    Requiere headers delegados.
    """
    url = f"{BASE_URL}/me/drive/items/{item_id}/content"
    # Headers para actualizar contenido con texto plano
    update_headers = headers.copy()
    update_headers['Content-Type'] = 'text/plain' # Indicar que enviamos texto plano

    response: Optional[requests.Response] = None
    try:
        logger.warning(f"API Call: PUT {url} (REEMPLAZANDO contenido Word ID '{item_id}' con texto plano)")
        response = requests.put(url, headers=update_headers, data=texto.encode('utf-8'), timeout=GRAPH_API_TIMEOUT * 2)
        response.raise_for_status() # Espera 200 OK
        data = response.json() # La respuesta contiene metadatos del archivo actualizado
        logger.info(f"Contenido del documento Word ID '{item_id}' reemplazado.")
        return data # Devolver metadatos actualizados
        # Devolver un status simple podría ser mejor:
        # return {"status": "Contenido Reemplazado", "code": response.status_code, "id": item_id}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en insertar_texto_word (reemplazar): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en insertar_texto_word (reemplazar): {e}", exc_info=True)
        raise

def obtener_documento_word(headers: Dict[str, str], item_id: str) -> bytes:
    """Obtiene el contenido binario (.docx) de un documento de Word. Requiere headers."""
    url = f"{BASE_URL}/me/drive/items/{item_id}/content"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo contenido Word ID '{item_id}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT * 2)
        response.raise_for_status()
        logger.info(f"Contenido Word ID '{item_id}' obtenido.")
        return response.content # Devuelve los bytes del archivo .docx
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en obtener_documento_word: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_documento_word: {e}", exc_info=True)
        raise


# ---- EXCEL ONLINE (via OneDrive /me/drive) ----
# Requieren headers delegados

def crear_excel(headers: Dict[str, str], nombre_archivo: str, ruta: str = "/") -> dict:
    """Crea un nuevo libro de Excel (.xlsx) en OneDrive. Requiere headers."""
    if not nombre_archivo.lower().endswith(".xlsx"):
        nombre_archivo += ".xlsx"
    drive_path = f"/me/drive/root:{ruta.strip('/')}/{nombre_archivo}:" if ruta != "/" else f"/me/drive/root:/{nombre_archivo}:"
    url = f"{BASE_URL}{drive_path}/content"

    create_headers = headers.copy()
    create_headers.setdefault('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PUT {url} (Creando Excel '{nombre_archivo}' en ruta '{ruta}')")
        response = requests.put(url, headers=create_headers, data=b'', timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 201 Created
        data = response.json()
        logger.info(f"Libro Excel '{nombre_archivo}' creado en ruta '{ruta}'. ID: {data.get('id')}")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en crear_excel: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_excel: {e}", exc_info=True)
        raise

def escribir_celda_excel(headers: Dict[str, str], item_id: str, hoja: str, celda: str, valor: Union[str, int, float, bool]) -> dict:
    """Escribe un valor en una celda Excel. Requiere headers."""
    # La dirección de la celda puede ser 'A1' o 'Sheet1!A1'
    # El endpoint maneja la hoja en la URL, así que 'celda' debe ser tipo 'A1'
    # Graph infiere el tipo de valor (string, number, boolean)
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/range(address='{celda}')"
    body = {"values": [[valor]]} # Debe ser una lista de listas
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Escribiendo en celda '{celda}' hoja '{hoja}' item '{item_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.patch(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 200 OK
        data = response.json()
        logger.info(f"Valor escrito en celda '{celda}', hoja '{hoja}', item '{item_id}'.")
        return data # Devuelve info del rango actualizado
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en escribir_celda_excel: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en escribir_celda_excel: {e}", exc_info=True)
        raise

def leer_celda_excel(headers: Dict[str, str], item_id: str, hoja: str, celda: str) -> dict:
    """Lee el valor de una celda Excel. Requiere headers."""
    # Similar a escribir, pero con GET y seleccionando 'values' o 'text'
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/range(address='{celda}')?$select=text,values,address"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Leyendo celda '{celda}' hoja '{hoja}' item '{item_id}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Valor leído de celda '{celda}', hoja '{hoja}', item '{item_id}'.")
        # Devuelve el objeto range con text, values, address
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en leer_celda_excel: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en leer_celda_excel: {e}", exc_info=True)
        raise

def crear_tabla_excel(headers: Dict[str, str], item_id: str, hoja: str, rango: str, tiene_headers: bool = False) -> dict:
    """Crea una tabla Excel en un rango. Requiere headers."""
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/tables/add" # Endpoint para añadir tabla
    body = {
        "address": f"{hoja}!{rango}", # Dirección completa requerida aquí
        "hasHeaders": tiene_headers
        }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando tabla en rango '{rango}' hoja '{hoja}' item '{item_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 201 Created
        data = response.json()
        table_id = data.get('id')
        logger.info(f"Tabla creada ID '{table_id}' en hoja '{hoja}', item '{item_id}'.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en crear_tabla_excel: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_tabla_excel: {e}", exc_info=True)
        raise

def agregar_datos_tabla_excel(headers: Dict[str, str], item_id: str, tabla_id_o_nombre: str, valores: List[List[Any]]) -> dict:
    """Agrega filas de datos a una tabla Excel. Requiere headers."""
    # El endpoint usa ID o nombre de tabla
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/tables/{tabla_id_o_nombre}/rows" # POST a /rows
    body = {
        # "index": null, # Añadir al final
        "values": valores # Lista de listas, cada lista interna es una fila
    }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Agregando {len(valores)} filas a tabla '{tabla_id_o_nombre}' item '{item_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 201 Created
        data = response.json()
        logger.info(f"Datos agregados a tabla '{tabla_id_o_nombre}', item '{item_id}'.")
        return data # Devuelve info sobre las filas añadidas (index)
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en agregar_datos_tabla_excel: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en agregar_datos_tabla_excel: {e}", exc_info=True)
        raise
