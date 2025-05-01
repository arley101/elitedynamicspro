# actions/office.py (Refactorizado)

import logging
import requests
import os
from typing import Dict, List, Optional, Union

# Usar logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar constantes globales desde __init__.py
try:
    from .. import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    # Fallback
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_TIMEOUT = 45
    logger.warning("No se pudo importar BASE_URL/GRAPH_API_TIMEOUT desde el padre, usando defaults.")

# (Eliminada configuración redundante y _actualizar_headers)

# ---- WORD ONLINE (Opera sobre /me/drive) ----

def crear_documento_word(headers: Dict[str, str], nombre_archivo: str, ruta: str = "/") -> dict:
    """Crea un nuevo documento de Word en OneDrive (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    # Asegurar que la ruta sea relativa a la raíz y construir endpoint
    safe_path = ruta.strip()
    if not safe_path.startswith('/'): safe_path = '/' + safe_path
    item_path = os.path.join(safe_path, f"{nombre_archivo}.docx").replace('\\','/')
    url = f"{BASE_URL}/me/drive/root:{item_path}"

    # PUT vacío para crear archivo
    body = None # {} # El ejemplo original usaba @odata.type, pero PUT en path crea archivo vacío
    put_headers = headers.copy()
    put_headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' # Opcional

    try:
        logger.info(f"API Call: PUT {url} (Creando Word /me '{nombre_archivo}.docx' en ruta '{ruta}')")
        response = requests.put(url, headers=put_headers, json=body, timeout=GRAPH_API_TIMEOUT) # json=body si se usa, sino data=b''
        response.raise_for_status()
        data = response.json()
        logger.info(f"Documento Word '{nombre_archivo}.docx' creado.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en crear_documento_word: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_documento_word: {e}", exc_info=True)
        raise

def insertar_texto_word(headers: Dict[str, str], item_id: str, texto: str) -> dict:
    """Inserta texto al final de un documento Word en OneDrive (/me). (Requiere revisión, PUT /content sobreescribe)"""
    # ADVERTENCIA: PUT a /content generalmente SOBREESCRIBE el archivo.
    # Para insertar o modificar contenido específico se necesitan APIs más avanzadas o
    # descargar, modificar localmente y volver a subir.
    # Esta implementación PROBABLEMENTE REEMPLACE el contenido con el texto.
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    url = f"{BASE_URL}/me/drive/items/{item_id}/content"
    upload_headers = headers.copy()
    upload_headers['Content-Type'] = 'text/plain' # O el tipo correcto si se quiere mantener formato

    try:
        logger.warning(f"API Call: PUT {url} (INTENTANDO insertar texto en Word /me '{item_id}' - ¡PUEDE SOBREESCRIBIR!)")
        response = requests.put(url, headers=upload_headers, data=texto.encode('utf-8'), timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        logger.info(f"Operación PUT content completada para Word ID '{item_id}'.")
        # PUT /content no devuelve body JSON útil normalmente
        return {"status": "Operación PUT content completada", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en insertar_texto_word: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en insertar_texto_word: {e}", exc_info=True)
        raise

def obtener_documento_word(headers: Dict[str, str], item_id: str) -> bytes:
    """Obtiene el contenido de un documento de Word (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    url = f"{BASE_URL}/me/drive/items/{item_id}/content"
    try:
        logger.info(f"API Call: GET {url} (Obteniendo contenido Word /me '{item_id}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT * 2)
        response.raise_for_status()
        logger.info(f"Contenido Word '{item_id}' obtenido.")
        return response.content
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en obtener_documento_word: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_documento_word: {e}", exc_info=True)
        raise


# ---- EXCEL ONLINE (Opera sobre /me/drive) ----

def crear_excel(headers: Dict[str, str], nombre_archivo: str, ruta: str = "/") -> dict:
    """Crea un nuevo libro de Excel en OneDrive (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    safe_path = ruta.strip()
    if not safe_path.startswith('/'): safe_path = '/' + safe_path
    item_path = os.path.join(safe_path, f"{nombre_archivo}.xlsx").replace('\\','/')
    url = f"{BASE_URL}/me/drive/root:{item_path}"
    body = None # {} # PUT vacío crea archivo
    put_headers = headers.copy()
    put_headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    try:
        logger.info(f"API Call: PUT {url} (Creando Excel /me '{nombre_archivo}.xlsx' en ruta '{ruta}')")
        response = requests.put(url, headers=put_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Libro Excel '{nombre_archivo}.xlsx' creado.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en crear_excel: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_excel: {e}", exc_info=True)
        raise

def escribir_celda_excel(headers: Dict[str, str], item_id: str, hoja: str, celda: str, valor: Union[str, int, float, bool]) -> dict:
    """Escribe un valor en una celda de Excel (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    # Escapar nombre de hoja si contiene caracteres especiales? MSAL maneja esto a veces.
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/range(address='{celda}')"
    # El valor debe estar en un array de arrays
    body = {"values": [[valor]]}
    patch_headers = headers.copy()
    patch_headers['Content-Type'] = 'application/json' # Asegurar JSON para PATCH

    try:
        logger.info(f"API Call: PATCH {url} (Escribiendo celda Excel /me '{hoja}!{celda}' en item '{item_id}')")
        response = requests.patch(url, headers=patch_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        logger.info(f"Valor '{valor}' escrito en celda '{celda}', hoja '{hoja}'.")
        return response.json()
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en escribir_celda_excel: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en escribir_celda_excel: {e}", exc_info=True)
        raise

def leer_celda_excel(headers: Dict[str, str], item_id: str, hoja: str, celda: str) -> dict:
    """Lee el valor de una celda de Excel (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/range(address='{celda}')?$select=values" # Pedir solo values
    try:
        logger.info(f"API Call: GET {url} (Leyendo celda Excel /me '{hoja}!{celda}' en item '{item_id}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Valor leído de celda '{celda}', hoja '{hoja}'.")
        return data # Devuelve un objeto Range con 'values'
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en leer_celda_excel: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en leer_celda_excel: {e}", exc_info=True)
        raise

def crear_tabla_excel(headers: Dict[str, str], item_id: str, hoja: str, rango: str, tiene_headers: bool = False) -> dict:
    """Crea una tabla en Excel (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/tables/add" # Endpoint 'add'
    body = {"address": rango, "hasHeaders": tiene_headers}
    post_headers = headers.copy()
    post_headers['Content-Type'] = 'application/json'

    try:
        logger.info(f"API Call: POST {url} (Creando tabla Excel /me en '{hoja}' rango '{rango}')")
        response = requests.post(url, headers=post_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        table_id = data.get('id')
        logger.info(f"Tabla Excel creada ID '{table_id}'.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en crear_tabla_excel: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_tabla_excel: {e}", exc_info=True)
        raise

def agregar_datos_tabla_excel(headers: Dict[str, str], item_id: str, tabla_id: str, valores: List[List[Union[str, int, float, bool]]]) -> dict:
    """Agrega filas de datos a una tabla Excel (/me)."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    # La tabla se puede referenciar por ID o nombre
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/tables/{tabla_id}/rows/add"
    body = {"values": valores} # index opcional para insertar en posición específica
    post_headers = headers.copy()
    post_headers['Content-Type'] = 'application/json'

    try:
        logger.info(f"API Call: POST {url} (Agregando datos tabla Excel /me '{tabla_id}' en item '{item_id}')")
        response = requests.post(url, headers=post_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        # La respuesta contiene el índice de la fila insertada
        logger.info(f"Datos agregados a tabla '{tabla_id}'.")
        return response.json()
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en agregar_datos_tabla_excel: {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en agregar_datos_tabla_excel: {e}", exc_info=True)
        raise
