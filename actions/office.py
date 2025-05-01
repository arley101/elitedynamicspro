# actions/office.py (Refactorizado v2 con Helper)

import logging
import requests # Solo para tipos de excepción
import json
import os
from typing import Dict, List, Optional, Union, Any

# Usar logger principal
logger = logging.getLogger("azure.functions")

# Importar helper y constantes
try:
    from helpers.http_client import hacer_llamada_api
    from shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    logger.error("Error importando helpers/constantes en Office.")
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs): raise NotImplementedError("Helper no importado")

# ---- WORD ONLINE (via OneDrive /me/drive) ----
def crear_documento_word(headers: Dict[str, str], nombre_archivo: str, ruta: str = "/") -> dict:
    if not nombre_archivo.lower().endswith(".docx"): nombre_archivo += ".docx"
    target_folder_path = ruta.strip('/')
    target_file_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    url = f"{BASE_URL}/me/drive/root:{target_file_path}"
    create_headers = headers.copy(); create_headers.setdefault('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    body = {"name": nombre_archivo, "file": {}}
    logger.info(f"Creando Word '{nombre_archivo}' en ruta '{ruta}'")
    # Usamos PUT pero el helper espera JSON por defecto, forzamos expect_json=True
    return hacer_llamada_api("PUT", url, create_headers, json_data=body, expect_json=True)

def insertar_texto_word(headers: Dict[str, str], item_id: str, texto: str) -> dict:
    """REEMPLAZA contenido con texto plano."""
    url = f"{BASE_URL}/me/drive/items/{item_id}/content"
    update_headers = headers.copy(); update_headers['Content-Type'] = 'text/plain';
    logger.warning(f"REEMPLAZANDO contenido Word ID '{item_id}' con texto plano")
    # Usamos PUT con data binaria, el helper puede manejarlo si se pasa 'data'
    return hacer_llamada_api("PUT", url, update_headers, data=texto.encode('utf-8'), timeout=GRAPH_API_TIMEOUT * 2)

def obtener_documento_word(headers: Dict[str, str], item_id: str) -> bytes:
    """Obtiene contenido binario (.docx)."""
    url = f"{BASE_URL}/me/drive/items/{item_id}/content"
    logger.info(f"Obteniendo contenido Word ID '{item_id}'")
    # Necesitamos la respuesta cruda para obtener bytes, llamar a requests directo
    try:
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT * 2)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en obtener_documento_word: {e}", exc_info=True); raise Exception(f"Error API obteniendo doc Word: {e}")
    except Exception as e: logger.error(f"Error inesperado en obtener_documento_word: {e}", exc_info=True); raise

# ---- EXCEL ONLINE (via OneDrive /me/drive) ----
def crear_excel(headers: Dict[str, str], nombre_archivo: str, ruta: str = "/") -> dict:
    if not nombre_archivo.lower().endswith(".xlsx"): nombre_archivo += ".xlsx"
    target_folder_path = ruta.strip('/')
    target_file_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    url = f"{BASE_URL}/me/drive/root:{target_file_path}"
    create_headers = headers.copy(); create_headers.setdefault('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    body = {"name": nombre_archivo, "file": {}}
    logger.info(f"Creando Excel '{nombre_archivo}' en ruta '{ruta}'")
    return hacer_llamada_api("PUT", url, create_headers, json_data=body, expect_json=True)

def escribir_celda_excel(headers: Dict[str, str], item_id: str, hoja: str, celda: str, valor: Union[str, int, float, bool]) -> dict:
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/range(address='{celda}')"
    body = {"values": [[valor]]}
    logger.info(f"Escribiendo en celda '{celda}' hoja '{hoja}' item '{item_id}'")
    return hacer_llamada_api("PATCH", url, headers, json_data=body)

def leer_celda_excel(headers: Dict[str, str], item_id: str, hoja: str, celda: str) -> dict:
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/range(address='{celda}')?$select=text,values,address"
    logger.info(f"Leyendo celda '{celda}' hoja '{hoja}' item '{item_id}'")
    return hacer_llamada_api("GET", url, headers)

def crear_tabla_excel(headers: Dict[str, str], item_id: str, hoja: str, rango: str, tiene_headers: bool = False) -> dict:
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/tables/add"
    body = {"address": f"{hoja}!{rango}", "hasHeaders": tiene_headers}
    logger.info(f"Creando tabla en rango '{rango}' hoja '{hoja}' item '{item_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def agregar_datos_tabla_excel(headers: Dict[str, str], item_id: str, tabla_id_o_nombre: str, valores: List[List[Any]]) -> dict:
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/tables/{tabla_id_o_nombre}/rows"
    body = {"values": valores}
    logger.info(f"Agregando {len(valores)} filas a tabla '{tabla_id_o_nombre}' item '{item_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)
