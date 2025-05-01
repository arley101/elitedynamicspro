# actions/onedrive.py (Refactorizado y Corregido - Final)

import logging
import requests
import os
import json
# Corregido: Añadir Any
from typing import Dict, Optional, Union, List, Any

# Usar el logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar constantes globales desde shared/constants.py
try:
    from shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    # Fallback
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_TIMEOUT = 45
    logger.warning("No se pudo importar constantes desde shared (OneDrive), usando defaults.")


# ---- Helpers Locales (Usan BASE_URL global) ----
def _get_od_me_drive_endpoint() -> str:
    return f"{BASE_URL}/me/drive"

def _get_od_me_item_path_endpoint(ruta_relativa: str) -> str:
    drive_endpoint = _get_od_me_drive_endpoint()
    safe_path = ruta_relativa.strip()
    if safe_path and not safe_path.startswith('/'): safe_path = '/' + safe_path
    return f"{drive_endpoint}/root" if not safe_path or safe_path == '/' else f"{drive_endpoint}/root:{safe_path}"

# ---- FUNCIONES ONEDRIVE (Refactorizadas) ----
def listar_archivos(headers: Dict[str, str], ruta: str = "/", top: int = 100) -> dict:
    item_endpoint = _get_od_me_item_path_endpoint(ruta)
    url = f"{item_endpoint}/children"; params: Dict[str, Any] = {'$top': min(int(top), 999)}; all_items: List[Dict[str, Any]] = []; current_url: Optional[str] = url; current_headers = headers.copy(); response: Optional[requests.Response] = None
    try:
        page_count = 0
        while current_url:
            page_count += 1; logger.info(f"API Call: GET {current_url} Page: {page_count} (Listando OD /me ruta '{ruta}')")
            current_params = params if page_count == 1 else None
            # Corregido: Añadir assert para current_url
            assert current_url is not None
            response = requests.get(current_url, headers=current_headers, params=current_params, timeout=GRAPH_API_TIMEOUT)
            response.raise_for_status(); data = response.json(); page_items = data.get('value', []); all_items.extend(page_items)
            current_url = data.get('@odata.nextLink')
        logger.info(f"Total items OD /me en '{ruta}': {len(all_items)}"); return {'value': all_items}
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en od_listar_archivos: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en od_listar_archivos (/me): {e}", exc_info=True); raise

def subir_archivo(headers: Dict[str, str], nombre_archivo: str, contenido_bytes: bytes, ruta: str = "/", conflict_behavior: str = "rename") -> dict:
    target_folder_path = ruta.strip('/'); target_file_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"; item_endpoint = _get_od_me_item_path_endpoint(target_file_path); url = f"{item_endpoint}:/content?@microsoft.graph.conflictBehavior={conflict_behavior}"; upload_headers = headers.copy(); upload_headers['Content-Type'] = 'application/octet-stream'; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PUT {item_endpoint}:/content (Subiendo OD /me '{nombre_archivo}' a ruta '{ruta}')")
        if len(contenido_bytes) > 4*1024*1024: logger.warning(f"Archivo OD '{nombre_archivo}' > 4MB.")
        response = requests.put(url, headers=upload_headers, data=contenido_bytes, timeout=GRAPH_API_TIMEOUT * 3)
        response.raise_for_status(); data = response.json(); logger.info(f"Archivo OD '{nombre_archivo}' subido."); return data
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en od_subir_archivo: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en od_subir_archivo (/me): {e}", exc_info=True); raise

def descargar_archivo(headers: Dict[str, str], nombre_archivo: str, ruta: str = "/") -> bytes:
    target_folder_path = ruta.strip('/'); target_file_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"; item_endpoint = _get_od_me_item_path_endpoint(target_file_path); url = f"{item_endpoint}/content"; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Descargando OD /me '{nombre_archivo}' de ruta '{ruta}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT * 2)
        response.raise_for_status(); logger.info(f"Archivo OD '{nombre_archivo}' descargado."); return response.content
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en od_descargar_archivo: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en od_descargar_archivo (/me): {e}", exc_info=True); raise

def eliminar_archivo(headers: Dict[str, str], nombre_archivo: str, ruta: str = "/") -> dict:
    target_folder_path = ruta.strip('/'); target_file_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"; item_endpoint = _get_od_me_item_path_endpoint(target_file_path); url = item_endpoint; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando OD /me '{nombre_archivo}' de ruta '{ruta}')")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); logger.info(f"Archivo/Carpeta OD '{nombre_archivo}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en od_eliminar_archivo: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en od_eliminar_archivo (/me): {e}", exc_info=True); raise

def crear_carpeta(headers: Dict[str, str], nombre_carpeta: str, ruta: str = "/", conflict_behavior: str = "rename") -> dict:
    parent_folder_endpoint = _get_od_me_item_path_endpoint(ruta); url = f"{parent_folder_endpoint}/children"; body: Dict[str, Any] = {"name": nombre_carpeta, "folder": {}, "@microsoft.graph.conflictBehavior": conflict_behavior}; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando OD /me carpeta '{nombre_carpeta}' en ruta '{ruta}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Carpeta OD '{nombre_carpeta}' creada. ID: {data.get('id')}"); return data
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en od_crear_carpeta: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en od_crear_carpeta (/me): {e}", exc_info=True); raise

def mover_archivo(headers: Dict[str, str], nombre_archivo: str, ruta_origen: str = "/", ruta_destino: str = "/NuevaCarpeta", nuevo_nombre: Optional[str] = None) -> dict:
    target_folder_path_origen = ruta_origen.strip('/'); item_path_origen = f"/{nombre_archivo}" if not target_folder_path_origen else f"/{target_folder_path_origen}/{nombre_archivo}"; item_origen_endpoint = _get_od_me_item_path_endpoint(item_path_origen); url = item_origen_endpoint; parent_dest_path = ruta_destino.strip();
    if not parent_dest_path.startswith('/'): parent_dest_path = '/' + parent_dest_path
    parent_reference_path = "/drive/root" if parent_dest_path == '/' else f"/drive/root:{parent_dest_path}"; body: Dict[str, Any] = { "parentReference": { "path": parent_reference_path } }; body["name"] = nuevo_nombre if nuevo_nombre is not None else nombre_archivo; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Moviendo OD /me '{nombre_archivo}' de '{ruta_origen}' a '{ruta_destino}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        response = requests.patch(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Archivo/Carpeta OD '{nombre_archivo}' movido a '{ruta_destino}'."); return data
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en od_mover_archivo: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en od_mover_archivo (/me): {e}", exc_info=True); raise

def copiar_archivo(headers: Dict[str, str], nombre_archivo: str, ruta_origen: str = "/", ruta_destino: str = "/Copias", nuevo_nombre_copia: Optional[str] = None) -> dict:
    drive_endpoint = _get_od_me_drive_endpoint();
    try: drive_resp = requests.get(drive_endpoint, headers=headers, params={'$select':'id'}, timeout=GRAPH_API_TIMEOUT); drive_resp.raise_for_status(); actual_drive_id = drive_resp.json().get('id'); assert actual_drive_id is not None
    except Exception as drive_err: logger.error(f"Error obteniendo ID drive /me para copiar: {drive_err}", exc_info=True); raise Exception(f"Error obteniendo ID drive /me para copia: {drive_err}")
    target_folder_path_origen = ruta_origen.strip('/'); item_path_origen = f"/{nombre_archivo}" if not target_folder_path_origen else f"/{target_folder_path_origen}/{nombre_archivo}"; item_origen_endpoint = _get_od_me_item_path_endpoint(item_path_origen); url = f"{item_origen_endpoint}/copy"; parent_dest_path = ruta_destino.strip();
    if not parent_dest_path.startswith('/'): parent_dest_path = '/' + parent_dest_path
    parent_reference_path = "/drive/root" if parent_dest_path == '/' else f"/drive/root:{parent_dest_path}"; body: Dict[str, Any] = {"parentReference": { "driveId": actual_drive_id, "path": parent_reference_path }}; body["name"] = nuevo_nombre_copia if nuevo_nombre_copia is not None else f"Copia de {nombre_archivo}"; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Copiando OD /me '{nombre_archivo}' de '{ruta_origen}' a '{ruta_destino}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); monitor_url = response.headers.get('Location'); logger.info(f"Copia OD '{nombre_archivo}' iniciada. Monitor: {monitor_url}"); return {"status": "Copia Iniciada", "code": response.status_code, "monitorUrl": monitor_url}
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en od_copiar_archivo: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en od_copiar_archivo (/me): {e}", exc_info=True); raise

def obtener_metadatos_archivo(headers: Dict[str, str], nombre_archivo: str, ruta: str = "/") -> dict:
    target_folder_path = ruta.strip('/'); item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"; item_endpoint = _get_od_me_item_path_endpoint(item_path); url = item_endpoint; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo metadatos OD /me '{nombre_archivo}' ruta '{ruta}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Metadatos OD '{nombre_archivo}' obtenidos."); return data
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en od_obtener_metadatos_archivo: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en od_obtener_metadatos_archivo (/me): {e}", exc_info=True); raise

def actualizar_metadatos_archivo(headers: Dict[str, str], nombre_archivo: str, nuevos_valores: dict, ruta:str = "/") -> dict:
    target_folder_path = ruta.strip('/'); item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"; item_endpoint = _get_od_me_item_path_endpoint(item_path); url = item_endpoint; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Actualizando metadatos OD /me '{nombre_archivo}' ruta '{ruta}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        etag = nuevos_valores.pop('@odata.etag', None)
        if etag: current_headers['If-Match'] = etag
        response = requests.patch(url, headers=current_headers, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Metadatos OD '{nombre_archivo}' actualizados."); return data
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request en od_actualizar_metadatos_archivo: {req_ex}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en od_actualizar_metadatos_archivo (/me): {e}", exc_info=True); raise
