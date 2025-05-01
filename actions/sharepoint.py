import logging
import requests
import json
import os
from typing import Dict, List, Optional, Union, Any

# Usar el logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar constantes globales desde shared/constants.py
try:
    from shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    # Fallback
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_TIMEOUT = 45
    logger.warning("No se pudo importar constantes desde shared (SharePoint), usando defaults.")

# Leer variables de entorno específicas de SharePoint (si existen)
SHAREPOINT_DEFAULT_SITE_ID = os.environ.get('SHAREPOINT_DEFAULT_SITE_ID')
SHAREPOINT_DEFAULT_DRIVE_ID = os.environ.get('SHAREPOINT_DEFAULT_DRIVE_ID', 'Documents')


# --- Helper para obtener Site ID (Refactorizado - Sin Cache Global) ---
def _obtener_site_id_sp(headers: Dict[str, str], site_id_o_hostname_path: Optional[str] = None) -> str:
    """
    Obtiene el ID de un sitio de SharePoint.
    Prioridad: ID específico (si contiene ','), hostname:/ruta/relativa, SHAREPOINT_DEFAULT_SITE_ID, lookup raíz.
    Requiere headers autenticados.
    """
    # 1. Si es un ID compuesto (hostname,spsite-guid,spweb-guid)
    if site_id_o_hostname_path and ',' in site_id_o_hostname_path:
        logger.debug(f"Usando Site ID provisto directamente: {site_id_o_hostname_path}")
        return site_id_o_hostname_path

    # 2. Si es un path tipo hostname:/ruta/relativa
    if site_id_o_hostname_path and ':' in site_id_o_hostname_path and not site_id_o_hostname_path.startswith('/'):
        site_path_lookup = site_id_o_hostname_path
        url = f"{BASE_URL}/sites/{site_path_lookup}?$select=id"
        logger.debug(f"Buscando Site ID por path: GET {url}")
        try:
            response = None # Reinicializar antes de usarla para la llamada a /sites/root
            response.raise_for_status()
            site_data = response.json()
            site_id = site_data.get('id')
            if not site_id: raise ValueError("Respuesta inválida, falta 'id'.")
            logger.info(f"Site ID obtenido por path '{site_path_lookup}': {site_id}")
            return site_id
        except requests.exceptions.RequestException as req_ex:
             # Si falla el lookup por path (ej: 404), intentar con default o raíz
             logger.warning(f"No se encontró sitio por path '{site_path_lookup}': {req_ex}. Intentando default/raíz.")
        except Exception as e:
             logger.warning(f"Error inesperado buscando sitio por path '{site_path_lookup}': {e}. Intentando default/raíz.")

    # 3. Si se configuró un default en las variables de entorno
    if SHAREPOINT_DEFAULT_SITE_ID:
        logger.debug(f"Usando Site ID default de configuración: {SHAREPOINT_DEFAULT_SITE_ID}")
        return SHAREPOINT_DEFAULT_SITE_ID

    # 4. Si no hay nada más, intentar obtener el sitio raíz
    url = f"{BASE_URL}/sites/root?$select=id"
    logger.debug(f"Obteniendo sitio raíz SP: GET {url}")
    response: Optional[requests.Response] = None
    try:
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        site_data = response.json(); site_id = site_data.get('id')
        if not site_id: raise ValueError("Respuesta de sitio raíz inválida.")
        logger.info(f"Site ID raíz obtenido: {site_id}"); return site_id
    except requests.exceptions.RequestException as req_ex: logger.error(f"Error Request obteniendo Site ID raíz SP: {req_ex}", exc_info=True); raise Exception(f"Error API obteniendo Site ID raíz: {req_ex}")
    except Exception as e: logger.error(f"Error inesperado obteniendo Site ID raíz SP: {e}", exc_info=True); raise

# --- Helpers para Endpoints (Refactorizados) ---
def _get_sp_drive_endpoint(headers: Dict[str, str], site_id: str, drive_id_or_name: Optional[str] = None) -> str:
    target_drive = drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID
    return f"{BASE_URL}/sites/{site_id}/drives/{target_drive}"

def _get_sp_item_path_endpoint(headers: Dict[str, str], site_id: str, item_path: str, drive_id_or_name: Optional[str] = None) -> str:
    drive_endpoint = _get_sp_drive_endpoint(headers, site_id, drive_id_or_name)
    safe_path = item_path.strip()
    if not safe_path.startswith('/'): safe_path = '/' + safe_path
    return f"{drive_endpoint}/root" if safe_path == '/' else f"{drive_endpoint}/root:{safe_path}"

def _get_drive_id(headers: Dict[str, str], site_id: str, drive_id_or_name: Optional[str] = None) -> str:
    drive_endpoint = _get_sp_drive_endpoint(headers, site_id, drive_id_or_name)
    url = f"{drive_endpoint}?$select=id"; response: Optional[requests.Response] = None
    try:
        logger.debug(f"Obteniendo ID real del drive: GET {url}")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); drive_data = response.json(); actual_drive_id = drive_data.get('id')
        if not actual_drive_id: raise ValueError(f"No se pudo obtener 'id' del drive en {drive_endpoint}")
        logger.debug(f"Drive ID obtenido: {actual_drive_id}"); return actual_drive_id
    except requests.exceptions.RequestException as e: logger.error(f"Error Request obteniendo Drive ID para '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}': {e}", exc_info=True); raise Exception(f"Error API obteniendo Drive ID: {e}")
    except Exception as e: logger.error(f"Error inesperado obteniendo Drive ID: {e}", exc_info=True); raise

# ---- FUNCIONES DE LISTAS (Refactorizadas) ----
def crear_lista(headers: Dict[str, str], nombre_lista: str, site_id: Optional[str] = None) -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists"
    body = {"displayName": nombre_lista, "columns": [{"name": "Clave", "text": {}}, {"name": "Valor", "text": {}}], "list": {"template": "genericList"}}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando lista SP '{nombre_lista}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data=response.json(); logger.info(f"Lista SP '{nombre_lista}' creada."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en crear_lista (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en crear_lista (SP): {e}", exc_info=True); raise

def listar_listas(headers: Dict[str, str], site_id: Optional[str] = None) -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists"; params = {'$select': 'id,name,displayName,webUrl'}; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Listando listas SP sitio '{target_site_id}')")
        response = requests.get(url, headers=headers, params=params, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data=response.json(); logger.info(f"Listadas {len(data.get('value',[]))} listas SP."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en listar_listas (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en listar_listas (SP): {e}", exc_info=True); raise

def agregar_elemento(headers: Dict[str, str], nombre_lista: str, clave: str, valor: str, site_id: Optional[str] = None) -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{nombre_lista}/items"; body = {"fields": {"Clave": clave, "Valor": valor}}; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Agregando elemento SP a lista '{nombre_lista}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data=response.json(); logger.info(f"Elemento SP agregado a lista '{nombre_lista}'. ID: {data.get('id')}"); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en agregar_elemento (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en agregar_elemento (SP): {e}", exc_info=True); raise

def listar_elementos(headers: Dict[str, str], nombre_lista: str, site_id: Optional[str] = None, expand_fields: bool = True, top: int = 100, filter_query: Optional[str]=None) -> dict:
    # Añadido filter_query como opción
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url_base = f"{BASE_URL}/sites/{target_site_id}/lists/{nombre_lista}/items";
    params: Dict[str, Any] = {'$top': min(int(top), 999)}
    if expand_fields: params['$expand'] = 'fields'
    if filter_query: params['$filter'] = filter_query # Añadir filtro si se provee
    all_items = []; current_url: Optional[str] = url_base; current_headers = headers.copy(); response: Optional[requests.Response] = None
    try:
        page_count = 0
        while current_url:
            page_count += 1; logger.info(f"API Call: GET {current_url} Page: {page_count} (Listando elementos SP lista '{nombre_lista}')")
            current_params = params if page_count == 1 else None
            response = requests.get(current_url, headers=current_headers, params=current_params, timeout=GRAPH_API_TIMEOUT)
            response.raise_for_status(); data = response.json(); page_items = data.get('value', []); all_items.extend(page_items)
            current_url = data.get('@odata.nextLink')
        logger.info(f"Total elementos SP lista '{nombre_lista}': {len(all_items)}"); return {'value': all_items}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en listar_elementos (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en listar_elementos (SP): {e}", exc_info=True); raise

def actualizar_elemento(headers: Dict[str, str], nombre_lista: str, item_id: str, nuevos_valores: dict, site_id: Optional[str] = None) -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{nombre_lista}/items/{item_id}/fields"; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Actualizando elemento SP '{item_id}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        etag = nuevos_valores.pop('@odata.etag', None)
        if etag: current_headers['If-Match'] = etag
        response = requests.patch(url, headers=current_headers, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Elemento SP '{item_id}' actualizado."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en actualizar_elemento (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en actualizar_elemento (SP): {e}", exc_info=True); raise

def eliminar_elemento(headers: Dict[str, str], nombre_lista: str, item_id: str, site_id: Optional[str] = None) -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{nombre_lista}/items/{item_id}"; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando elemento SP '{item_id}')")
        # Podría necesitar If-Match header con ETag
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); logger.info(f"Elemento SP '{item_id}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en eliminar_elemento (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en eliminar_elemento (SP): {e}", exc_info=True); raise

# ---- FUNCIONES DE DOCUMENTOS (Bibliotecas / Drives) ----
def listar_documentos_biblioteca(headers: Dict[str, str], biblioteca: Optional[str] = None, site_id: Optional[str] = None, top: int = 100, ruta_carpeta: str = '/') -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, ruta_carpeta, target_drive)
    url_base = f"{item_endpoint}/children"; params = {'$top': min(int(top), 999)};
    all_files = []; current_url: Optional[str] = url_base; current_headers = headers.copy(); response: Optional[requests.Response] = None
    try:
        page_count = 0
        while current_url:
            page_count += 1; logger.info(f"API Call: GET {current_url} Page: {page_count} (Listando docs SP biblioteca '{target_drive}', ruta '{ruta_carpeta}')")
            current_params = params if page_count == 1 else None
            response = requests.get(current_url, headers=current_headers, params=current_params, timeout=GRAPH_API_TIMEOUT)
            response.raise_for_status(); data = response.json(); page_items = data.get('value', []); all_files.extend(page_items)
            current_url = data.get('@odata.nextLink')
        logger.info(f"Total docs SP biblioteca '{target_drive}', ruta '{ruta_carpeta}': {len(all_files)}"); return {'value': all_files}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en listar_documentos_biblioteca (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en listar_documentos_biblioteca (SP): {e}", exc_info=True); raise

def subir_documento(headers: Dict[str, str], nombre_archivo: str, contenido_bytes: bytes, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta_destino: str = '/', conflict_behavior: str = "rename") -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    # Normalizar path destino
    target_folder_path = ruta_carpeta_destino.strip('/')
    target_file_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, target_file_path, target_drive)
    url = f"{item_endpoint}:/content?@microsoft.graph.conflictBehavior={conflict_behavior}"; upload_headers = headers.copy(); upload_headers['Content-Type'] = 'application/octet-stream'; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PUT {item_endpoint}:/content (Subiendo doc SP '{nombre_archivo}' a '{ruta_carpeta_destino}')")
        if len(contenido_bytes) > 4*1024*1024: logger.warning(f"Archivo SP '{nombre_archivo}' > 4MB.")
        response = requests.put(url, headers=upload_headers, data=contenido_bytes, timeout=GRAPH_API_TIMEOUT * 3)
        response.raise_for_status(); data = response.json(); logger.info(f"Doc SP '{nombre_archivo}' subido. ID: {data.get('id')}"); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en subir_documento (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en subir_documento (SP): {e}", exc_info=True); raise

def eliminar_archivo(headers: Dict[str, str], nombre_archivo: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, item_path, target_drive)
    url = item_endpoint; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando archivo/carpeta SP '{item_path}')")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); logger.info(f"Archivo/Carpeta SP '{item_path}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en eliminar_archivo (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en eliminar_archivo (SP): {e}", exc_info=True); raise

# ---- FUNCIONES AVANZADAS DE ARCHIVOS (Refactorizadas) ----
def crear_carpeta_biblioteca(headers: Dict[str, str], nombre_carpeta: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta_padre: str = '/', conflict_behavior: str = "rename") -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    parent_folder_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, ruta_carpeta_padre, target_drive)
    url = f"{parent_folder_endpoint}/children"; body = {"name": nombre_carpeta, "folder": {}, "@microsoft.graph.conflictBehavior": conflict_behavior}; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando carpeta SP '{nombre_carpeta}' en '{ruta_carpeta_padre}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Carpeta SP '{nombre_carpeta}' creada. ID: {data.get('id')}"); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en crear_carpeta_biblioteca (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en crear_carpeta_biblioteca (SP): {e}", exc_info=True); raise

def mover_archivo(headers: Dict[str, str], nombre_archivo: str, nueva_ubicacion: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta_origen: str = '/', nuevo_nombre: Optional[str]=None) -> dict:
    # Añadido nuevo_nombre opcional
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path_origen = ruta_carpeta_origen.strip('/')
    item_path_origen = f"/{nombre_archivo}" if not target_folder_path_origen else f"/{target_folder_path_origen}/{nombre_archivo}"
    item_endpoint_origen = _get_sp_item_path_endpoint(headers, target_site_id, item_path_origen, target_drive)
    url = item_endpoint_origen
    try: actual_drive_id = _get_drive_id(headers, target_site_id, target_drive)
    except Exception as drive_err: raise Exception(f"Error obteniendo Drive ID para mover: {drive_err}")
    parent_dest_path = nueva_ubicacion.strip()
    if not parent_dest_path.startswith('/'): parent_dest_path = '/' + parent_dest_path
    parent_path = f"/drives/{actual_drive_id}/root" if parent_dest_path == '/' else f"/drives/{actual_drive_id}/root:{parent_dest_path}"
    body: Dict[str, Any] = {"parentReference": {"path": parent_path}, "name": nombre_archivo }
    # Usar nuevo nombre si se proporciona, si no, mantener el original
    body["name"] = nuevo_nombre if nuevo_nombre is not None else nombre_archivo
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Moviendo SP '{item_path_origen}' a '{nueva_ubicacion}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        response = requests.patch(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Archivo/Carpeta SP '{nombre_archivo}' movido a '{nueva_ubicacion}'."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en mover_archivo (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en mover_archivo (SP): {e}", exc_info=True); raise

def copiar_archivo(headers: Dict[str, str], nombre_archivo: str, nueva_ubicacion: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta_origen: str = '/', nuevo_nombre_copia: Optional[str] = None) -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path_origen = ruta_carpeta_origen.strip('/')
    item_path_origen = f"/{nombre_archivo}" if not target_folder_path_origen else f"/{target_folder_path_origen}/{nombre_archivo}"
    item_endpoint_origen = _get_sp_item_path_endpoint(headers, target_site_id, item_path_origen, target_drive)
    url = f"{item_endpoint_origen}/copy"
    try: actual_drive_id = _get_drive_id(headers, target_site_id, target_drive)
    except Exception as drive_err: raise Exception(f"Error obteniendo Drive ID para copiar: {drive_err}")
    parent_dest_path = nueva_ubicacion.strip()
    if not parent_dest_path.startswith('/'): parent_dest_path = '/' + parent_dest_path
    parent_path = "/drive/root" if parent_dest_path == '/' else f"/drive/root:{parent_dest_path}"
    body = {"parentReference": {"driveId": actual_drive_id, "path": parent_path}, "name": nuevo_nombre_copia or f"Copia de {nombre_archivo}" }; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Iniciando copia SP '{item_path_origen}' a '{nueva_ubicacion}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); monitor_url = response.headers.get('Location'); logger.info(f"Copia SP '{nombre_archivo}' iniciada. Monitor: {monitor_url}"); return {"status": "Copia Iniciada", "code": response.status_code, "monitorUrl": monitor_url}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en copiar_archivo (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en copiar_archivo (SP): {e}", exc_info=True); raise

def obtener_metadatos_archivo(headers: Dict[str, str], nombre_archivo: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, item_path, target_drive)
    url = item_endpoint; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo metadatos SP '{item_path}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Metadatos SP '{item_path}' obtenidos."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en obtener_metadatos_archivo (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en obtener_metadatos_archivo (SP): {e}", exc_info=True); raise

def actualizar_metadatos_archivo(headers: Dict[str, str], nombre_archivo: str, nuevos_valores: dict, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, item_path, target_drive)
    url = item_endpoint; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Actualizando metadatos SP '{item_path}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        etag = nuevos_valores.pop('@odata.etag', None)
        if etag: current_headers['If-Match'] = etag
        response = requests.patch(url, headers=current_headers, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Metadatos SP '{item_path}' actualizados."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en actualizar_metadatos_archivo (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en actualizar_metadatos_archivo (SP): {e}", exc_info=True); raise

def obtener_contenido_archivo(headers: Dict[str, str], nombre_archivo: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> bytes:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, item_path, target_drive)
    url = f"{item_endpoint}/content"; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo contenido SP '{item_path}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT * 2)
        response.raise_for_status(); logger.info(f"Contenido SP '{item_path}' obtenido."); return response.content
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en obtener_contenido_archivo (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en obtener_contenido_archivo (SP): {e}", exc_info=True); raise

def actualizar_contenido_archivo(headers: Dict[str, str], nombre_archivo: str, nuevo_contenido: bytes, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, item_path, target_drive)
    url = f"{item_endpoint}/content"; upload_headers = headers.copy(); upload_headers['Content-Type'] = 'application/octet-stream'; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PUT {url} (Actualizando contenido SP '{item_path}')")
        if len(nuevo_contenido) > 4*1024*1024: logger.warning(f"Archivo SP '{nombre_archivo}' > 4MB.")
        response = requests.put(url, headers=upload_headers, data=nuevo_contenido, timeout=GRAPH_API_TIMEOUT * 3)
        response.raise_for_status(); data = response.json(); logger.info(f"Contenido SP '{item_path}' actualizado."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en actualizar_contenido_archivo (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en actualizar_contenido_archivo (SP): {e}", exc_info=True); raise

def crear_enlace_compartido_archivo(headers: Dict[str, str], nombre_archivo: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/', tipo_enlace: str = "view", alcance: str = "organization") -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, item_path, target_drive)
    url = f"{item_endpoint}/createLink"; body = {"type": tipo_enlace, "scope": alcance}; response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando enlace SP '{item_path}')")
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status(); data = response.json(); logger.info(f"Enlace SP creado para '{item_path}'."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en crear_enlace_compartido_archivo (SP): {e}", exc_info=True); raise
    except Exception as e: logger.error(f"Error inesperado en crear_enlace_compartido_archivo (SP): {e}", exc_info=True); raise
