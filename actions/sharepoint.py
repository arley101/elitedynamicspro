# actions/sharepoint.py (Refactorizado v2 con Helper y Memoria)

import logging
import requests # Necesario aquí solo para tipos de excepción
import os
import json # Para formateo de exportación
import csv # Para exportación CSV
from io import StringIO # Para exportación CSV
from typing import Dict, List, Optional, Any, Union

# Importar helper y constantes
try:
    from helpers.http_client import hacer_llamada_api
    from shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    # Fallback si la estructura no es reconocida
    logger = logging.getLogger("azure.functions")
    logger.error("Error importando helpers/constantes en SharePoint. Usando mocks/defaults.")
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs): raise NotImplementedError("Helper no importado")

# Usar logger
logger = logging.getLogger("azure.functions")

# Leer variables de entorno específicas de SharePoint (si existen)
SHAREPOINT_DEFAULT_SITE_ID = os.environ.get('SHAREPOINT_DEFAULT_SITE_ID')
SHAREPOINT_DEFAULT_DRIVE_ID = os.environ.get('SHAREPOINT_DEFAULT_DRIVE_ID', 'Documents')
# Nombre de la lista para memoria persistente (configurable)
MEMORIA_LIST_NAME = os.environ.get('SHAREPOINT_MEMORY_LIST', 'MemoriaPersistenteAsistente')

# --- Helper para obtener Site ID (Usa http helper) ---
def _obtener_site_id_sp(headers: Dict[str, str], site_id_input: Optional[str] = None) -> str:
    """Obtiene el ID de un sitio SharePoint. Prioridad: input, default env var, raíz."""
    if site_id_input and ',' in site_id_input: return site_id_input
    if site_id_input and ':' in site_id_input:
        site_path_lookup = site_id_input
        url = f"{BASE_URL}/sites/{site_path_lookup}?$select=id"
        try:
            logger.debug(f"Buscando Site ID por path: GET {url}")
            site_data = hacer_llamada_api("GET", url, headers)
            site_id = site_data.get("id")
            if site_id: logger.info(f"Site ID por path '{site_path_lookup}': {site_id}"); return site_id
            else: raise ValueError("Respuesta inválida, falta 'id'.")
        except Exception as e: logger.warning(f"No se encontró sitio por path '{site_path_lookup}' o error: {e}. Intentando default/raíz.")
    if SHAREPOINT_DEFAULT_SITE_ID: return SHAREPOINT_DEFAULT_SITE_ID
    url = f"{BASE_URL}/sites/root?$select=id"
    try:
        logger.debug(f"Obteniendo sitio raíz SP: GET {url}")
        site_data = hacer_llamada_api("GET", url, headers)
        site_id = site_data.get("id")
        if not site_id: raise ValueError("Respuesta de sitio raíz inválida.")
        logger.info(f"Site ID raíz obtenido: {site_id}"); return site_id
    except Exception as e: logger.error(f"Fallo crítico al obtener Site ID: {e}", exc_info=True); raise

# --- Helpers para Endpoints (Simplificados) ---
def _get_sp_drive_endpoint(site_id: str, drive_id_or_name: Optional[str] = None) -> str:
    target_drive = drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID
    return f"{BASE_URL}/sites/{site_id}/drives/{target_drive}"

def _get_sp_item_path_endpoint(site_id: str, item_path: str, drive_id_or_name: Optional[str] = None) -> str:
    drive_endpoint = _get_sp_drive_endpoint(site_id, drive_id_or_name) # No necesita headers aquí
    safe_path = item_path.strip()
    if not safe_path.startswith('/'): safe_path = '/' + safe_path
    return f"{drive_endpoint}/root" if safe_path == '/' else f"{drive_endpoint}/root:{safe_path}"

def _get_drive_id(headers: Dict[str, str], site_id: str, drive_id_or_name: Optional[str] = None) -> str:
    drive_endpoint = _get_sp_drive_endpoint(site_id, drive_id_or_name) # No necesita headers aquí
    url = f"{drive_endpoint}?$select=id"
    try:
        drive_data = hacer_llamada_api("GET", url, headers)
        actual_drive_id = drive_data.get('id')
        if not actual_drive_id: raise ValueError("No se pudo obtener 'id' del drive.")
        return actual_drive_id
    except Exception as e: raise Exception(f"Error API obteniendo Drive ID para '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}': {e}")

# ---- FUNCIONES DE LISTAS ----
# Usan el helper hacer_llamada_api

def crear_lista(headers: Dict[str, str], nombre_lista: str, columnas: Optional[List[Dict[str, Any]]] = None, site_id: Optional[str] = None) -> dict:
    """Crea una nueva lista en SharePoint con columnas personalizadas."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists"
    default_columnas = [{"name": "Title", "text": {}}] # Title siempre existe
    body = {
        "displayName": nombre_lista,
        "columns": default_columnas + (columnas if columnas else []),
        "list": {"template": "genericList"}
    }
    logger.info(f"Creando lista SP '{nombre_lista}' en sitio {target_site_id}")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def listar_listas(headers: Dict[str, str], site_id: Optional[str] = None, select: str = "id,name,displayName,webUrl") -> dict:
    """Lista las listas del sitio especificado."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists"
    params = {"$select": select} if select else None
    logger.info(f"Listando listas SP del sitio {target_site_id}")
    return hacer_llamada_api("GET", url, headers, params=params)

def agregar_elemento_lista(headers: Dict[str, str], lista_id_o_nombre: str, datos_campos: Dict[str, Any], site_id: Optional[str] = None) -> dict:
    """Agrega un elemento a una lista de SharePoint."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items"
    body = {"fields": datos_campos}
    logger.info(f"Agregando elemento a lista SP '{lista_id_o_nombre}' en sitio {target_site_id}")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def listar_elementos_lista(headers: Dict[str, str], lista_id_o_nombre: str, site_id: Optional[str] = None, expand_fields: bool = True, top: int = 100, filter_query: Optional[str]=None, select: Optional[str]=None, order_by: Optional[str]=None) -> dict:
    """Lista elementos de una lista, con paginación, filtro, selección y orden."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url_base = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items"
    params: Dict[str, Any] = {'$top': min(int(top), 999)}
    if expand_fields: params['$expand'] = 'fields'
    if filter_query: params['$filter'] = filter_query
    if select: params['$select'] = select
    if order_by: params['$orderby'] = order_by

    all_items = []
    current_url: Optional[str] = url_base
    current_headers = headers.copy()
    try:
        page_count = 0
        while current_url:
            page_count += 1
            logger.info(f"Listando elementos SP lista '{lista_id_o_nombre}', Página: {page_count}")
            current_params = params if page_count == 1 else None # Params solo en la primera
            # Necesitamos pasar la URL completa a hacer_llamada_api si es paginada
            # El helper no maneja paginación interna
            response = requests.get(current_url, headers=current_headers, params=current_params, timeout=GRAPH_API_TIMEOUT)
            response.raise_for_status()
            data = response.json()
            page_items = data.get('value', [])
            all_items.extend(page_items)
            current_url = data.get('@odata.nextLink') # Obtener siguiente link

        logger.info(f"Total elementos SP lista '{lista_id_o_nombre}': {len(all_items)}")
        return {'value': all_items}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en listar_elementos_lista (SP): {e}", exc_info=True); raise Exception(f"Error API listando elementos SP: {e}")
    except Exception as e: logger.error(f"Error inesperado en listar_elementos_lista (SP): {e}", exc_info=True); raise

def actualizar_elemento_lista(headers: Dict[str, str], lista_id_o_nombre: str, item_id: str, nuevos_valores_campos: dict, site_id: Optional[str] = None) -> dict:
    """Actualiza campos de un item de lista."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items/{item_id}/fields"
    # Añadir ETag para concurrencia si se incluye en nuevos_valores
    etag = nuevos_valores_campos.pop('@odata.etag', None)
    current_headers = headers.copy()
    if etag: current_headers['If-Match'] = etag
    logger.info(f"Actualizando elemento SP '{item_id}' en lista '{lista_id_o_nombre}'")
    return hacer_llamada_api("PATCH", url, current_headers, json_data=nuevos_valores_campos)

def eliminar_elemento_lista(headers: Dict[str, str], lista_id_o_nombre: str, item_id: str, site_id: Optional[str] = None, etag: Optional[str] = None) -> Optional[dict]:
    """Elimina un item de lista."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items/{item_id}"
    current_headers = headers.copy()
    if etag: current_headers['If-Match'] = etag
    else: logger.warning(f"Eliminando item SP {item_id} sin ETag.")
    logger.info(f"Eliminando elemento SP '{item_id}' de lista '{lista_id_o_nombre}'")
    # Hacer llamada devuelve None en caso de éxito 204
    hacer_llamada_api("DELETE", url, current_headers)
    return {"status": "Eliminado", "item_id": item_id} # Devolver confirmación

# ---- FUNCIONES DE DOCUMENTOS (Bibliotecas / Drives) ----
def listar_documentos_biblioteca(headers: Dict[str, str], biblioteca: Optional[str] = None, site_id: Optional[str] = None, top: int = 100, ruta_carpeta: str = '/') -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_carpeta, target_drive)
    url_base = f"{item_endpoint}/children"; params = {'$top': min(int(top), 999)};
    # Manejo de paginación (similar a listar_elementos_lista)
    all_files = []; current_url: Optional[str] = url_base; current_headers = headers.copy(); response: Optional[requests.Response] = None
    try:
        page_count = 0
        while current_url:
            page_count += 1; logger.info(f"Listando docs SP biblioteca '{target_drive}', Página: {page_count}")
            current_params = params if page_count == 1 else None
            response = requests.get(current_url, headers=current_headers, params=current_params, timeout=GRAPH_API_TIMEOUT)
            response.raise_for_status(); data = response.json(); page_items = data.get('value', []); all_files.extend(page_items)
            current_url = data.get('@odata.nextLink')
        logger.info(f"Total docs SP biblioteca '{target_drive}', ruta '{ruta_carpeta}': {len(all_files)}"); return {'value': all_files}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en listar_documentos_biblioteca (SP): {e}", exc_info=True); raise Exception(f"Error API listando documentos SP: {e}")
    except Exception as e: logger.error(f"Error inesperado en listar_documentos_biblioteca (SP): {e}", exc_info=True); raise

def subir_documento(headers: Dict[str, str], nombre_archivo: str, contenido_bytes: bytes, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta_destino: str = '/', conflict_behavior: str = "rename") -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta_destino.strip('/')
    target_file_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, target_file_path, target_drive)
    url = f"{item_endpoint}:/content"; params = {"@microsoft.graph.conflictBehavior": conflict_behavior}; upload_headers = headers.copy(); upload_headers['Content-Type'] = 'application/octet-stream';
    try:
        logger.info(f"Subiendo doc SP '{nombre_archivo}' a '{ruta_carpeta_destino}'")
        if len(contenido_bytes) > 4*1024*1024: logger.warning(f"Archivo SP '{nombre_archivo}' > 4MB.")
        # Hacer llamada API no maneja data binaria directamente, usar requests
        response = requests.put(url, headers=upload_headers, params=params, data=contenido_bytes, timeout=GRAPH_API_TIMEOUT * 3)
        response.raise_for_status(); data = response.json(); logger.info(f"Doc SP '{nombre_archivo}' subido. ID: {data.get('id')}"); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en subir_documento (SP): {e}", exc_info=True); raise Exception(f"Error API subiendo documento: {e}")
    except Exception as e: logger.error(f"Error inesperado en subir_documento (SP): {e}", exc_info=True); raise

def eliminar_archivo(headers: Dict[str, str], nombre_archivo: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> Optional[dict]:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = item_endpoint;
    logger.info(f"Eliminando archivo/carpeta SP '{item_path}'")
    hacer_llamada_api("DELETE", url, headers) # Devuelve None en éxito 204
    return {"status": "Eliminado", "path": item_path}

# --- FUNCIONES AVANZADAS DE ARCHIVOS ---
def crear_carpeta_biblioteca(headers: Dict[str, str], nombre_carpeta: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta_padre: str = '/', conflict_behavior: str = "rename") -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    parent_folder_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_carpeta_padre, target_drive)
    url = f"{parent_folder_endpoint}/children"; body = {"name": nombre_carpeta, "folder": {}, "@microsoft.graph.conflictBehavior": conflict_behavior};
    logger.info(f"Creando carpeta SP '{nombre_carpeta}' en '{ruta_carpeta_padre}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def mover_archivo(headers: Dict[str, str], nombre_archivo: str, nueva_ubicacion: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta_origen: str = '/', nuevo_nombre: Optional[str]=None) -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive_name = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path_origen = ruta_carpeta_origen.strip('/')
    item_path_origen = f"/{nombre_archivo}" if not target_folder_path_origen else f"/{target_folder_path_origen}/{nombre_archivo}"
    item_endpoint_origen = _get_sp_item_path_endpoint(target_site_id, item_path_origen, target_drive_name)
    url = item_endpoint_origen
    try: actual_drive_id = _get_drive_id(headers, target_site_id, target_drive_name)
    except Exception as drive_err: raise Exception(f"Error obteniendo Drive ID para mover: {drive_err}")
    parent_dest_path = nueva_ubicacion.strip();
    if not parent_dest_path.startswith('/'): parent_dest_path = '/' + parent_dest_path
    parent_path = f"/drives/{actual_drive_id}/root" if parent_dest_path == '/' else f"/drives/{actual_drive_id}/root:{parent_dest_path}"
    body = {"parentReference": {"path": parent_path}}; body["name"] = nuevo_nombre if nuevo_nombre is not None else nombre_archivo
    logger.info(f"Moviendo SP '{item_path_origen}' a '{nueva_ubicacion}'")
    return hacer_llamada_api("PATCH", url, headers, json_data=body)

def copiar_archivo(headers: Dict[str, str], nombre_archivo: str, nueva_ubicacion: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta_origen: str = '/', nuevo_nombre_copia: Optional[str] = None) -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive_name = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path_origen = ruta_carpeta_origen.strip('/')
    item_path_origen = f"/{nombre_archivo}" if not target_folder_path_origen else f"/{target_folder_path_origen}/{nombre_archivo}"
    item_endpoint_origen = _get_sp_item_path_endpoint(target_site_id, item_path_origen, target_drive_name)
    url = f"{item_endpoint_origen}/copy"
    try: actual_drive_id = _get_drive_id(headers, target_site_id, target_drive_name)
    except Exception as drive_err: raise Exception(f"Error obteniendo Drive ID para copiar: {drive_err}")
    parent_dest_path = nueva_ubicacion.strip();
    if not parent_dest_path.startswith('/'): parent_dest_path = '/' + parent_dest_path
    parent_path = "/drive/root" if parent_dest_path == '/' else f"/drive/root:{parent_dest_path}"
    body = {"parentReference": {"driveId": actual_drive_id, "path": parent_path}, "name": nuevo_nombre_copia or f"Copia de {nombre_archivo}" };
    logger.info(f"Iniciando copia SP '{item_path_origen}' a '{nueva_ubicacion}'")
    # Copia es asíncrona, devuelve 202 con URL de monitor. Helper actual no maneja esto bien. Llamada directa:
    try:
        current_headers = headers.copy(); current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        monitor_url = response.headers.get('Location')
        logger.info(f"Copia SP '{nombre_archivo}' iniciada. Monitor: {monitor_url}")
        return {"status": "Copia Iniciada", "code": response.status_code, "monitorUrl": monitor_url}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en copiar_archivo (SP): {e}", exc_info=True); raise Exception(f"Error API iniciando copia SP: {e}")
    except Exception as e: logger.error(f"Error inesperado en copiar_archivo (SP): {e}", exc_info=True); raise

def obtener_metadatos_archivo(headers: Dict[str, str], nombre_archivo: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = item_endpoint;
    logger.info(f"Obteniendo metadatos SP '{item_path}'")
    return hacer_llamada_api("GET", url, headers)

def actualizar_metadatos_archivo(headers: Dict[str, str], nombre_archivo: str, nuevos_valores: dict, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = item_endpoint;
    etag = nuevos_valores.pop('@odata.etag', None)
    current_headers = headers.copy()
    if etag: current_headers['If-Match'] = etag
    logger.info(f"Actualizando metadatos SP '{item_path}'")
    return hacer_llamada_api("PATCH", url, current_headers, json_data=nuevos_valores)

def obtener_contenido_archivo(headers: Dict[str, str], nombre_archivo: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> bytes:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = f"{item_endpoint}/content";
    try:
        logger.info(f"Obteniendo contenido SP '{item_path}'")
        # El helper espera JSON, hacemos llamada directa para bytes
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT * 2)
        response.raise_for_status(); logger.info(f"Contenido SP '{item_path}' obtenido."); return response.content
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en obtener_contenido_archivo (SP): {e}", exc_info=True); raise Exception(f"Error API obteniendo contenido: {e}")
    except Exception as e: logger.error(f"Error inesperado en obtener_contenido_archivo (SP): {e}", exc_info=True); raise

def actualizar_contenido_archivo(headers: Dict[str, str], nombre_archivo: str, nuevo_contenido: bytes, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = f"{item_endpoint}/content"; upload_headers = headers.copy(); upload_headers['Content-Type'] = 'application/octet-stream';
    try:
        logger.info(f"Actualizando contenido SP '{item_path}'")
        if len(nuevo_contenido) > 4*1024*1024: logger.warning(f"Archivo SP '{nombre_archivo}' > 4MB.")
        # Helper no maneja data binaria, usar requests directo
        response = requests.put(url, headers=upload_headers, data=nuevo_contenido, timeout=GRAPH_API_TIMEOUT * 3)
        response.raise_for_status(); data = response.json(); logger.info(f"Contenido SP '{item_path}' actualizado."); return data
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en actualizar_contenido_archivo (SP): {e}", exc_info=True); raise Exception(f"Error API actualizando contenido: {e}")
    except Exception as e: logger.error(f"Error inesperado en actualizar_contenido_archivo (SP): {e}", exc_info=True); raise

def crear_enlace_compartido_archivo(headers: Dict[str, str], nombre_archivo: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/', tipo_enlace: str = "view", alcance: str = "organization") -> dict:
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = f"{item_endpoint}/createLink"; body = {"type": tipo_enlace, "scope": alcance};
    logger.info(f"Creando enlace SP '{item_path}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

# --- Funciones de Memoria Persistente (Usando Lista SP) ---
def crear_lista_memoria(headers: Dict[str, str], site_id: Optional[str] = None) -> dict:
    """Crea la lista para memoria persistente si no existe."""
    columnas = [
        {"name": "SessionID", "text": {}, "indexed": True},
        {"name": "Clave", "text": {}, "indexed": True},
        {"name": "Valor", "text": {}} # Texto multilinea permitiría JSON como string
        # Podríamos añadir timestamp: {"name": "Timestamp", "dateTime": {}}
    ]
    try:
        logger.info(f"Intentando crear lista de memoria '{MEMORIA_LIST_NAME}'")
        # Verificar si ya existe podría ser útil antes de llamar a crear
        return crear_lista(headers=headers, nombre_lista=MEMORIA_LIST_NAME, columnas=columnas, site_id=site_id)
    except Exception as e:
        # Podría fallar si ya existe, verificar el error específico
        if "already exists" in str(e).lower():
            logger.warning(f"La lista de memoria '{MEMORIA_LIST_NAME}' ya existe.")
            # Devolver info de la lista existente? Necesitaríamos listar_listas y buscarla.
            # Por ahora, devolvemos un status indicando que existe.
            return {"status": "Lista ya existente", "nombre": MEMORIA_LIST_NAME}
        else:
            logger.error(f"Error al crear lista de memoria: {e}", exc_info=True)
            raise

def guardar_dato_memoria(headers: Dict[str, str], session_id: str, clave: str, valor: Any, site_id: Optional[str] = None) -> dict:
    """Guarda (o actualiza) un dato clave-valor para una sesión en la lista de memoria."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    # Convertir valor a string (JSON si es dict/list) para guardar en campo de texto
    if isinstance(valor, (dict, list)):
        valor_str = json.dumps(valor)
    else:
        valor_str = str(valor)

    # Intentar buscar si ya existe un item con esa SessionID y Clave para actualizarlo (PATCH)
    # Si no, crear uno nuevo (POST)
    filter_query = f"fields/SessionID eq '{session_id}' and fields/Clave eq '{clave}'"
    try:
        existing_items = listar_elementos_lista(headers=headers, lista_id_o_nombre=MEMORIA_LIST_NAME, site_id=target_site_id, filter_query=filter_query, top=1, select="id")
        item_id = existing_items.get("value", [])[0].get("id") if existing_items.get("value") else None
    except Exception as e:
        logger.warning(f"Error buscando item existente para memoria ({session_id}/{clave}): {e}. Se intentará crear.")
        item_id = None

    if item_id:
        # Actualizar existente
        logger.info(f"Actualizando dato en memoria: Session={session_id}, Clave={clave}")
        return actualizar_elemento_lista(headers=headers, lista_id_o_nombre=MEMORIA_LIST_NAME, item_id=item_id, nuevos_valores_campos={"Valor": valor_str}, site_id=target_site_id)
    else:
        # Crear nuevo
        logger.info(f"Guardando nuevo dato en memoria: Session={session_id}, Clave={clave}")
        datos_campos = {"SessionID": session_id, "Clave": clave, "Valor": valor_str}
        return agregar_elemento_lista(headers=headers, lista_id_o_nombre=MEMORIA_LIST_NAME, datos_campos=datos_campos, site_id=target_site_id) # Usar función específica si la creamos

def recuperar_datos_sesion(headers: Dict[str, str], session_id: str, site_id: Optional[str] = None) -> Dict[str, Any]:
    """Recupera todos los datos (clave-valor) asociados a una sesión."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    filter_query = f"fields/SessionID eq '{session_id}'"
    logger.info(f"Recuperando datos de memoria para Session={session_id}")
    items_data = listar_elementos_lista(headers=headers, lista_id_o_nombre=MEMORIA_LIST_NAME, site_id=target_site_id, filter_query=filter_query, expand_fields=True, select="fields/Clave,fields/Valor")

    # Reconstruir diccionario clave-valor
    memoria: Dict[str, Any] = {}
    for item in items_data.get("value", []):
        fields = item.get("fields", {})
        clave = fields.get("Clave")
        valor_str = fields.get("Valor")
        if clave and valor_str:
            try:
                # Intentar decodificar JSON si es un dict/list guardado
                memoria[clave] = json.loads(valor_str)
            except json.JSONDecodeError:
                memoria[clave] = valor_str # Guardar como string si no es JSON
    logger.info(f"Recuperados {len(memoria)} datos para Session={session_id}")
    return memoria

def eliminar_dato_memoria(headers: Dict[str, str], session_id: str, clave: str, site_id: Optional[str] = None) -> Optional[dict]:
    """Elimina un dato específico (por clave) de una sesión en memoria."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    filter_query = f"fields/SessionID eq '{session_id}' and fields/Clave eq '{clave}'"
    try:
        existing_items = listar_elementos_lista(headers=headers, lista_id_o_nombre=MEMORIA_LIST_NAME, site_id=target_site_id, filter_query=filter_query, top=1, select="id")
        item_id = existing_items.get("value", [])[0].get("id") if existing_items.get("value") else None
        if item_id:
            logger.info(f"Eliminando dato de memoria: Session={session_id}, Clave={clave}")
            return eliminar_elemento_lista(headers=headers, lista_id_o_nombre=MEMORIA_LIST_NAME, item_id=item_id, site_id=target_site_id)
        else:
            logger.warning(f"No se encontró dato para eliminar: Session={session_id}, Clave={clave}")
            return {"status": "No encontrado"}
    except Exception as e:
        logger.error(f"Error eliminando dato de memoria ({session_id}/{clave}): {e}", exc_info=True)
        raise

def eliminar_memoria_sesion(headers: Dict[str, str], session_id: str, site_id: Optional[str] = None) -> dict:
    """Elimina TODOS los datos asociados a una sesión específica."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    filter_query = f"fields/SessionID eq '{session_id}'"
    logger.info(f"Eliminando TODOS los datos de memoria para Session={session_id}")
    items_data = listar_elementos_lista(headers=headers, lista_id_o_nombre=MEMORIA_LIST_NAME, site_id=target_site_id, filter_query=filter_query, select="id")

    count = 0
    for item in items_data.get("value", []):
        item_id = item.get("id")
        if item_id:
            try:
                eliminar_elemento_lista(headers=headers, lista_id_o_nombre=MEMORIA_LIST_NAME, item_id=item_id, site_id=target_site_id)
                count += 1
            except Exception as del_err:
                logger.error(f"Error eliminando item {item_id} de memoria para sesión {session_id}: {del_err}")
                # Continuar eliminando los demás si falla uno? O parar? Decidimos continuar.
    logger.info(f"Eliminados {count} datos para Session={session_id}")
    return {"status": "Eliminados", "items_eliminados": count, "session_id": session_id}

# --- Funciones Avanzadas Adicionales ---
def exportar_datos_lista(headers: Dict[str, str], lista_id_o_nombre: str, formato: str = "json", site_id: Optional[str] = None) -> Union[dict, str]:
    """Exporta los datos de una lista en formato JSON o CSV."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    logger.info(f"Exportando datos de lista '{lista_id_o_nombre}' en formato {formato}")
    # Obtener todos los items (asume que listar_elementos_lista devuelve todos)
    # Podríamos añadir un parámetro 'top=-1' o similar para indicar "todos"
    items_data = listar_elementos_lista(headers=headers, lista_id_o_nombre=lista_id_o_nombre, site_id=target_site_id, expand_fields=True, top=999) # Limitar a 999 por llamada, necesita paginación real para >999

    items = items_data.get("value", [])
    if not items: return {} if formato.lower() == "json" else ""

    # Extraer solo los campos (fields)
    field_data = [item.get("fields", {}) for item in items]

    if formato.lower() == "json":
        return {"value": field_data} # Devolver como JSON
    elif formato.lower() == "csv":
        if not field_data: return ""
        output = StringIO()
        # Usar las claves del primer item como cabeceras CSV
        field_names = list(field_data[0].keys())
        writer = csv.DictWriter(output, fieldnames=field_names)
        writer.writeheader()
        writer.writerows(field_data)
        return output.getvalue()
    else:
        raise ValueError("Formato no soportado. Use 'json' o 'csv'.")

# --- FIN: Funciones SharePoint ---
