# actions/sharepoint.py (Refactorizado v2 - Corrección MyPy)

import logging
import requests # Necesario aquí solo para tipos de excepción (RequestException)
import os
import json # Para formateo de exportación y memoria
import csv # Para exportación CSV
from io import StringIO # Para exportación CSV
# Corrección MyPy: Añadir Callable para el tipo de acciones_disponibles si se importara aquí
from typing import Dict, List, Optional, Any, Union, cast

# Importar helper y constantes desde la estructura compartida
try:
    # Asume que shared está un nivel arriba de actions
    from ..shared.helpers.http_client import hacer_llamada_api
    from ..shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError as e:
    # Fallback crítico si la estructura no es reconocida
    logging.critical(f"Error CRÍTICO importando helpers/constantes en SharePoint: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs):
        raise NotImplementedError("Dependencia 'hacer_llamada_api' no importada correctamente.")

# Usar logger estándar de Azure Functions
logger = logging.getLogger("azure.functions")

# --- Configuración Leída de Variables de Entorno ---
SHAREPOINT_DEFAULT_SITE_ID = os.environ.get('SHAREPOINT_DEFAULT_SITE_ID')
SHAREPOINT_DEFAULT_DRIVE_ID = os.environ.get('SHAREPOINT_DEFAULT_DRIVE_ID', 'Documents')
MEMORIA_LIST_NAME = os.environ.get('SHAREPOINT_MEMORY_LIST', 'MemoriaPersistenteAsistente')

# --- Helper Interno para Obtener Site ID (Sin cambios) ---
def _obtener_site_id_sp(parametros: Dict[str, Any], headers: Dict[str, str]) -> str:
    # ... (código igual a la versión anterior) ...
    site_id_input: Optional[str] = parametros.get("site_id")
    if site_id_input and ',' in site_id_input: return site_id_input
    if site_id_input and (':' in site_id_input or '.' in site_id_input):
        site_path_lookup = site_id_input
        if ':' not in site_path_lookup: site_path_lookup = f"{site_path_lookup}:/"
        url = f"{BASE_URL}/sites/{site_path_lookup}?$select=id"
        try:
            logger.debug(f"Buscando Site ID por path/hostname: GET {url}")
            site_data = hacer_llamada_api("GET", url, headers)
            site_id = site_data.get("id")
            if site_id: logger.info(f"Site ID encontrado por path/hostname '{site_id_input}': {site_id}"); return site_id
            else: raise ValueError(f"Respuesta inválida buscando sitio '{site_id_input}', falta 'id'.")
        except requests.exceptions.RequestException as e:
            if e.response is not None and e.response.status_code == 404: logger.warning(f"No se encontró sitio por path/hostname '{site_id_input}' (404). Intentando default/raíz.")
            else: logger.warning(f"Error API buscando sitio por path/hostname '{site_id_input}': {e}. Intentando default/raíz.")
        except Exception as e: logger.warning(f"Error inesperado buscando sitio por path/hostname '{site_id_input}': {e}. Intentando default/raíz.")
    if SHAREPOINT_DEFAULT_SITE_ID: return SHAREPOINT_DEFAULT_SITE_ID
    url = f"{BASE_URL}/sites/root?$select=id"
    try:
        logger.debug(f"Obteniendo sitio raíz SP del tenant: GET {url}")
        site_data = hacer_llamada_api("GET", url, headers)
        site_id = site_data.get("id")
        if not site_id: raise ValueError("Respuesta de sitio raíz inválida, falta 'id'.")
        logger.info(f"Site ID raíz del tenant obtenido: {site_id}"); return site_id
    except Exception as e: logger.critical(f"Fallo crítico al obtener Site ID: {e}", exc_info=True); raise ValueError(f"No se pudo determinar el Site ID de SharePoint: {e}") from e


# --- Helpers Internos para Endpoints de Drive/Items (Sin cambios) ---
def _get_sp_drive_endpoint(site_id: str, drive_id_or_name: Optional[str] = None) -> str:
    # ... (código igual a la versión anterior) ...
    target_drive = drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
    return f"{BASE_URL}/sites/{site_id}/drives/{target_drive}"

def _get_sp_item_path_endpoint(site_id: str, item_path: str, drive_id_or_name: Optional[str] = None) -> str:
    # ... (código igual a la versión anterior) ...
    drive_endpoint = _get_sp_drive_endpoint(site_id, drive_id_or_name)
    safe_path = item_path.strip();
    if not safe_path: safe_path = '/'
    if not safe_path.startswith('/'): safe_path = '/' + safe_path
    if safe_path == '/': return f"{drive_endpoint}/root"
    else: return f"{drive_endpoint}/root:{safe_path}"


def _get_drive_id(headers: Dict[str, str], site_id: str, drive_id_or_name: Optional[str] = None) -> str:
    # ... (código igual a la versión anterior) ...
    drive_endpoint = _get_sp_drive_endpoint(site_id, drive_id_or_name)
    url = f"{drive_endpoint}?$select=id"
    try:
        logger.debug(f"Obteniendo Drive ID para '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'}': GET {url}")
        drive_data = hacer_llamada_api("GET", url, headers)
        actual_drive_id = drive_data.get('id')
        if not actual_drive_id: raise ValueError("Respuesta inválida, no se pudo obtener 'id' del drive.")
        logger.info(f"Drive ID obtenido: {actual_drive_id}"); return actual_drive_id
    except Exception as e: logger.error(f"Error API obteniendo Drive ID para '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}': {e}", exc_info=True); raise Exception(f"Error obteniendo Drive ID para biblioteca '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}': {e}") from e


# ============================================
# ==== FUNCIONES DE ACCIÓN PARA LISTAS SP ====
# ============================================
# (Funciones crear_lista, listar_listas, agregar_elemento_lista, listar_elementos_lista,
#  actualizar_elemento_lista, eliminar_elemento_lista sin cambios respecto a v2)

def crear_lista(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    # ... (código igual a la versión anterior) ...
    nombre_lista: Optional[str] = parametros.get("nombre_lista")
    columnas: Optional[List[Dict[str, Any]]] = parametros.get("columnas")
    if not nombre_lista: raise ValueError("Parámetro 'nombre_lista' es requerido.")
    if columnas and not isinstance(columnas, list): raise ValueError("Parámetro 'columnas' debe ser una lista de diccionarios.")
    target_site_id = _obtener_site_id_sp(parametros, headers)
    url = f"{BASE_URL}/sites/{target_site_id}/lists"
    columnas_final = []
    if columnas:
        for col in columnas:
            if isinstance(col, dict) and col.get("name"): # Validar que sea dict con name
                 columnas_final.append(col)
            else:
                 logger.warning(f"Ignorando columna inválida: {col}")
    body = {"displayName": nombre_lista, "columns": columnas_final, "list": {"template": "genericList"}}
    logger.info(f"Creando lista SP '{nombre_lista}' en sitio {target_site_id}")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def listar_listas(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    # ... (código igual a la versión anterior) ...
    select: str = parametros.get("select", "id,name,displayName,webUrl")
    target_site_id = _obtener_site_id_sp(parametros, headers)
    url = f"{BASE_URL}/sites/{target_site_id}/lists"
    params_query = {"$select": select} if select else None
    logger.info(f"Listando listas SP del sitio {target_site_id} (campos: {select})")
    return hacer_llamada_api("GET", url, headers, params=params_query)

def agregar_elemento_lista(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    # ... (código igual a la versión anterior) ...
    lista_id_o_nombre: Optional[str] = parametros.get("lista_id_o_nombre")
    datos_campos: Optional[Dict[str, Any]] = parametros.get("datos_campos")
    if not lista_id_o_nombre: raise ValueError("Parámetro 'lista_id_o_nombre' es requerido.")
    if not datos_campos or not isinstance(datos_campos, dict): raise ValueError("Parámetro 'datos_campos' (diccionario) es requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers)
    body = {"fields": datos_campos}
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items"
    logger.info(f"Agregando elemento a lista SP '{lista_id_o_nombre}' en sitio {target_site_id}")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def listar_elementos_lista(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    # ... (código igual a la versión anterior, con paginación usando helper) ...
    lista_id_o_nombre: Optional[str] = parametros.get("lista_id_o_nombre")
    expand_fields: bool = parametros.get("expand_fields", True)
    top: int = int(parametros.get("top", 100)); filter_query: Optional[str] = parametros.get("filter_query")
    select: Optional[str] = parametros.get("select"); order_by: Optional[str] = parametros.get("order_by")
    if not lista_id_o_nombre: raise ValueError("Parámetro 'lista_id_o_nombre' es requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers)
    url_base = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items"
    params_query: Dict[str, Any] = {'$top': min(top, 999)}
    if expand_fields:
        if select and 'fields/' in select:
            params_query['$expand'] = 'fields($select=' + ','.join(s.split('/')[1] for s in select.split(',') if s.startswith('fields/')) + ')'
            select_final = ','.join(s for s in select.split(',') if not s.startswith('fields/'))
            if select_final: params_query['$select'] = select_final
            elif '$select' in params_query: del params_query['$select']
        else: params_query['$expand'] = 'fields'
    if filter_query: params_query['$filter'] = filter_query
    if select and '$select' not in params_query: params_query['$select'] = select
    if order_by: params_query['$orderby'] = order_by
    all_items: List[Dict[str, Any]] = []; current_url: Optional[str] = url_base; page_count = 0; max_pages = 100
    try:
        while current_url and page_count < max_pages:
            page_count += 1; logger.info(f"Listando elementos SP lista '{lista_id_o_nombre}', Página: {page_count}")
            current_params = params_query if page_count == 1 else None
            # Asegurar que current_url no es None antes de llamar
            assert current_url is not None, "current_url se volvió None inesperadamente en bucle de paginación"
            data = hacer_llamada_api("GET", current_url, headers, params=current_params)
            if data:
                page_items = data.get('value', []); all_items.extend(page_items)
                current_url = data.get('@odata.nextLink');
                if not current_url: break
            else: logger.warning(f"Llamada a {current_url} devolvió None/vacío."); break
        if page_count >= max_pages: logger.warning(f"Límite de {max_pages} páginas alcanzado.")
        logger.info(f"Total elementos SP lista '{lista_id_o_nombre}': {len(all_items)}"); return {'value': all_items}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en listar_elementos_lista (SP) p{page_count}: {e}", exc_info=True); raise Exception(f"Error API listando elementos SP: {e}") from e
    except Exception as e: logger.error(f"Error inesperado en listar_elementos_lista (SP) p{page_count}: {e}", exc_info=True); raise

def actualizar_elemento_lista(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    # ... (código igual a la versión anterior) ...
    lista_id_o_nombre: Optional[str] = parametros.get("lista_id_o_nombre"); item_id: Optional[str] = parametros.get("item_id")
    nuevos_valores_campos: Optional[Dict[str, Any]] = parametros.get("nuevos_valores_campos")
    if not lista_id_o_nombre: raise ValueError("'lista_id_o_nombre' requerido.")
    if not item_id: raise ValueError("'item_id' requerido.")
    if not nuevos_valores_campos or not isinstance(nuevos_valores_campos, dict): raise ValueError("'nuevos_valores_campos' (dict) requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items/{item_id}/fields"
    current_headers = headers.copy(); body_data = nuevos_valores_campos.copy()
    etag = body_data.pop('@odata.etag', None)
    if etag: current_headers['If-Match'] = etag; logger.debug(f"Usando ETag '{etag}'.")
    logger.info(f"Actualizando elemento SP '{item_id}' en lista '{lista_id_o_nombre}'")
    return hacer_llamada_api("PATCH", url, current_headers, json_data=body_data)

def eliminar_elemento_lista(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    # ... (código igual a la versión anterior) ...
    lista_id_o_nombre: Optional[str] = parametros.get("lista_id_o_nombre"); item_id: Optional[str] = parametros.get("item_id")
    etag: Optional[str] = parametros.get("etag")
    if not lista_id_o_nombre: raise ValueError("'lista_id_o_nombre' requerido.")
    if not item_id: raise ValueError("'item_id' requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items/{item_id}"
    current_headers = headers.copy()
    if etag: current_headers['If-Match'] = etag; logger.debug(f"Usando ETag '{etag}'.")
    else: logger.warning(f"Eliminando item SP {item_id} sin ETag.")
    logger.info(f"Eliminando elemento SP '{item_id}' de lista '{lista_id_o_nombre}'")
    hacer_llamada_api("DELETE", url, current_headers); return {"status": "Eliminado", "item_id": item_id, "lista": lista_id_o_nombre}


# ========================================================
# ==== FUNCIONES DE ACCIÓN PARA DOCUMENTOS (DRIVES) ====
# ========================================================
# (Funciones listar_documentos_biblioteca, subir_documento, eliminar_archivo,
#  crear_carpeta_biblioteca sin cambios respecto a v2)

def listar_documentos_biblioteca(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    # ... (código igual a la versión anterior, con paginación usando helper) ...
    biblioteca: Optional[str] = parametros.get("biblioteca"); ruta_carpeta: str = parametros.get("ruta_carpeta", '/')
    top: int = int(parametros.get("top", 100))
    target_site_id = _obtener_site_id_sp(parametros, headers)
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_carpeta, biblioteca)
    url_base = f"{item_endpoint}/children"; params_query = {'$top': min(top, 999)};
    all_files: List[Dict[str, Any]] = []; current_url: Optional[str] = url_base; page_count = 0; max_pages = 100
    try:
        while current_url and page_count < max_pages:
            page_count += 1; target_drive_name = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
            logger.info(f"Listando docs SP biblioteca '{target_drive_name}', Ruta: '{ruta_carpeta}', Página: {page_count}")
            current_params = params_query if page_count == 1 else None
            assert current_url is not None, "current_url None en paginación listar_documentos"
            data = hacer_llamada_api("GET", current_url, headers, params=current_params)
            if data:
                page_items = data.get('value', []); all_files.extend(page_items)
                current_url = data.get('@odata.nextLink');
                if not current_url: break
            else: logger.warning(f"Llamada a {current_url} devolvió None/vacío."); break
        if page_count >= max_pages: logger.warning(f"Límite de {max_pages} páginas alcanzado.")
        logger.info(f"Total docs/carpetas SP encontrados: {len(all_files)}"); return {'value': all_files}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en listar_documentos_biblioteca (SP) p{page_count}: {e}", exc_info=True); raise Exception(f"Error API listando documentos SP: {e}") from e
    except Exception as e: logger.error(f"Error inesperado en listar_documentos_biblioteca (SP) p{page_count}: {e}", exc_info=True); raise

def subir_documento(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    # ... (código igual a la versión anterior, con lógica de sesión de carga) ...
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo"); contenido_bytes: Optional[bytes] = parametros.get("contenido_bytes")
    biblioteca: Optional[str] = parametros.get("biblioteca"); ruta_carpeta_destino: str = parametros.get("ruta_carpeta_destino", '/')
    conflict_behavior: str = parametros.get("conflict_behavior", "rename")
    if not nombre_archivo: raise ValueError("'nombre_archivo' requerido.")
    if contenido_bytes is None or not isinstance(contenido_bytes, bytes): raise ValueError("'contenido_bytes' (bytes) requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers); target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
    target_folder_path = ruta_carpeta_destino.strip('/'); target_file_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, target_file_path, target_drive)
    file_size_mb = len(contenido_bytes) / (1024 * 1024)
    logger.info(f"Subiendo doc SP '{nombre_archivo}' ({file_size_mb:.2f} MB) a '{ruta_carpeta_destino}' conflict='{conflict_behavior}'")
    if file_size_mb > 4.0:
        create_session_url = f"{item_endpoint}:/createUploadSession"; session_body = {"item": {"@microsoft.graph.conflictBehavior": conflict_behavior}}
        try:
            logger.info(f"Archivo > 4MB. Creando sesión de carga..."); session_info = hacer_llamada_api("POST", create_session_url, headers, json_data=session_body)
            upload_url = session_info.get("uploadUrl");
            if not upload_url: raise ValueError("No se pudo obtener 'uploadUrl'.")
            logger.info(f"Sesión creada. URL: {upload_url[:50]}...")
            chunk_size = 5 * 1024 * 1024; start_byte = 0; total_bytes = len(contenido_bytes); last_response_json = {}
            while start_byte < total_bytes:
                end_byte = min(start_byte + chunk_size - 1, total_bytes - 1); chunk_data = contenido_bytes[start_byte : end_byte + 1]
                content_range = f"bytes {start_byte}-{end_byte}/{total_bytes}"; chunk_headers = {'Content-Length': str(len(chunk_data)), 'Content-Range': content_range}
                logger.debug(f"Subiendo chunk SP: {content_range}")
                chunk_timeout = max(GRAPH_API_TIMEOUT, int(file_size_mb * 5));
                chunk_response = requests.put(upload_url, headers=chunk_headers, data=chunk_data, timeout=chunk_timeout)
                chunk_response.raise_for_status(); start_byte = end_byte + 1
                if chunk_response.content:
                    try: last_response_json = chunk_response.json()
                    except json.JSONDecodeError: pass
            logger.info(f"Doc SP '{nombre_archivo}' subido por sesión."); return last_response_json
        except requests.exceptions.RequestException as e: logger.error(f"Error Request sesión carga SP '{nombre_archivo}': {e}", exc_info=True); raise Exception(f"Error API sesión carga SP: {e}") from e
        except Exception as e: logger.error(f"Error inesperado sesión carga SP '{nombre_archivo}': {e}", exc_info=True); raise
    else:
        url_put_simple = f"{item_endpoint}:/content"; params_query = {"@microsoft.graph.conflictBehavior": conflict_behavior}
        upload_headers = headers.copy(); upload_headers['Content-Type'] = 'application/octet-stream';
        try:
             simple_upload_timeout = max(GRAPH_API_TIMEOUT, int(file_size_mb * 10))
             resultado = hacer_llamada_api("PUT", url_put_simple, upload_headers, params=params_query, data=contenido_bytes, timeout=simple_upload_timeout, expect_json=True)
             logger.info(f"Doc SP '{nombre_archivo}' subido (simple)."); return resultado
        except requests.exceptions.RequestException as e: logger.error(f"Error Request subida simple SP '{nombre_archivo}': {e}", exc_info=True); raise Exception(f"Error API subiendo doc SP (simple): {e}") from e
        except Exception as e: logger.error(f"Error inesperado subida simple SP '{nombre_archivo}': {e}", exc_info=True); raise

# Renombrado para mapeo_acciones.py
def eliminar_archivo_biblioteca(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Elimina un archivo o carpeta de una biblioteca."""
    # (Código igual a eliminar_archivo de onedrive.py, adaptado a SP)
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta")
    biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta: str = parametros.get("ruta_carpeta", '/')
    if not nombre_archivo_o_carpeta: raise ValueError("'nombre_archivo_o_carpeta' requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
    target_folder_path = ruta_carpeta.strip('/'); item_path = f"/{nombre_archivo_o_carpeta}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo_o_carpeta}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = item_endpoint
    logger.info(f"Eliminando archivo/carpeta SP '{item_path}' en biblioteca '{target_drive}'")
    hacer_llamada_api("DELETE", url, headers); return {"status": "Eliminado", "path": item_path}

# Renombrado para mapeo_acciones.py
def crear_carpeta_biblioteca(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Crea una nueva carpeta en una biblioteca/carpeta padre."""
    # (Código igual a crear_carpeta de onedrive.py, adaptado a SP)
    nombre_carpeta: Optional[str] = parametros.get("nombre_carpeta")
    biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta_padre: str = parametros.get("ruta_carpeta_padre", '/')
    conflict_behavior: str = parametros.get("conflict_behavior", "rename")
    if not nombre_carpeta: raise ValueError("'nombre_carpeta' requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
    parent_folder_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_carpeta_padre, target_drive)
    url = f"{parent_folder_endpoint}/children"
    body: Dict[str, Any] = {"name": nombre_carpeta, "folder": {}, "@microsoft.graph.conflictBehavior": conflict_behavior}
    logger.info(f"Creando carpeta SP '{nombre_carpeta}' en '{ruta_carpeta_padre}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)


# ======================================================
# ==== FUNCIONES AVANZADAS DE ARCHIVOS (MOVER/COPIAR) ====
# ======================================================

# Renombrado para mapeo_acciones.py
def mover_archivo_biblioteca(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Mueve un archivo o carpeta a una nueva ubicación (dentro del mismo Drive/Biblioteca)."""
    # (Código igual a mover_archivo de onedrive.py, adaptado a SP)
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta")
    ruta_origen: str = parametros.get("ruta_origen", '/'); nueva_ruta_carpeta_padre: Optional[str] = parametros.get("nueva_ruta_carpeta_padre")
    nuevo_nombre: Optional[str] = parametros.get("nuevo_nombre"); biblioteca: Optional[str] = parametros.get("biblioteca")
    if not nombre_archivo_o_carpeta: raise ValueError("'nombre_archivo_o_carpeta' requerido.")
    if nueva_ruta_carpeta_padre is None: raise ValueError("'nueva_ruta_carpeta_padre' requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers); target_drive_name = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
    target_folder_path_origen = ruta_origen.strip('/'); item_path_origen = f"/{nombre_archivo_o_carpeta}" if not target_folder_path_origen else f"/{target_folder_path_origen}/{nombre_archivo_o_carpeta}"
    item_origen_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path_origen, target_drive_name); url = item_origen_endpoint
    try: actual_drive_id = _get_drive_id(headers, target_site_id, target_drive_name)
    except Exception as drive_err: raise Exception(f"Error obteniendo Drive ID para mover SP: {drive_err}") from drive_err
    parent_dest_path = nueva_ruta_carpeta_padre.strip();
    if not parent_dest_path.startswith('/'): parent_dest_path = '/' + parent_dest_path
    parent_path_ref = f"/drives/{actual_drive_id}/root" if parent_dest_path == '/' else f"/drives/{actual_drive_id}/root:{parent_dest_path}"
    body: Dict[str, Any] = {"parentReference": {"path": parent_path_ref}}; body["name"] = nuevo_nombre if nuevo_nombre is not None else nombre_archivo_o_carpeta
    logger.info(f"Moviendo SP '{item_path_origen}' a '{nueva_ruta_carpeta_padre}'")
    return hacer_llamada_api("PATCH", url, headers, json_data=body)

# Renombrado para mapeo_acciones.py
def copiar_archivo_biblioteca(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Copia un archivo a una nueva ubicación. Operación asíncrona."""
    # Corrección MyPy L383: Añadir type hint explícito a current_headers
    # (Código igual a copiar_archivo de onedrive.py, adaptado a SP)
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo"); ruta_origen: str = parametros.get("ruta_origen", '/')
    nueva_ruta_carpeta_padre: Optional[str] = parametros.get("nueva_ruta_carpeta_padre"); nuevo_nombre_copia: Optional[str] = parametros.get("nuevo_nombre_copia")
    biblioteca: Optional[str] = parametros.get("biblioteca"); drive_id_destino: Optional[str] = parametros.get("drive_id_destino")
    if not nombre_archivo: raise ValueError("'nombre_archivo' requerido.")
    if nueva_ruta_carpeta_padre is None: raise ValueError("'nueva_ruta_carpeta_padre' requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers); target_drive_name_origen = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
    target_folder_path_origen = ruta_origen.strip('/'); item_path_origen = f"/{nombre_archivo}" if not target_folder_path_origen else f"/{target_folder_path_origen}/{nombre_archivo}"
    item_origen_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path_origen, target_drive_name_origen); url = f"{item_origen_endpoint}/copy"
    if not drive_id_destino:
        try: drive_id_destino = _get_drive_id(headers, target_site_id, target_drive_name_origen)
        except Exception as drive_err: raise Exception(f"Error obteniendo Drive ID origen SP para copiar: {drive_err}") from drive_err
    parent_dest_path = nueva_ruta_carpeta_padre.strip();
    if not parent_dest_path.startswith('/'): parent_dest_path = '/' + parent_dest_path
    parent_path_ref = f"/drives/{drive_id_destino}/root" if parent_dest_path == '/' else f"/drives/{drive_id_destino}/root:{parent_dest_path}"
    body: Dict[str, Any] = {"parentReference": {"driveId": drive_id_destino, "path": parent_path_ref}}; body["name"] = nuevo_nombre_copia or f"Copia de {nombre_archivo}";
    logger.info(f"Iniciando copia asíncrona SP de '{item_path_origen}' a Drive '{drive_id_destino}', Path: '{nueva_ruta_carpeta_padre}'")
    # Usar helper con expect_json=False
    response = hacer_llamada_api("POST", url, headers, json_data=body, expect_json=False)
    if isinstance(response, requests.Response) and response.status_code == 202:
        monitor_url = response.headers.get('Location'); logger.info(f"Copia SP '{nombre_archivo}' iniciada. Monitor URL: {monitor_url}");
        return {"status": "Copia Iniciada", "status_code": 202, "monitorUrl": monitor_url}
    elif isinstance(response, requests.Response): logger.error(f"Respuesta inesperada al iniciar copia SP: {response.status_code}."); raise Exception(f"Respuesta inesperada al iniciar copia SP: {response.status_code}")
    else: logger.error(f"Respuesta inesperada del helper al iniciar copia SP: {type(response)}"); raise Exception("Error interno al procesar copia SP.")


# ======================================================
# ==== FUNCIONES DE METADATOS Y CONTENIDO ARCHIVOS ====
# ======================================================

# Renombrado para mapeo_acciones.py
def obtener_metadatos_archivo_biblioteca(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Obtiene metadatos de un archivo o carpeta en una biblioteca."""
    # (Código igual a obtener_metadatos_archivo de onedrive.py, adaptado a SP)
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta"); biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta: str = parametros.get("ruta_carpeta", '/');
    if not nombre_archivo_o_carpeta: raise ValueError("'nombre_archivo_o_carpeta' requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers); target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
    target_folder_path = ruta_carpeta.strip('/'); item_path = f"/{nombre_archivo_o_carpeta}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo_o_carpeta}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive); url = item_endpoint;
    logger.info(f"Obteniendo metadatos SP '{item_path}'")
    return hacer_llamada_api("GET", url, headers)

# Renombrado para mapeo_acciones.py
def actualizar_metadatos_archivo_biblioteca(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Actualiza metadatos de un archivo/carpeta en una biblioteca. Soporta ETag."""
    # (Código igual a actualizar_metadatos_archivo de onedrive.py, adaptado a SP)
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta"); nuevos_valores: Optional[Dict[str, Any]] = parametros.get("nuevos_valores")
    biblioteca: Optional[str] = parametros.get("biblioteca"); ruta_carpeta: str = parametros.get("ruta_carpeta", '/');
    if not nombre_archivo_o_carpeta: raise ValueError("'nombre_archivo_o_carpeta' requerido.")
    if not nuevos_valores or not isinstance(nuevos_valores, dict): raise ValueError("'nuevos_valores' (dict) requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers); target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
    target_folder_path = ruta_carpeta.strip('/'); item_path = f"/{nombre_archivo_o_carpeta}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo_o_carpeta}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive); url = item_endpoint;
    current_headers = headers.copy(); body_data = nuevos_valores.copy(); etag = body_data.pop('@odata.etag', None)
    if etag: current_headers['If-Match'] = etag; logger.debug("Usando ETag.")
    logger.info(f"Actualizando metadatos SP '{item_path}'")
    return hacer_llamada_api("PATCH", url, current_headers, json_data=body_data)

# Renombrado para mapeo_acciones.py
def obtener_contenido_archivo_biblioteca(parametros: Dict[str, Any], headers: Dict[str, str]) -> bytes:
    """Descarga el contenido de un archivo de una biblioteca."""
    # (Código igual a descargar_archivo de onedrive.py, adaptado a SP)
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo"); biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta: str = parametros.get("ruta_carpeta", '/');
    if not nombre_archivo: raise ValueError("'nombre_archivo' requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers); target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
    target_folder_path = ruta_carpeta.strip('/'); item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive); url = f"{item_endpoint}/content";
    logger.info(f"Obteniendo contenido SP '{item_path}'")
    download_timeout = max(GRAPH_API_TIMEOUT, 60)
    response = hacer_llamada_api("GET", url, headers, timeout=download_timeout, expect_json=False)
    if isinstance(response, requests.Response): logger.info(f"Contenido SP '{item_path}' obtenido ({len(response.content)} bytes)."); return response.content
    else: logger.error(f"Respuesta inesperada del helper al obtener contenido SP: {type(response)}"); raise Exception("Error interno al obtener contenido SP.")

# Renombrado para mapeo_acciones.py
def actualizar_contenido_archivo_biblioteca(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Actualiza/Reemplaza el contenido de un archivo existente en una biblioteca."""
    # (Código igual a subir_archivo de onedrive.py, adaptado a SP, pero usando PUT en /content)
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo"); nuevo_contenido_bytes: Optional[bytes] = parametros.get("nuevo_contenido_bytes")
    biblioteca: Optional[str] = parametros.get("biblioteca"); ruta_carpeta: str = parametros.get("ruta_carpeta", '/');
    if not nombre_archivo: raise ValueError("'nombre_archivo' requerido.")
    if nuevo_contenido_bytes is None or not isinstance(nuevo_contenido_bytes, bytes): raise ValueError("'nuevo_contenido_bytes' (bytes) requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers); target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
    target_folder_path = ruta_carpeta.strip('/'); item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive); url = f"{item_endpoint}/content";
    upload_headers = headers.copy(); upload_headers['Content-Type'] = 'application/octet-stream';
    file_size_mb = len(nuevo_contenido_bytes) / (1024 * 1024)
    logger.info(f"Actualizando contenido SP '{item_path}' ({file_size_mb:.2f} MB)")
    if file_size_mb > 4.0: logger.warning(f"Actualizando archivo SP > 4MB con PUT simple. Considera sesión de carga."); # Podría fallar
    try:
        update_timeout = max(GRAPH_API_TIMEOUT, int(file_size_mb * 10))
        resultado = hacer_llamada_api("PUT", url, upload_headers, data=nuevo_contenido_bytes, timeout=update_timeout, expect_json=True)
        logger.info(f"Contenido SP '{item_path}' actualizado."); return resultado
    except requests.exceptions.RequestException as e: logger.error(f"Error Request actualizando contenido SP '{item_path}': {e}", exc_info=True); raise Exception(f"Error API actualizando contenido SP: {e}") from e
    except Exception as e: logger.error(f"Error inesperado actualizando contenido SP '{item_path}': {e}", exc_info=True); raise

# Renombrado para mapeo_acciones.py
def crear_enlace_compartido_archivo_biblioteca(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """Crea un enlace para compartir un archivo o carpeta en una biblioteca."""
    # (Código igual a crear_enlace_compartido_archivo de onedrive.py, adaptado a SP)
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta"); biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta: str = parametros.get("ruta_carpeta", '/'); tipo_enlace: str = parametros.get("tipo_enlace", "view"); alcance: str = parametros.get("alcance", "organization")
    password: Optional[str] = parametros.get("password"); expirationDateTime: Optional[str] = parametros.get("expirationDateTime")
    if not nombre_archivo_o_carpeta: raise ValueError("'nombre_archivo_o_carpeta' requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers); target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
    target_folder_path = ruta_carpeta.strip('/'); item_path = f"/{nombre_archivo_o_carpeta}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo_o_carpeta}"
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive); url = f"{item_endpoint}/createLink";
    body: Dict[str, Any] = {"type": tipo_enlace, "scope": alcance};
    if password: body["password"] = password
    if expirationDateTime: body["expirationDateTime"] = expirationDateTime
    logger.info(f"Creando enlace SP (tipo: {tipo_enlace}, alcance: {alcance}) para '{item_path}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)


# ======================================================
# ==== FUNCIONES DE MEMORIA PERSISTENTE (LISTA SP) ====
# ======================================================
# (Funciones _ensure_memory_list_exists, guardar_dato_memoria,
#  recuperar_datos_sesion, eliminar_dato_memoria, eliminar_memoria_sesion
#  sin cambios respecto a v2, pero se añade corrección MyPy L580)

def _ensure_memory_list_exists(headers: Dict[str, str], site_id: str) -> bool:
    # ... (código igual a la versión anterior) ...
    try:
        list_url = f"{BASE_URL}/sites/{site_id}/lists/{MEMORIA_LIST_NAME}?$select=id"; hacer_llamada_api("GET", list_url, headers)
        logger.debug(f"Lista memoria '{MEMORIA_LIST_NAME}' existe."); return True
    except requests.exceptions.RequestException as e:
        if e.response is not None and e.response.status_code == 404:
            logger.info(f"Lista memoria '{MEMORIA_LIST_NAME}' no encontrada. Creando...");
            columnas = [{"name": "SessionID", "text": {}, "indexed": True}, {"name": "Clave", "text": {}, "indexed": True}, {"name": "Valor", "text": {"allowMultipleLines": True, "textType": "plain"}}, {"name": "Timestamp", "dateTime": {}, "indexed": True}]
            params_crear = {"nombre_lista": MEMORIA_LIST_NAME, "columnas": columnas, "site_id": site_id}
            try: crear_lista(params_crear, headers); logger.info(f"Lista memoria '{MEMORIA_LIST_NAME}' creada."); return True
            except Exception as create_err: logger.critical(f"Fallo al crear lista memoria!: {create_err}", exc_info=True); return False
        else: logger.error(f"Error verificando lista memoria: {e}", exc_info=True); return False
    except Exception as e: logger.error(f"Error inesperado verificando lista memoria: {e}", exc_info=True); return False

def guardar_dato_memoria(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    # ... (código igual a la versión anterior) ...
    session_id: Optional[str] = parametros.get("session_id"); clave: Optional[str] = parametros.get("clave"); valor: Any = parametros.get("valor")
    if not session_id: raise ValueError("'session_id' requerido.");
    if not clave: raise ValueError("'clave' requerido.");
    if valor is None: raise ValueError("'valor' requerido (no None).")
    target_site_id = _obtener_site_id_sp(parametros, headers)
    if not _ensure_memory_list_exists(headers, target_site_id): raise Exception(f"No se pudo asegurar lista memoria '{MEMORIA_LIST_NAME}'.")
    try:
        if isinstance(valor, (dict, list, bool)): valor_str = json.dumps(valor)
        elif isinstance(valor, (int, float)): valor_str = str(valor)
        elif isinstance(valor, str): valor_str = valor
        else: valor_str = str(valor); logger.warning(f"Guardando tipo no estándar '{type(valor)}' como string.")
    except Exception as json_err: raise ValueError(f"Error serializando valor para clave '{clave}': {json_err}") from json_err
    filter_query = f"fields/SessionID eq '{session_id}' and fields/Clave eq '{clave}'"
    params_listar = {"lista_id_o_nombre": MEMORIA_LIST_NAME, "site_id": target_site_id, "filter_query": filter_query, "top": 1, "select": "id,@odata.etag"}
    item_id: Optional[str] = None; item_etag: Optional[str] = None
    try:
        existing_items_data = listar_elementos_lista(params_listar, headers); existing_items = existing_items_data.get("value", [])
        if existing_items: item_id = existing_items[0].get("id"); item_etag = existing_items[0].get("@odata.etag")
    except Exception as e: logger.warning(f"Error buscando item memoria ({session_id}/{clave}): {e}. Se intentará crear.")
    datos_campos = {"SessionID": session_id, "Clave": clave, "Valor": valor_str, "Timestamp": datetime.utcnow().isoformat() + "Z"}
    if item_id:
        logger.info(f"Actualizando memoria: Session={session_id}, Clave={clave}")
        params_actualizar = {"lista_id_o_nombre": MEMORIA_LIST_NAME, "item_id": item_id, "nuevos_valores_campos": datos_campos, "site_id": target_site_id}
        if item_etag: params_actualizar["nuevos_valores_campos"]["@odata.etag"] = item_etag
        return actualizar_elemento_lista(params_actualizar, headers)
    else:
        logger.info(f"Guardando nuevo dato memoria: Session={session_id}, Clave={clave}")
        params_agregar = {"lista_id_o_nombre": MEMORIA_LIST_NAME, "datos_campos": datos_campos, "site_id": target_site_id}
        return agregar_elemento_lista(params_agregar, headers)

def recuperar_datos_sesion(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    # Corrección MyPy L580: Añadir validación explícita de tipo para clave
    # ... (código igual a la versión anterior hasta el bucle) ...
    session_id: Optional[str] = parametros.get("session_id");
    if not session_id: raise ValueError("'session_id' requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers)
    if not _ensure_memory_list_exists(headers, target_site_id): logger.warning(f"Lista memoria no encontrada."); return {}
    filter_query = f"fields/SessionID eq '{session_id}'"; select_fields = "id,fields/Clave,fields/Valor,fields/Timestamp"; order_by = "fields/Timestamp desc"
    logger.info(f"Recuperando memoria Session={session_id}")
    params_listar = {"lista_id_o_nombre": MEMORIA_LIST_NAME, "site_id": target_site_id, "filter_query": filter_query, "expand_fields": True, "select": select_fields, "order_by": order_by, "top": 999}
    items_data = listar_elementos_lista(params_listar, headers)
    memoria: Dict[str, Any] = {}
    for item in items_data.get("value", []):
        fields = item.get("fields", {})
        clave_any = fields.get("Clave") # Obtener como Any primero
        valor_str = fields.get("Valor")
        # Corrección MyPy L580: Asegurar que clave es string antes de usarla como índice
        if isinstance(clave_any, str) and clave_any and valor_str:
            clave: str = clave_any # Ahora clave es definitivamente str
            if clave not in memoria: # Solo guardar el más reciente (por order_by)
                try: memoria[clave] = json.loads(valor_str) # Intentar parsear JSON
                except json.JSONDecodeError: memoria[clave] = valor_str # Guardar como string si falla
                except Exception as parse_err: logger.warning(f"Error parseando valor memoria '{clave}': {parse_err}."); memoria[clave] = valor_str
        elif clave_any:
             logger.warning(f"Clave inválida encontrada en memoria (no es string o está vacía): {clave_any}")

    logger.info(f"Recuperados {len(memoria)} datos únicos para Session={session_id}"); return memoria

def eliminar_dato_memoria(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    # ... (código igual a la versión anterior) ...
    session_id: Optional[str] = parametros.get("session_id"); clave: Optional[str] = parametros.get("clave")
    if not session_id: raise ValueError("'session_id' requerido.");
    if not clave: raise ValueError("'clave' requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers)
    if not _ensure_memory_list_exists(headers, target_site_id): return {"status": "Lista no encontrada"}
    filter_query = f"fields/SessionID eq '{session_id}' and fields/Clave eq '{clave}'"
    params_listar = {"lista_id_o_nombre": MEMORIA_LIST_NAME, "site_id": target_site_id, "filter_query": filter_query, "top": 1, "select": "id"}
    item_id: Optional[str] = None
    try:
        existing_items_data = listar_elementos_lista(params_listar, headers); existing_items = existing_items_data.get("value", [])
        if existing_items: item_id = existing_items[0].get("id")
    except Exception as e: logger.error(f"Error buscando item a eliminar ({session_id}/{clave}): {e}", exc_info=True); raise Exception(f"Error buscando item a eliminar: {e}") from e
    if item_id:
        logger.info(f"Eliminando memoria: Session={session_id}, Clave={clave}")
        params_eliminar = {"lista_id_o_nombre": MEMORIA_LIST_NAME, "item_id": item_id, "site_id": target_site_id}
        return eliminar_elemento_lista(params_eliminar, headers)
    else: logger.warning(f"No se encontró dato memoria para eliminar: Session={session_id}, Clave={clave}"); return {"status": "No encontrado", "session_id": session_id, "clave": clave}

def eliminar_memoria_sesion(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    # ... (código igual a la versión anterior) ...
    session_id: Optional[str] = parametros.get("session_id");
    if not session_id: raise ValueError("'session_id' requerido.")
    target_site_id = _obtener_site_id_sp(parametros, headers)
    if not _ensure_memory_list_exists(headers, target_site_id): return {"status": "Lista no encontrada", "items_eliminados": 0}
    filter_query = f"fields/SessionID eq '{session_id}'"; logger.info(f"Listando TODOS los items memoria para eliminar Session={session_id}")
    params_listar = {"lista_id_o_nombre": MEMORIA_LIST_NAME, "site_id": target_site_id, "filter_query": filter_query, "select": "id", "top": 999}
    items_data = listar_elementos_lista(params_listar, headers); items_to_delete = items_data.get("value", []); item_ids_to_delete = [item.get("id") for item in items_to_delete if item.get("id")]
    if not item_ids_to_delete: logger.info(f"No se encontraron datos para eliminar Session={session_id}"); return {"status": "Sin datos", "items_eliminados": 0, "session_id": session_id}
    logger.info(f"Se eliminarán {len(item_ids_to_delete)} datos para Session={session_id}"); count_deleted = 0; count_failed = 0
    for item_id in item_ids_to_delete:
        try: params_eliminar = {"lista_id_o_nombre": MEMORIA_LIST_NAME, "item_id": item_id, "site_id": target_site_id}; eliminar_elemento_lista(params_eliminar, headers); count_deleted += 1
        except Exception as del_err: logger.error(f"Error eliminando item {item_id} memoria sesión {session_id}: {del_err}"); count_failed += 1
    logger.info(f"Eliminación memoria sesión {session_id}: {count_deleted} exitosos, {count_failed} fallidos.")
    return {"status": "Completado" if count_failed == 0 else "Completado con errores", "items_eliminados": count_deleted, "items_fallidos": count_failed, "session_id": session_id}

# ======================================================
# ==== FUNCIONES ADICIONALES (EJ: EXPORTAR)         ====
# ======================================================
# (Función exportar_datos_lista sin cambios respecto a v2)
def exportar_datos_lista(parametros: Dict[str, Any], headers: Dict[str, str]) -> Union[Dict[str, Any], str]:
    # ... (código igual a la versión anterior) ...
    lista_id_o_nombre: Optional[str] = parametros.get("lista_id_o_nombre"); formato: str = parametros.get("formato", "json").lower()
    if not lista_id_o_nombre: raise ValueError("'lista_id_o_nombre' requerido.");
    if formato not in ["json", "csv"]: raise ValueError("Formato no soportado. Use 'json' o 'csv'.")
    target_site_id = _obtener_site_id_sp(parametros, headers); logger.info(f"Exportando lista '{lista_id_o_nombre}' formato {formato}")
    params_listar = {"lista_id_o_nombre": lista_id_o_nombre, "site_id": target_site_id, "expand_fields": True, "top": 999}
    items_data = listar_elementos_lista(params_listar, headers); items = items_data.get("value", [])
    if not items: return {"value": []} if formato == "json" else ""
    field_data = [];
    for item in items: fields = item.get("fields", {}); fields["_ItemID_"] = item.get("id"); field_data.append(fields)
    if formato == "json": logger.info(f"Exportando {len(field_data)} items como JSON."); return {"value": field_data}
    else:
        logger.info(f"Exportando {len(field_data)} items como CSV.");
        if not field_data: return ""
        output = StringIO(); field_names = list(field_data[0].keys())
        if "_ItemID_" in field_names: field_names.insert(0, field_names.pop(field_names.index("_ItemID_")))
        writer = csv.DictWriter(output, fieldnames=field_names, extrasaction='ignore', quoting=csv.QUOTE_MINIMAL)
        writer.writeheader(); writer.writerows(field_data); csv_content = output.getvalue(); output.close(); return csv_content

# --- FIN DEL MÓDULO actions/sharepoint.py ---
