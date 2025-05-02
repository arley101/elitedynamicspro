"""
actions/sharepoint.py (Refactorizado)

Funciones para interactuar con SharePoint (Listas y Documentos) usando
el helper centralizado `hacer_llamada_api`. Incluye funciones para
simular memoria persistente usando una lista de SharePoint.

Refactorización:
- Firmas de funciones unificadas a `(parametros: Dict[str, Any], headers: Dict[str, str])`.
- Uso consistente de `hacer_llamada_api` para TODAS las interacciones HTTP.
- Manejo de paginación, datos binarios y respuestas 202 usando el helper.
- Extracción de argumentos desde el diccionario `parametros`.
"""

import logging
import requests # Necesario aquí solo para tipos de excepción (RequestException)
import os
import json # Para formateo de exportación y memoria
import csv # Para exportación CSV
from io import StringIO # Para exportación CSV
from typing import Dict, List, Optional, Any, Union

# Importar helper y constantes desde la estructura compartida
try:
    # Asume que shared está un nivel arriba de actions
    from ..shared.helpers.http_client import hacer_llamada_api
    from ..shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError as e:
    # Fallback crítico si la estructura no es reconocida
    logging.critical(f"Error CRÍTICO importando helpers/constantes en SharePoint: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    # Definir mocks o lanzar error para evitar ejecución parcial
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs):
        raise NotImplementedError("Dependencia 'hacer_llamada_api' no importada correctamente.")

# Usar logger estándar de Azure Functions
logger = logging.getLogger("azure.functions")

# --- Configuración Leída de Variables de Entorno ---
# Es buena práctica permitir configuración externa
SHAREPOINT_DEFAULT_SITE_ID = os.environ.get('SHAREPOINT_DEFAULT_SITE_ID')
SHAREPOINT_DEFAULT_DRIVE_ID = os.environ.get('SHAREPOINT_DEFAULT_DRIVE_ID', 'Documents') # 'Documents' es común
MEMORIA_LIST_NAME = os.environ.get('SHAREPOINT_MEMORY_LIST', 'MemoriaPersistenteAsistente') # Nombre configurable para la lista de memoria

# --- Helper Interno para Obtener Site ID ---
def _obtener_site_id_sp(parametros: Dict[str, Any], headers: Dict[str, str]) -> str:
    """
    Obtiene el ID de un sitio SharePoint.
    Prioridad: 'site_id' en params, variable de entorno SHAREPOINT_DEFAULT_SITE_ID, sitio raíz.
    Soporta búsqueda por path relativo (ej. "teams/MiEquipo") o ID directo (guid,hostname,guid).
    """
    site_id_input: Optional[str] = parametros.get("site_id")

    # 1. Si se proporciona ID directo (contiene comas)
    if site_id_input and ',' in site_id_input:
        logger.debug(f"Usando Site ID directo proporcionado: {site_id_input}")
        return site_id_input

    # 2. Si se proporciona path (contiene :) o nombre de host
    if site_id_input and (':' in site_id_input or '.' in site_id_input):
        # Asumimos que es un path como "hostname:/sites/MySite" o "hostname"
        site_path_lookup = site_id_input
        # Si no contiene ':', asumimos que es solo hostname y buscamos el sitio raíz de ese host
        if ':' not in site_path_lookup:
             site_path_lookup = f"{site_path_lookup}:/"

        url = f"{BASE_URL}/sites/{site_path_lookup}?$select=id"
        try:
            logger.debug(f"Buscando Site ID por path/hostname: GET {url}")
            site_data = hacer_llamada_api("GET", url, headers)
            site_id = site_data.get("id")
            if site_id:
                logger.info(f"Site ID encontrado por path/hostname '{site_id_input}': {site_id}")
                return site_id
            else:
                # Esto no debería ocurrir si la llamada fue exitosa (2xx)
                raise ValueError(f"Respuesta inválida de Graph API buscando sitio '{site_id_input}', falta 'id'.")
        except requests.exceptions.RequestException as e:
            # Error 404 (Not Found) es común si el path no existe
            if e.response is not None and e.response.status_code == 404:
                 logger.warning(f"No se encontró sitio por path/hostname '{site_id_input}' (404). Intentando default/raíz.")
            else:
                 logger.warning(f"Error API buscando sitio por path/hostname '{site_id_input}': {e}. Intentando default/raíz.")
        except Exception as e:
            logger.warning(f"Error inesperado buscando sitio por path/hostname '{site_id_input}': {e}. Intentando default/raíz.")

    # 3. Usar variable de entorno si existe
    if SHAREPOINT_DEFAULT_SITE_ID:
        logger.debug(f"Usando Site ID por defecto de variable de entorno: {SHAREPOINT_DEFAULT_SITE_ID}")
        return SHAREPOINT_DEFAULT_SITE_ID

    # 4. Obtener el sitio raíz del tenant
    url = f"{BASE_URL}/sites/root?$select=id"
    try:
        logger.debug(f"Obteniendo sitio raíz SP del tenant: GET {url}")
        site_data = hacer_llamada_api("GET", url, headers)
        site_id = site_data.get("id")
        if not site_id:
            # Esto sería muy raro si la llamada fue exitosa
            raise ValueError("Respuesta de sitio raíz inválida, falta 'id'.")
        logger.info(f"Site ID raíz del tenant obtenido: {site_id}")
        return site_id
    except Exception as e:
        logger.critical(f"Fallo crítico al obtener Site ID (ni input, ni default, ni raíz funcionaron): {e}", exc_info=True)
        raise ValueError(f"No se pudo determinar el Site ID de SharePoint: {e}") from e

# --- Helpers Internos para Endpoints de Drive/Items ---
# Estos solo construyen URLs, no hacen llamadas API
def _get_sp_drive_endpoint(site_id: str, drive_id_or_name: Optional[str] = None) -> str:
    """Construye la URL base para un Drive específico dentro de un Sitio."""
    # Usa el drive_id proporcionado, o el default de env var, o 'Documents' como último recurso
    target_drive = drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
    return f"{BASE_URL}/sites/{site_id}/drives/{target_drive}"

def _get_sp_item_path_endpoint(site_id: str, item_path: str, drive_id_or_name: Optional[str] = None) -> str:
    """Construye la URL para un item específico por path dentro de un Drive."""
    drive_endpoint = _get_sp_drive_endpoint(site_id, drive_id_or_name)
    # Limpiar y asegurar que el path empiece con '/'
    safe_path = item_path.strip()
    if not safe_path: # Si el path es vacío, apuntar a la raíz
        safe_path = '/'
    if not safe_path.startswith('/'):
        safe_path = '/' + safe_path

    # Si el path es solo '/', apunta a la raíz del drive
    if safe_path == '/':
        return f"{drive_endpoint}/root"
    else:
        # Para otros paths, se usa el formato /root:/path/to/item
        return f"{drive_endpoint}/root:{safe_path}"

def _get_drive_id(headers: Dict[str, str], site_id: str, drive_id_or_name: Optional[str] = None) -> str:
    """Obtiene el ID real de un Drive (biblioteca) usando su nombre o ID."""
    drive_endpoint = _get_sp_drive_endpoint(site_id, drive_id_or_name)
    url = f"{drive_endpoint}?$select=id" # Solo necesitamos el ID
    try:
        logger.debug(f"Obteniendo Drive ID para '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'}': GET {url}")
        drive_data = hacer_llamada_api("GET", url, headers)
        actual_drive_id = drive_data.get('id')
        if not actual_drive_id:
            raise ValueError("Respuesta inválida, no se pudo obtener 'id' del drive.")
        logger.info(f"Drive ID obtenido: {actual_drive_id}")
        return actual_drive_id
    except Exception as e:
        logger.error(f"Error API obteniendo Drive ID para '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}': {e}", exc_info=True)
        # Re-lanzar con mensaje claro
        raise Exception(f"Error obteniendo Drive ID para biblioteca '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}': {e}") from e


# ============================================
# ==== FUNCIONES DE ACCIÓN PARA LISTAS SP ====
# ============================================
# Todas usan la firma (parametros: Dict[str, Any], headers: Dict[str, str])
# y llaman a hacer_llamada_api

def crear_lista(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea una nueva lista en SharePoint con columnas personalizadas.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_lista'.
                                     Opcional: 'site_id', 'columnas' (List[Dict]).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: La información de la lista creada devuelta por Graph API.
    """
    nombre_lista: Optional[str] = parametros.get("nombre_lista")
    columnas: Optional[List[Dict[str, Any]]] = parametros.get("columnas") # Lista de dicts

    if not nombre_lista:
        raise ValueError("Parámetro 'nombre_lista' es requerido.")
    if columnas and not isinstance(columnas, list):
         raise ValueError("Parámetro 'columnas' debe ser una lista de diccionarios.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    url = f"{BASE_URL}/sites/{target_site_id}/lists"

    # Graph requiere que 'Title' exista, aunque no se incluya explícitamente en 'columnas'
    # Asegurémonos de que no intentamos añadir 'Title' si ya viene en 'columnas'
    columnas_final = []
    has_title = False
    if columnas:
        for col in columnas:
            if col.get("name", "").lower() == "title":
                has_title = True
            columnas_final.append(col)
    # No es necesario añadir Title manualmente, Graph lo maneja.

    body = {
        "displayName": nombre_lista,
        "columns": columnas_final, # Enviar las columnas proporcionadas
        "list": {"template": "genericList"} # Template estándar
    }
    logger.info(f"Creando lista SP '{nombre_lista}' en sitio {target_site_id}")
    return hacer_llamada_api("POST", url, headers, json_data=body)


def listar_listas(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista las listas del sitio especificado.

    Args:
        parametros (Dict[str, Any]): Opcional: 'site_id', 'select' (campos a devolver).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Respuesta de Graph API, usualmente {'value': [...]}.
    """
    select: str = parametros.get("select", "id,name,displayName,webUrl") # Campos por defecto

    target_site_id = _obtener_site_id_sp(parametros, headers)
    url = f"{BASE_URL}/sites/{target_site_id}/lists"
    params_query = {"$select": select} if select else None

    logger.info(f"Listando listas SP del sitio {target_site_id} (campos: {select})")
    return hacer_llamada_api("GET", url, headers, params=params_query)


def agregar_elemento_lista(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Agrega un elemento a una lista de SharePoint.

    Args:
        parametros (Dict[str, Any]): Debe contener 'lista_id_o_nombre' y 'datos_campos' (dict).
                                     Opcional: 'site_id'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El item creado devuelto por Graph API.
    """
    lista_id_o_nombre: Optional[str] = parametros.get("lista_id_o_nombre")
    datos_campos: Optional[Dict[str, Any]] = parametros.get("datos_campos")

    if not lista_id_o_nombre: raise ValueError("Parámetro 'lista_id_o_nombre' es requerido.")
    if not datos_campos or not isinstance(datos_campos, dict):
        raise ValueError("Parámetro 'datos_campos' (diccionario) es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    # Graph API requiere que los campos estén dentro de un objeto 'fields'
    body = {"fields": datos_campos}
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items"

    logger.info(f"Agregando elemento a lista SP '{lista_id_o_nombre}' en sitio {target_site_id}")
    return hacer_llamada_api("POST", url, headers, json_data=body)


def listar_elementos_lista(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista elementos de una lista, manejando paginación.

    Args:
        parametros (Dict[str, Any]): Debe contener 'lista_id_o_nombre'.
                                     Opcional: 'site_id', 'expand_fields' (bool, default True),
                                     'top' (int, default 100), 'filter_query', 'select', 'order_by'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Un diccionario {'value': [lista_completa_de_items]}.
    """
    lista_id_o_nombre: Optional[str] = parametros.get("lista_id_o_nombre")
    expand_fields: bool = parametros.get("expand_fields", True)
    top: int = int(parametros.get("top", 100)) # Asegurar que sea int
    filter_query: Optional[str] = parametros.get("filter_query")
    select: Optional[str] = parametros.get("select")
    order_by: Optional[str] = parametros.get("order_by")

    if not lista_id_o_nombre: raise ValueError("Parámetro 'lista_id_o_nombre' es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    url_base = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items"

    # Construir parámetros de query iniciales
    params_query: Dict[str, Any] = {'$top': min(top, 999)} # Graph limita top a 999 usualmente
    if expand_fields:
        # Si se pide 'select', Graph a veces requiere expandir 'fields' explícitamente
        if select and 'fields/' in select:
            params_query['$expand'] = 'fields($select=' + ','.join(s.split('/')[1] for s in select.split(',') if s.startswith('fields/')) + ')'
            # Quitar 'fields/' del select principal si se expandió
            select_final = ','.join(s for s in select.split(',') if not s.startswith('fields/'))
            if select_final: params_query['$select'] = select_final
            elif '$select' in params_query: del params_query['$select'] # Eliminar si quedó vacío
        else:
             params_query['$expand'] = 'fields' # Expandir todos los campos si no hay select específico de fields

    if filter_query: params_query['$filter'] = filter_query
    if select and '$select' not in params_query: params_query['$select'] = select # Añadir select si no se manejó con expand
    if order_by: params_query['$orderby'] = order_by

    all_items: List[Dict[str, Any]] = []
    current_url: Optional[str] = url_base # URL para la primera llamada

    page_count = 0
    max_pages = 100 # Límite de seguridad para evitar bucles infinitos

    try:
        while current_url and page_count < max_pages:
            page_count += 1
            logger.info(f"Listando elementos SP lista '{lista_id_o_nombre}', Página: {page_count}")

            # Usar el helper centralizado para cada página
            # Los parámetros de query solo se pasan en la primera llamada (url_base)
            # Las llamadas siguientes usan la URL completa de @odata.nextLink que ya incluye los params
            current_params = params_query if page_count == 1 else None
            data = hacer_llamada_api("GET", current_url, headers, params=current_params)

            if data: # Hacer llamada puede devolver None si hay error o 204
                page_items = data.get('value', [])
                all_items.extend(page_items)
                current_url = data.get('@odata.nextLink') # Obtener siguiente link
                if not current_url:
                    logger.debug("No hay '@odata.nextLink', se termina paginación.")
                    break # Salir si no hay más páginas
            else:
                 logger.warning(f"La llamada a {current_url} devolvió None o vacío. Terminando paginación.")
                 break # Salir si la llamada falla o no devuelve datos

        if page_count >= max_pages:
             logger.warning(f"Se alcanzó el límite de {max_pages} páginas al listar elementos de '{lista_id_o_nombre}'. Puede haber más resultados.")

        logger.info(f"Total elementos SP lista '{lista_id_o_nombre}': {len(all_items)}")
        # Devolver siempre la estructura {'value': [...]}
        return {'value': all_items}

    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_elementos_lista (SP) página {page_count}: {e}", exc_info=True)
        raise Exception(f"Error API listando elementos SP: {e}") from e
    except Exception as e:
        logger.error(f"Error inesperado en listar_elementos_lista (SP) página {page_count}: {e}", exc_info=True)
        raise


def actualizar_elemento_lista(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Actualiza campos de un item de lista. Soporta ETag para concurrencia.

    Args:
        parametros (Dict[str, Any]): Debe contener 'lista_id_o_nombre', 'item_id',
                                     'nuevos_valores_campos' (dict).
                                     Opcional: 'site_id', '@odata.etag' dentro de nuevos_valores.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Los campos actualizados devueltos por Graph API.
    """
    lista_id_o_nombre: Optional[str] = parametros.get("lista_id_o_nombre")
    item_id: Optional[str] = parametros.get("item_id")
    nuevos_valores_campos: Optional[Dict[str, Any]] = parametros.get("nuevos_valores_campos")

    if not lista_id_o_nombre: raise ValueError("Parámetro 'lista_id_o_nombre' es requerido.")
    if not item_id: raise ValueError("Parámetro 'item_id' es requerido.")
    if not nuevos_valores_campos or not isinstance(nuevos_valores_campos, dict):
        raise ValueError("Parámetro 'nuevos_valores_campos' (diccionario) es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items/{item_id}/fields"

    # Extraer ETag si viene en los datos y añadirlo a headers
    # Copiar headers para no modificar el original
    current_headers = headers.copy()
    # Crear copia de valores para no modificar el dict original de params
    body_data = nuevos_valores_campos.copy()
    etag = body_data.pop('@odata.etag', None) # Quitar etag del cuerpo si existe
    if etag:
        current_headers['If-Match'] = etag
        logger.debug(f"Usando ETag '{etag}' para actualización concurrente.")

    logger.info(f"Actualizando elemento SP '{item_id}' en lista '{lista_id_o_nombre}'")
    # Usar el helper centralizado con PATCH
    return hacer_llamada_api("PATCH", url, current_headers, json_data=body_data)


def eliminar_elemento_lista(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina un item de lista. Soporta ETag para concurrencia.

    Args:
        parametros (Dict[str, Any]): Debe contener 'lista_id_o_nombre', 'item_id'.
                                     Opcional: 'site_id', 'etag'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de eliminación.
    """
    lista_id_o_nombre: Optional[str] = parametros.get("lista_id_o_nombre")
    item_id: Optional[str] = parametros.get("item_id")
    etag: Optional[str] = parametros.get("etag") # ETag puede venir como param separado

    if not lista_id_o_nombre: raise ValueError("Parámetro 'lista_id_o_nombre' es requerido.")
    if not item_id: raise ValueError("Parámetro 'item_id' es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{lista_id_o_nombre}/items/{item_id}"

    current_headers = headers.copy()
    if etag:
        current_headers['If-Match'] = etag
        logger.debug(f"Usando ETag '{etag}' para eliminación concurrente.")
    else:
        logger.warning(f"Eliminando item SP {item_id} sin ETag. Podría fallar si fue modificado.")

    logger.info(f"Eliminando elemento SP '{item_id}' de lista '{lista_id_o_nombre}'")
    # Hacer llamada devuelve None en caso de éxito 204
    hacer_llamada_api("DELETE", url, current_headers)
    # Devolver una confirmación explícita
    return {"status": "Eliminado", "item_id": item_id, "lista": lista_id_o_nombre}


# ========================================================
# ==== FUNCIONES DE ACCIÓN PARA DOCUMENTOS (DRIVES) ====
# ========================================================

def listar_documentos_biblioteca(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista documentos y carpetas en una biblioteca/carpeta, manejando paginación.

    Args:
        parametros (Dict[str, Any]): Opcional: 'site_id', 'biblioteca' (nombre o ID),
                                     'ruta_carpeta' (default '/'), 'top' (int, default 100).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Un diccionario {'value': [lista_completa_de_items]}.
    """
    biblioteca: Optional[str] = parametros.get("biblioteca") # Puede ser nombre o ID
    ruta_carpeta: str = parametros.get("ruta_carpeta", '/')
    top: int = int(parametros.get("top", 100))

    target_site_id = _obtener_site_id_sp(parametros, headers)
    # Obtener endpoint del item (puede ser carpeta o raíz)
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_carpeta, biblioteca)
    # La URL para listar hijos es /children
    url_base = f"{item_endpoint}/children"
    params_query = {'$top': min(top, 999)} # Limitar top

    all_files: List[Dict[str, Any]] = []
    current_url: Optional[str] = url_base
    page_count = 0
    max_pages = 100 # Límite de seguridad

    try:
        while current_url and page_count < max_pages:
            page_count += 1
            target_drive_name = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'
            logger.info(f"Listando docs SP biblioteca '{target_drive_name}', Ruta: '{ruta_carpeta}', Página: {page_count}")

            current_params = params_query if page_count == 1 else None
            # Usar helper centralizado
            data = hacer_llamada_api("GET", current_url, headers, params=current_params)

            if data:
                page_items = data.get('value', [])
                all_files.extend(page_items)
                current_url = data.get('@odata.nextLink')
                if not current_url: break
            else:
                 logger.warning(f"Llamada a {current_url} para listar docs devolvió None/vacío.")
                 break

        if page_count >= max_pages:
             logger.warning(f"Se alcanzó límite de {max_pages} páginas listando docs en '{ruta_carpeta}'.")

        logger.info(f"Total docs/carpetas SP encontrados: {len(all_files)}")
        return {'value': all_files}

    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_documentos_biblioteca (SP) página {page_count}: {e}", exc_info=True)
        raise Exception(f"Error API listando documentos SP: {e}") from e
    except Exception as e:
        logger.error(f"Error inesperado en listar_documentos_biblioteca (SP) página {page_count}: {e}", exc_info=True)
        raise


def subir_documento(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Sube un documento a una biblioteca/carpeta. Maneja archivos > 4MB (requiere lógica adicional no implementada aquí).

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo', 'contenido_bytes'.
                                     Opcional: 'site_id', 'biblioteca', 'ruta_carpeta_destino' (default '/'),
                                     'conflict_behavior' ('rename', 'replace', 'fail', default 'rename').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos del archivo subido.
    """
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo")
    contenido_bytes: Optional[bytes] = parametros.get("contenido_bytes") # Espera bytes
    biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta_destino: str = parametros.get("ruta_carpeta_destino", '/')
    conflict_behavior: str = parametros.get("conflict_behavior", "rename")

    if not nombre_archivo: raise ValueError("Parámetro 'nombre_archivo' es requerido.")
    if contenido_bytes is None or not isinstance(contenido_bytes, bytes): # Validar que sean bytes
        raise ValueError("Parámetro 'contenido_bytes' (bytes) es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'

    # Construir path relativo al root del drive
    target_folder_path = ruta_carpeta_destino.strip('/')
    target_file_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"

    # Endpoint para subir contenido
    item_endpoint = _get_sp_item_path_endpoint(target_site_id, target_file_path, target_drive)
    url = f"{item_endpoint}:/content"
    params_query = {"@microsoft.graph.conflictBehavior": conflict_behavior}

    # Headers específicos para subida de contenido
    upload_headers = headers.copy()
    # Determinar Content-Type sería ideal, pero octet-stream es genérico
    upload_headers['Content-Type'] = 'application/octet-stream'

    file_size_mb = len(contenido_bytes) / (1024 * 1024)
    logger.info(f"Subiendo doc SP '{nombre_archivo}' ({file_size_mb:.2f} MB) a '{ruta_carpeta_destino}' con conflict='{conflict_behavior}'")

    # --- Lógica de Subida ---
    # Graph API tiene límite de 4MB para subida simple PUT.
    # Para archivos más grandes, se necesita una sesión de carga (uploadSession).
    if file_size_mb > 4.0:
        # --- INICIO: Lógica de Sesión de Carga (Simplificada) ---
        # 1. Crear sesión de carga
        create_session_url = f"{item_endpoint}:/createUploadSession"
        session_body = {
            "item": {
                "@microsoft.graph.conflictBehavior": conflict_behavior
                # Podrías añadir más metadatos aquí si es necesario
            }
        }
        try:
            logger.info(f"Archivo > 4MB. Creando sesión de carga para '{nombre_archivo}'...")
            session_info = hacer_llamada_api("POST", create_session_url, headers, json_data=session_body)
            upload_url = session_info.get("uploadUrl")
            if not upload_url:
                raise ValueError("No se pudo obtener 'uploadUrl' de la sesión de carga.")
            logger.info(f"Sesión de carga creada. URL: {upload_url[:50]}...")

            # 2. Subir fragmentos (Chunks) - Ejemplo simple, requiere manejo de errores y reintentos
            # Graph recomienda chunks de 5-10 MB, múltiplos de 320 KiB.
            chunk_size = 5 * 1024 * 1024 # 5 MB
            start_byte = 0
            total_bytes = len(contenido_bytes)
            while start_byte < total_bytes:
                end_byte = min(start_byte + chunk_size - 1, total_bytes - 1)
                chunk_data = contenido_bytes[start_byte : end_byte + 1]
                content_range = f"bytes {start_byte}-{end_byte}/{total_bytes}"
                chunk_headers = {
                    'Content-Length': str(len(chunk_data)),
                    'Content-Range': content_range
                    # No necesita Authorization ni Content-Type aquí
                }
                logger.debug(f"Subiendo chunk: {content_range}")
                # Usar requests directamente para PUT a uploadUrl (no necesita auth header)
                # Aumentar timeout para chunks grandes
                chunk_timeout = max(GRAPH_API_TIMEOUT, int(file_size_mb * 5)) # Timeout más largo
                chunk_response = requests.put(upload_url, headers=chunk_headers, data=chunk_data, timeout=chunk_timeout)
                chunk_response.raise_for_status() # Lanza error si falla la subida del chunk
                start_byte = end_byte + 1

            # La última respuesta (201 Created o 200 OK) contiene los metadatos del archivo final
            logger.info(f"Doc SP '{nombre_archivo}' subido exitosamente mediante sesión de carga.")
            return chunk_response.json() # Devuelve los metadatos del archivo

        except requests.exceptions.RequestException as e:
            logger.error(f"Error Request durante sesión de carga para '{nombre_archivo}': {e}", exc_info=True)
            # Podríamos intentar cancelar la sesión si falla
            raise Exception(f"Error API durante sesión de carga: {e}") from e
        except Exception as e:
            logger.error(f"Error inesperado durante sesión de carga para '{nombre_archivo}': {e}", exc_info=True)
            raise
        # --- FIN: Lógica de Sesión de Carga ---
    else:
        # --- Subida Simple (<= 4MB) ---
        try:
             # Usar el helper centralizado pasando los bytes en 'data'
             # Timeout podría necesitar ajuste incluso para <4MB si la red es lenta
             simple_upload_timeout = max(GRAPH_API_TIMEOUT, int(file_size_mb * 10))
             resultado = hacer_llamada_api(
                 metodo="PUT",
                 url=url,
                 headers=upload_headers,
                 params=params_query,
                 data=contenido_bytes, # Pasar bytes aquí
                 timeout=simple_upload_timeout,
                 expect_json=True # Esperamos los metadatos del archivo
             )
             logger.info(f"Doc SP '{nombre_archivo}' subido (subida simple). ID: {resultado.get('id')}")
             return resultado
        except requests.exceptions.RequestException as e:
            logger.error(f"Error Request en subida simple de '{nombre_archivo}': {e}", exc_info=True)
            raise Exception(f"Error API subiendo documento (simple): {e}") from e
        except Exception as e:
            logger.error(f"Error inesperado en subida simple de '{nombre_archivo}': {e}", exc_info=True)
            raise


def eliminar_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina un archivo o carpeta de una biblioteca.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo_o_carpeta'.
                                     Opcional: 'site_id', 'biblioteca', 'ruta_carpeta' (default '/').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de eliminación.
    """
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta")
    biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta: str = parametros.get("ruta_carpeta", '/')

    if not nombre_archivo_o_carpeta: raise ValueError("Parámetro 'nombre_archivo_o_carpeta' es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'

    # Construir path relativo al root del drive
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo_o_carpeta}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo_o_carpeta}"

    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = item_endpoint

    logger.info(f"Eliminando archivo/carpeta SP '{item_path}' en biblioteca '{target_drive}'")
    # Usar helper. Devuelve None en éxito 204.
    hacer_llamada_api("DELETE", url, headers)
    return {"status": "Eliminado", "path": item_path}


# ======================================================
# ==== FUNCIONES AVANZADAS DE ARCHIVOS (MOVER/COPIAR) ====
# ======================================================

def crear_carpeta_biblioteca(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea una nueva carpeta en una biblioteca/carpeta padre.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_carpeta'.
                                     Opcional: 'site_id', 'biblioteca', 'ruta_carpeta_padre' (default '/'),
                                     'conflict_behavior' ('rename', 'replace', 'fail', default 'rename').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos de la carpeta creada.
    """
    nombre_carpeta: Optional[str] = parametros.get("nombre_carpeta")
    biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta_padre: str = parametros.get("ruta_carpeta_padre", '/')
    conflict_behavior: str = parametros.get("conflict_behavior", "rename")

    if not nombre_carpeta: raise ValueError("Parámetro 'nombre_carpeta' es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'

    # Endpoint de la carpeta padre donde se creará la nueva carpeta
    parent_folder_endpoint = _get_sp_item_path_endpoint(target_site_id, ruta_carpeta_padre, target_drive)
    url = f"{parent_folder_endpoint}/children" # Crear item en los hijos del padre

    body = {
        "name": nombre_carpeta,
        "folder": {}, # Indica que es una carpeta
        "@microsoft.graph.conflictBehavior": conflict_behavior
    }
    logger.info(f"Creando carpeta SP '{nombre_carpeta}' en '{ruta_carpeta_padre}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)


def mover_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Mueve un archivo o carpeta a una nueva ubicación (dentro del mismo Drive).

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo_o_carpeta', 'nueva_ruta_carpeta_padre'.
                                     Opcional: 'site_id', 'biblioteca', 'ruta_carpeta_origen' (default '/'),
                                     'nuevo_nombre' (para renombrar al mover).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos del item movido/renombrado.
    """
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta")
    nueva_ruta_carpeta_padre: Optional[str] = parametros.get("nueva_ruta_carpeta_padre")
    biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta_origen: str = parametros.get("ruta_carpeta_origen", '/')
    nuevo_nombre: Optional[str] = parametros.get("nuevo_nombre")

    if not nombre_archivo_o_carpeta: raise ValueError("Parámetro 'nombre_archivo_o_carpeta' es requerido.")
    if nueva_ruta_carpeta_padre is None: raise ValueError("Parámetro 'nueva_ruta_carpeta_padre' es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    target_drive_name = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'

    # Construir path de origen
    target_folder_path_origen = ruta_carpeta_origen.strip('/')
    item_path_origen = f"/{nombre_archivo_o_carpeta}" if not target_folder_path_origen else f"/{target_folder_path_origen}/{nombre_archivo_o_carpeta}"

    # Endpoint del item a mover
    item_endpoint_origen = _get_sp_item_path_endpoint(target_site_id, item_path_origen, target_drive_name)
    url = item_endpoint_origen # La operación PATCH se hace sobre el item mismo

    # Construir referencia a la carpeta padre de destino
    # Necesitamos el ID del Drive para la referencia parentReference
    try:
        actual_drive_id = _get_drive_id(headers, target_site_id, target_drive_name)
    except Exception as drive_err:
        raise Exception(f"Error obteniendo Drive ID para mover: {drive_err}") from drive_err

    parent_dest_path = nueva_ruta_carpeta_padre.strip()
    if not parent_dest_path.startswith('/'): parent_dest_path = '/' + parent_dest_path
    # La referencia al padre usa /drives/{drive-id}/root:/path/to/parent
    parent_path_ref = f"/drives/{actual_drive_id}/root" if parent_dest_path == '/' else f"/drives/{actual_drive_id}/root:{parent_dest_path}"

    body = {
        "parentReference": {
             # Se puede usar 'id' de la carpeta padre o 'path' relativo al drive
            "path": parent_path_ref
            # "id": "ID_CARPETA_DESTINO" # Alternativa si tienes el ID
        }
    }
    # Añadir nuevo nombre si se proporcionó
    body["name"] = nuevo_nombre if nuevo_nombre is not None else nombre_archivo_o_carpeta

    logger.info(f"Moviendo SP '{item_path_origen}' a '{nueva_ruta_carpeta_padre}' (nuevo nombre: {body['name']})")
    # Usar PATCH en el item de origen con la nueva referencia padre y/o nombre
    return hacer_llamada_api("PATCH", url, headers, json_data=body)


def copiar_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Copia un archivo a una nueva ubicación (puede ser otro Drive al que se tenga acceso).
    Esta operación es asíncrona.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo', 'nueva_ruta_carpeta_padre'.
                                     Opcional: 'site_id', 'biblioteca' (origen), 'ruta_carpeta_origen' (default '/'),
                                     'nuevo_nombre_copia', 'drive_id_destino' (si es a otro drive).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Respuesta 202 Accepted con la URL para monitorizar la copia.
    """
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo")
    nueva_ruta_carpeta_padre: Optional[str] = parametros.get("nueva_ruta_carpeta_padre")
    biblioteca: Optional[str] = parametros.get("biblioteca") # Drive origen
    ruta_carpeta_origen: str = parametros.get("ruta_carpeta_origen", '/')
    nuevo_nombre_copia: Optional[str] = parametros.get("nuevo_nombre_copia")
    drive_id_destino: Optional[str] = parametros.get("drive_id_destino") # Opcional, para copiar a otro drive

    if not nombre_archivo: raise ValueError("Parámetro 'nombre_archivo' es requerido.")
    if nueva_ruta_carpeta_padre is None: raise ValueError("Parámetro 'nueva_ruta_carpeta_padre' es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers) # Sitio origen
    target_drive_name_origen = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'

    # Path de origen
    target_folder_path_origen = ruta_carpeta_origen.strip('/')
    item_path_origen = f"/{nombre_archivo}" if not target_folder_path_origen else f"/{target_folder_path_origen}/{nombre_archivo}"

    # Endpoint del item a copiar
    item_endpoint_origen = _get_sp_item_path_endpoint(target_site_id, item_path_origen, target_drive_name_origen)
    url = f"{item_endpoint_origen}/copy" # Endpoint para la acción de copia

    # Referencia a la carpeta padre de destino
    # Si no se da drive_id_destino, se asume el mismo drive de origen
    if not drive_id_destino:
        try:
            drive_id_destino = _get_drive_id(headers, target_site_id, target_drive_name_origen)
        except Exception as drive_err:
            raise Exception(f"Error obteniendo Drive ID de origen para copiar: {drive_err}") from drive_err

    parent_dest_path = nueva_ruta_carpeta_padre.strip()
    if not parent_dest_path.startswith('/'): parent_dest_path = '/' + parent_dest_path
    # La referencia al padre usa /drive/root:/path si es el mismo drive, o /drives/... si es otro
    # Para simplificar, usamos siempre la referencia completa con driveId
    parent_path_ref = f"/drives/{drive_id_destino}/root" if parent_dest_path == '/' else f"/drives/{drive_id_destino}/root:{parent_dest_path}"

    body = {
        "parentReference": {
            "driveId": drive_id_destino,
             # Se puede usar 'id' de la carpeta padre o 'path' relativo al drive
            "path": parent_path_ref
            # "id": "ID_CARPETA_DESTINO" # Alternativa
        },
        # Nombre opcional para la copia
        "name": nuevo_nombre_copia or f"Copia de {nombre_archivo}"
    }

    logger.info(f"Iniciando copia asíncrona SP de '{item_path_origen}' a Drive '{drive_id_destino}', Path: '{nueva_ruta_carpeta_padre}'")

    # La copia devuelve 202 Accepted. Necesitamos el objeto Response para leer el header 'Location'.
    # Llamamos al helper con expect_json=False
    response = hacer_llamada_api("POST", url, headers, json_data=body, expect_json=False)

    # Verificar que la respuesta sea un objeto Response y tenga status 202
    if isinstance(response, requests.Response) and response.status_code == 202:
        monitor_url = response.headers.get('Location')
        logger.info(f"Copia SP '{nombre_archivo}' iniciada. Monitor URL: {monitor_url}")
        # Devolver la información relevante
        return {
            "status": "Copia Iniciada",
            "status_code": response.status_code,
            "monitorUrl": monitor_url,
            "detail": "La copia se realiza en segundo plano. Usa la URL de monitorización para verificar el estado."
        }
    elif isinstance(response, requests.Response):
         # Si la llamada fue exitosa pero no 202 (inesperado para copy)
         logger.error(f"Respuesta inesperada al iniciar copia SP: {response.status_code} {response.reason}. Cuerpo: {response.text[:200]}")
         raise Exception(f"Respuesta inesperada al iniciar copia SP: {response.status_code}")
    else:
         # Si hacer_llamada_api devolvió algo que no es Response (inesperado si expect_json=False)
         logger.error(f"Respuesta inesperada del helper al iniciar copia SP: {type(response)}")
         raise Exception("Error interno al procesar la solicitud de copia.")


# ======================================================
# ==== FUNCIONES DE METADATOS Y CONTENIDO ARCHIVOS ====
# ======================================================

def obtener_metadatos_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene los metadatos de un archivo o carpeta.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo_o_carpeta'.
                                     Opcional: 'site_id', 'biblioteca', 'ruta_carpeta' (default '/').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos del item.
    """
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta")
    biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta: str = parametros.get("ruta_carpeta", '/')

    if not nombre_archivo_o_carpeta: raise ValueError("Parámetro 'nombre_archivo_o_carpeta' es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'

    # Path relativo al root
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo_o_carpeta}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo_o_carpeta}"

    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = item_endpoint # GET en el endpoint del item devuelve sus metadatos

    logger.info(f"Obteniendo metadatos SP '{item_path}'")
    return hacer_llamada_api("GET", url, headers)


def actualizar_metadatos_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Actualiza metadatos de un archivo o carpeta (ej. nombre). Soporta ETag.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo_o_carpeta', 'nuevos_valores' (dict).
                                     Opcional: 'site_id', 'biblioteca', 'ruta_carpeta' (default '/'),
                                     '@odata.etag' dentro de nuevos_valores.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos actualizados.
    """
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta")
    nuevos_valores: Optional[Dict[str, Any]] = parametros.get("nuevos_valores")
    biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta: str = parametros.get("ruta_carpeta", '/')

    if not nombre_archivo_o_carpeta: raise ValueError("Parámetro 'nombre_archivo_o_carpeta' es requerido.")
    if not nuevos_valores or not isinstance(nuevos_valores, dict):
        raise ValueError("Parámetro 'nuevos_valores' (diccionario) es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'

    # Path relativo al root
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo_o_carpeta}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo_o_carpeta}"

    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = item_endpoint # PATCH en el endpoint del item actualiza metadatos

    # Extraer ETag si viene y añadir a headers
    current_headers = headers.copy()
    body_data = nuevos_valores.copy() # Copia para no modificar params
    etag = body_data.pop('@odata.etag', None)
    if etag:
        current_headers['If-Match'] = etag
        logger.debug(f"Usando ETag '{etag}' para actualización de metadatos.")

    logger.info(f"Actualizando metadatos SP '{item_path}'")
    return hacer_llamada_api("PATCH", url, current_headers, json_data=body_data)


def obtener_contenido_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> bytes:
    """
    Descarga el contenido de un archivo.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo'.
                                     Opcional: 'site_id', 'biblioteca', 'ruta_carpeta' (default '/').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        bytes: El contenido binario del archivo.
    """
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo")
    biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta: str = parametros.get("ruta_carpeta", '/')

    if not nombre_archivo: raise ValueError("Parámetro 'nombre_archivo' es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'

    # Path relativo al root
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"

    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = f"{item_endpoint}/content" # Endpoint para descargar contenido

    logger.info(f"Obteniendo contenido SP '{item_path}'")
    # Llamar al helper esperando el objeto Response para acceder a .content
    response = hacer_llamada_api("GET", url, headers, expect_json=False)

    if isinstance(response, requests.Response):
        # raise_for_status ya fue llamado dentro del helper si hubo error 4xx/5xx
        logger.info(f"Contenido SP '{item_path}' obtenido ({len(response.content)} bytes).")
        return response.content
    else:
        # Esto no debería pasar si expect_json=False y no hubo error
        logger.error(f"Respuesta inesperada del helper al obtener contenido: {type(response)}")
        raise Exception("Error interno al obtener contenido del archivo.")


def actualizar_contenido_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Actualiza/Reemplaza el contenido de un archivo existente.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo', 'nuevo_contenido_bytes'.
                                     Opcional: 'site_id', 'biblioteca', 'ruta_carpeta' (default '/').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos del archivo actualizado.
    """
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo")
    nuevo_contenido_bytes: Optional[bytes] = parametros.get("nuevo_contenido_bytes")
    biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta: str = parametros.get("ruta_carpeta", '/')

    if not nombre_archivo: raise ValueError("Parámetro 'nombre_archivo' es requerido.")
    if nuevo_contenido_bytes is None or not isinstance(nuevo_contenido_bytes, bytes):
        raise ValueError("Parámetro 'nuevo_contenido_bytes' (bytes) es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'

    # Path relativo al root
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"

    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = f"{item_endpoint}/content" # PUT en /content reemplaza el contenido

    upload_headers = headers.copy()
    upload_headers['Content-Type'] = 'application/octet-stream'

    file_size_mb = len(nuevo_contenido_bytes) / (1024 * 1024)
    logger.info(f"Actualizando contenido SP '{item_path}' ({file_size_mb:.2f} MB)")

    # Aquí también aplica el límite de 4MB para PUT simple.
    # Si necesitas actualizar archivos grandes, se usa la misma lógica de uploadSession
    # que en subir_documento. Por simplicidad, aquí usamos solo PUT simple.
    if file_size_mb > 4.0:
         logger.warning(f"Actualizando archivo > 4MB ({nombre_archivo}) con PUT simple. Podría fallar. Considera usar sesión de carga.")
         # Podríamos lanzar un error aquí o intentar la subida simple de todas formas.
         # raise ValueError("Actualización de contenido para archivos > 4MB requiere sesión de carga (no implementado en esta función).")

    try:
        # Usar helper con 'data'
        # Timeout necesita ser potencialmente largo
        update_timeout = max(GRAPH_API_TIMEOUT, int(file_size_mb * 10))
        resultado = hacer_llamada_api(
            metodo="PUT",
            url=url,
            headers=upload_headers,
            data=nuevo_contenido_bytes,
            timeout=update_timeout,
            expect_json=True # PUT en /content devuelve metadatos
        )
        logger.info(f"Contenido SP '{item_path}' actualizado exitosamente.")
        return resultado
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request al actualizar contenido de '{item_path}': {e}", exc_info=True)
        raise Exception(f"Error API actualizando contenido: {e}") from e
    except Exception as e:
        logger.error(f"Error inesperado al actualizar contenido de '{item_path}': {e}", exc_info=True)
        raise


def crear_enlace_compartido_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea un enlace para compartir un archivo o carpeta.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo_o_carpeta'.
                                     Opcional: 'site_id', 'biblioteca', 'ruta_carpeta' (default '/'),
                                     'tipo_enlace' ('view', 'edit', default 'view'),
                                     'alcance' ('anonymous', 'organization', 'users', default 'organization').
                                     'password' (string, opcional), 'expirationDateTime' (ISO 8601 string, opcional).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Información del enlace creado (incluye 'link.webUrl').
    """
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta")
    biblioteca: Optional[str] = parametros.get("biblioteca")
    ruta_carpeta: str = parametros.get("ruta_carpeta", '/')
    tipo_enlace: str = parametros.get("tipo_enlace", "view")
    alcance: str = parametros.get("alcance", "organization")
    password: Optional[str] = parametros.get("password")
    expirationDateTime: Optional[str] = parametros.get("expirationDateTime")

    if not nombre_archivo_o_carpeta: raise ValueError("Parámetro 'nombre_archivo_o_carpeta' es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID or 'Documents'

    # Path relativo al root
    target_folder_path = ruta_carpeta.strip('/')
    item_path = f"/{nombre_archivo_o_carpeta}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo_o_carpeta}"

    item_endpoint = _get_sp_item_path_endpoint(target_site_id, item_path, target_drive)
    url = f"{item_endpoint}/createLink" # Endpoint para crear enlace

    body: Dict[str, Any] = {"type": tipo_enlace, "scope": alcance}
    if password: body["password"] = password
    if expirationDateTime: body["expirationDateTime"] = expirationDateTime
    # Podrían añadirse más opciones como 'retainInheritedPermissions'

    logger.info(f"Creando enlace SP (tipo: {tipo_enlace}, alcance: {alcance}) para '{item_path}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)


# ======================================================
# ==== FUNCIONES DE MEMORIA PERSISTENTE (LISTA SP) ====
# ======================================================

def _ensure_memory_list_exists(headers: Dict[str, str], site_id: str) -> bool:
    """Verifica si la lista de memoria existe, la crea si no."""
    try:
        # Intentar obtener la lista por nombre para ver si existe
        list_url = f"{BASE_URL}/sites/{site_id}/lists/{MEMORIA_LIST_NAME}?$select=id"
        hacer_llamada_api("GET", list_url, headers)
        logger.debug(f"Lista de memoria '{MEMORIA_LIST_NAME}' ya existe.")
        return True
    except requests.exceptions.RequestException as e:
        if e.response is not None and e.response.status_code == 404:
            # No existe, intentar crearla
            logger.info(f"Lista de memoria '{MEMORIA_LIST_NAME}' no encontrada. Intentando crearla...")
            columnas = [
                {"name": "SessionID", "text": {}, "indexed": True},
                {"name": "Clave", "text": {}, "indexed": True},
                # Usar texto multilinea para poder guardar JSON como string
                {"name": "Valor", "text": {"allowMultipleLines": True, "textType": "plain"}},
                {"name": "Timestamp", "dateTime": {}, "indexed": True} # Añadir timestamp
            ]
            params_crear = {"nombre_lista": MEMORIA_LIST_NAME, "columnas": columnas, "site_id": site_id}
            try:
                crear_lista(params_crear, headers) # Reutilizar la función de acción
                logger.info(f"Lista de memoria '{MEMORIA_LIST_NAME}' creada exitosamente.")
                return True
            except Exception as create_err:
                logger.critical(f"¡Fallo al crear lista de memoria '{MEMORIA_LIST_NAME}'!: {create_err}", exc_info=True)
                return False # Indicar fallo en la creación
        else:
            # Otro error al verificar existencia
            logger.error(f"Error verificando existencia de lista de memoria '{MEMORIA_LIST_NAME}': {e}", exc_info=True)
            return False # Indicar fallo en la verificación
    except Exception as e:
         logger.error(f"Error inesperado verificando lista de memoria: {e}", exc_info=True)
         return False


def guardar_dato_memoria(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Guarda (o actualiza) un dato clave-valor para una sesión en la lista de memoria.

    Args:
        parametros (Dict[str, Any]): Debe contener 'session_id', 'clave', 'valor'.
                                     Opcional: 'site_id'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El item guardado/actualizado.
    """
    session_id: Optional[str] = parametros.get("session_id")
    clave: Optional[str] = parametros.get("clave")
    valor: Any = parametros.get("valor") # Puede ser cualquier tipo serializable a JSON

    if not session_id: raise ValueError("Parámetro 'session_id' es requerido.")
    if not clave: raise ValueError("Parámetro 'clave' es requerido.")
    if valor is None: raise ValueError("Parámetro 'valor' es requerido (no puede ser None).")

    target_site_id = _obtener_site_id_sp(parametros, headers)

    # Asegurar que la lista exista
    if not _ensure_memory_list_exists(headers, target_site_id):
        raise Exception(f"No se pudo asegurar la existencia de la lista de memoria '{MEMORIA_LIST_NAME}'.")

    # Convertir valor a string (JSON si es dict/list)
    try:
        if isinstance(valor, (dict, list, bool)):
            valor_str = json.dumps(valor)
        elif isinstance(valor, (int, float)):
             valor_str = str(valor) # Guardar números como string también
        elif isinstance(valor, str):
             valor_str = valor
        else:
             # Intentar convertir a string otros tipos, puede fallar
             valor_str = str(valor)
             logger.warning(f"Guardando tipo no estándar '{type(valor)}' como string: {valor_str[:50]}...")
    except Exception as json_err:
        raise ValueError(f"Error al serializar el valor para la clave '{clave}': {json_err}") from json_err

    # Buscar item existente para actualizar (PATCH) o crear (POST)
    filter_query = f"fields/SessionID eq '{session_id}' and fields/Clave eq '{clave}'"
    params_listar = {
        "lista_id_o_nombre": MEMORIA_LIST_NAME,
        "site_id": target_site_id,
        "filter_query": filter_query,
        "top": 1,
        "select": "id,@odata.etag" # Obtener ID y ETag para actualización
    }
    item_id: Optional[str] = None
    item_etag: Optional[str] = None
    try:
        existing_items_data = listar_elementos_lista(params_listar, headers)
        existing_items = existing_items_data.get("value", [])
        if existing_items:
            item_id = existing_items[0].get("id")
            item_etag = existing_items[0].get("@odata.etag")
    except Exception as e:
        # No fallar si la búsqueda falla, intentar crear
        logger.warning(f"Error buscando item existente para memoria ({session_id}/{clave}): {e}. Se intentará crear.")

    # Datos a guardar/actualizar
    datos_campos = {
        "SessionID": session_id, # Asegurar que estos campos también se actualicen si cambian
        "Clave": clave,
        "Valor": valor_str,
        "Timestamp": datetime.utcnow().isoformat() + "Z" # Guardar timestamp UTC
    }

    if item_id:
        # Actualizar existente usando PATCH y ETag si se obtuvo
        logger.info(f"Actualizando dato en memoria: Session={session_id}, Clave={clave}")
        params_actualizar = {
            "lista_id_o_nombre": MEMORIA_LIST_NAME,
            "item_id": item_id,
            "nuevos_valores_campos": datos_campos,
            "site_id": target_site_id
            # Pasar ETag dentro de nuevos_valores para que la función lo maneje
        }
        if item_etag: params_actualizar["nuevos_valores_campos"]["@odata.etag"] = item_etag

        return actualizar_elemento_lista(params_actualizar, headers)
    else:
        # Crear nuevo usando POST
        logger.info(f"Guardando nuevo dato en memoria: Session={session_id}, Clave={clave}")
        params_agregar = {
            "lista_id_o_nombre": MEMORIA_LIST_NAME,
            "datos_campos": datos_campos,
            "site_id": target_site_id
        }
        return agregar_elemento_lista(params_agregar, headers)


def recuperar_datos_sesion(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Recupera todos los datos (clave-valor) asociados a una sesión, ordenados por timestamp descendente.

    Args:
        parametros (Dict[str, Any]): Debe contener 'session_id'. Opcional: 'site_id'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Un diccionario con los datos clave-valor de la sesión.
    """
    session_id: Optional[str] = parametros.get("session_id")
    if not session_id: raise ValueError("Parámetro 'session_id' es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)

    # Asegurar que la lista exista (opcional, podría fallar si no existe)
    if not _ensure_memory_list_exists(headers, target_site_id):
         logger.warning(f"Lista de memoria '{MEMORIA_LIST_NAME}' no encontrada al recuperar datos.")
         return {} # Devolver vacío si la lista no existe

    filter_query = f"fields/SessionID eq '{session_id}'"
    # Seleccionar campos necesarios y ordenar por Timestamp descendente
    select_fields = "id,fields/Clave,fields/Valor,fields/Timestamp"
    order_by = "fields/Timestamp desc"

    logger.info(f"Recuperando datos de memoria para Session={session_id}")
    params_listar = {
        "lista_id_o_nombre": MEMORIA_LIST_NAME,
        "site_id": target_site_id,
        "filter_query": filter_query,
        "expand_fields": True, # Necesario para acceder a fields/*
        "select": select_fields,
        "order_by": order_by,
        "top": 999 # Obtener hasta 999 items (límite práctico sin paginación compleja aquí)
    }
    items_data = listar_elementos_lista(params_listar, headers)

    # Reconstruir diccionario clave-valor, intentando decodificar JSON
    memoria: Dict[str, Any] = {}
    for item in items_data.get("value", []):
        fields = item.get("fields", {})
        clave = fields.get("Clave")
        valor_str = fields.get("Valor")
        timestamp = fields.get("Timestamp") # Podríamos añadirlo al valor si es útil

        if clave and valor_str:
            # Solo añadir la clave si no existe ya (por el order_by, el primero es el más reciente)
            if clave not in memoria:
                try:
                    # Intentar decodificar JSON
                    memoria[clave] = json.loads(valor_str)
                except json.JSONDecodeError:
                    # Si no es JSON, guardar como string
                    memoria[clave] = valor_str
                except Exception as parse_err:
                     logger.warning(f"Error parseando valor para clave '{clave}' (Session: {session_id}): {parse_err}. Guardando como string.")
                     memoria[clave] = valor_str # Guardar como string en caso de error

    logger.info(f"Recuperados {len(memoria)} datos únicos para Session={session_id}")
    return memoria


def eliminar_dato_memoria(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina un dato específico (por clave) de una sesión en memoria.

    Args:
        parametros (Dict[str, Any]): Debe contener 'session_id', 'clave'. Opcional: 'site_id'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Estado de la operación.
    """
    session_id: Optional[str] = parametros.get("session_id")
    clave: Optional[str] = parametros.get("clave")

    if not session_id: raise ValueError("Parámetro 'session_id' es requerido.")
    if not clave: raise ValueError("Parámetro 'clave' es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)

    # Asegurar que la lista exista
    if not _ensure_memory_list_exists(headers, target_site_id):
         logger.warning(f"Lista de memoria '{MEMORIA_LIST_NAME}' no encontrada al eliminar dato.")
         return {"status": "Lista no encontrada"}

    # Buscar el item a eliminar
    filter_query = f"fields/SessionID eq '{session_id}' and fields/Clave eq '{clave}'"
    params_listar = {
        "lista_id_o_nombre": MEMORIA_LIST_NAME,
        "site_id": target_site_id,
        "filter_query": filter_query,
        "top": 1,
        "select": "id" # Solo necesitamos el ID
    }
    item_id: Optional[str] = None
    try:
        existing_items_data = listar_elementos_lista(params_listar, headers)
        existing_items = existing_items_data.get("value", [])
        if existing_items:
            item_id = existing_items[0].get("id")
    except Exception as e:
        logger.error(f"Error buscando item a eliminar ({session_id}/{clave}): {e}", exc_info=True)
        raise Exception(f"Error buscando item a eliminar: {e}") from e

    if item_id:
        logger.info(f"Eliminando dato de memoria: Session={session_id}, Clave={clave}")
        params_eliminar = {
            "lista_id_o_nombre": MEMORIA_LIST_NAME,
            "item_id": item_id,
            "site_id": target_site_id
            # Podríamos pasar ETag si lo obtuviéramos en la búsqueda
        }
        return eliminar_elemento_lista(params_eliminar, headers) # Devuelve dict de confirmación
    else:
        logger.warning(f"No se encontró dato para eliminar: Session={session_id}, Clave={clave}")
        return {"status": "No encontrado", "session_id": session_id, "clave": clave}


def eliminar_memoria_sesion(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina TODOS los datos asociados a una sesión específica.
    ADVERTENCIA: Puede ser lento si hay muchos items. Considerar $batch para >20 items.

    Args:
        parametros (Dict[str, Any]): Debe contener 'session_id'. Opcional: 'site_id'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Resumen de la operación.
    """
    session_id: Optional[str] = parametros.get("session_id")
    if not session_id: raise ValueError("Parámetro 'session_id' es requerido.")

    target_site_id = _obtener_site_id_sp(parametros, headers)

    # Asegurar que la lista exista
    if not _ensure_memory_list_exists(headers, target_site_id):
         logger.warning(f"Lista de memoria '{MEMORIA_LIST_NAME}' no encontrada al eliminar sesión.")
         return {"status": "Lista no encontrada", "items_eliminados": 0}

    # Listar TODOS los items de la sesión (solo IDs)
    filter_query = f"fields/SessionID eq '{session_id}'"
    logger.info(f"Listando TODOS los items de memoria para eliminar Session={session_id}")
    params_listar = {
        "lista_id_o_nombre": MEMORIA_LIST_NAME,
        "site_id": target_site_id,
        "filter_query": filter_query,
        "select": "id",
        "top": 999 # Limitar por llamada, necesita paginación real para >999
    }
    # TODO: Implementar paginación aquí si se esperan >999 items por sesión
    items_data = listar_elementos_lista(params_listar, headers)
    items_to_delete = items_data.get("value", [])
    item_ids_to_delete = [item.get("id") for item in items_to_delete if item.get("id")]

    if not item_ids_to_delete:
        logger.info(f"No se encontraron datos para eliminar para Session={session_id}")
        return {"status": "Sin datos", "items_eliminados": 0, "session_id": session_id}

    logger.info(f"Se eliminarán {len(item_ids_to_delete)} datos para Session={session_id}")

    # --- Lógica de Eliminación ---
    # Opción 1: Eliminar uno por uno (simple pero lento)
    count_deleted = 0
    count_failed = 0
    for item_id in item_ids_to_delete:
        try:
            params_eliminar = {
                "lista_id_o_nombre": MEMORIA_LIST_NAME,
                "item_id": item_id,
                "site_id": target_site_id
            }
            eliminar_elemento_lista(params_eliminar, headers)
            count_deleted += 1
        except Exception as del_err:
            logger.error(f"Error eliminando item {item_id} de memoria para sesión {session_id}: {del_err}")
            count_failed += 1
            # ¿Parar o continuar? Decidimos continuar.

    # Opción 2: Usar $batch (más eficiente para >20 items, más complejo de implementar)
    # Si len(item_ids_to_delete) > 20:
    #    logger.info("Usando $batch para eliminación masiva...")
    #    # Construir payload de batch request
    #    batch_requests = []
    #    for i, item_id in enumerate(item_ids_to_delete):
    #        batch_req = {
    #            "id": str(i + 1),
    #            "method": "DELETE",
    #            "url": f"/sites/{target_site_id}/lists/{MEMORIA_LIST_NAME}/items/{item_id}"
    #            # Podría necesitar If-Match header si se requiere ETag
    #        }
    #        batch_requests.append(batch_req)
    #    batch_payload = {"requests": batch_requests}
    #    batch_url = f"{BASE_URL}/$batch"
    #    try:
    #        batch_response_data = hacer_llamada_api("POST", batch_url, headers, json_data=batch_payload)
    #        # Procesar respuestas individuales del batch
    #        responses = batch_response_data.get("responses", [])
    #        count_deleted = sum(1 for r in responses if r.get("status") == 204)
    #        count_failed = len(responses) - count_deleted
    #        if count_failed > 0: logger.warning(f"{count_failed} errores en eliminación batch.")
    #    except Exception as batch_err:
    #        logger.error(f"Error ejecutando $batch para eliminar memoria: {batch_err}", exc_info=True)
    #        # Podríamos intentar la eliminación individual como fallback?
    #        raise Exception(f"Error en $batch de eliminación: {batch_err}") from batch_err

    logger.info(f"Eliminación memoria sesión {session_id}: {count_deleted} exitosos, {count_failed} fallidos.")
    return {
        "status": "Completado" if count_failed == 0 else "Completado con errores",
        "items_eliminados": count_deleted,
        "items_fallidos": count_failed,
        "session_id": session_id
    }


# ======================================================
# ==== FUNCIONES ADICIONALES (EJ: EXPORTAR)         ====
# ======================================================

def exportar_datos_lista(parametros: Dict[str, Any], headers: Dict[str, str]) -> Union[Dict[str, Any], str]:
    """
    Exporta los datos de una lista en formato JSON o CSV.
    NOTA: Actualmente limitado por la paginación de listar_elementos_lista (ej. 999 items).

    Args:
        parametros (Dict[str, Any]): Debe contener 'lista_id_o_nombre'.
                                     Opcional: 'site_id', 'formato' ('json' o 'csv', default 'json').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Union[Dict[str, Any], str]: Los datos en formato JSON (dict) o CSV (string).
    """
    lista_id_o_nombre: Optional[str] = parametros.get("lista_id_o_nombre")
    formato: str = parametros.get("formato", "json").lower()

    if not lista_id_o_nombre: raise ValueError("Parámetro 'lista_id_o_nombre' es requerido.")
    if formato not in ["json", "csv"]: raise ValueError("Formato no soportado. Use 'json' o 'csv'.")

    target_site_id = _obtener_site_id_sp(parametros, headers)
    logger.info(f"Exportando datos de lista '{lista_id_o_nombre}' en formato {formato}")

    # Obtener todos los items (con límite actual de paginación)
    # TODO: Implementar paginación completa aquí si se necesitan >999 items
    params_listar = {
        "lista_id_o_nombre": lista_id_o_nombre,
        "site_id": target_site_id,
        "expand_fields": True,
        "top": 999 # Limitar a 999 por ahora
    }
    items_data = listar_elementos_lista(params_listar, headers)
    items = items_data.get("value", [])

    if not items:
        logger.warning(f"No se encontraron items para exportar en lista '{lista_id_o_nombre}'.")
        return {"value": []} if formato == "json" else ""

    # Extraer solo los campos (fields), incluyendo 'id' del item si es útil
    field_data = []
    for item in items:
        fields = item.get("fields", {})
        fields["_ItemID_"] = item.get("id") # Añadir ID del item por si acaso
        field_data.append(fields)

    if formato == "json":
        logger.info(f"Exportando {len(field_data)} items como JSON.")
        return {"value": field_data} # Devolver como dict JSON
    else: # formato == "csv"
        logger.info(f"Exportando {len(field_data)} items como CSV.")
        if not field_data: return "" # String vacío si no hay datos

        output = StringIO()
        # Usar las claves del primer item como cabeceras CSV (asume consistencia)
        # Mover '_ItemID_' al principio si existe
        field_names = list(field_data[0].keys())
        if "_ItemID_" in field_names:
            field_names.insert(0, field_names.pop(field_names.index("_ItemID_")))

        writer = csv.DictWriter(output, fieldnames=field_names, extrasaction='ignore', quoting=csv.QUOTE_MINIMAL)
        writer.writeheader()
        writer.writerows(field_data)
        csv_content = output.getvalue()
        output.close()
        return csv_content # Devolver string CSV

# --- FIN DEL MÓDULO actions/sharepoint.py ---


