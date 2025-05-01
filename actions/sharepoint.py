import logging
import requests
import json # Para manejo de errores
import os # Para os.path.join
from typing import Dict, List, Optional, Union, Any

# Usar el logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar constantes globales desde __init__.py
try:
    from .. import BASE_URL, GRAPH_API_TIMEOUT
    # Obtener defaults configurados en la función principal (leídos de env vars allí)
    from .. import SHAREPOINT_DEFAULT_SITE_ID, SHAREPOINT_DEFAULT_DRIVE_ID
except ImportError:
    # Fallback si se ejecuta standalone
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_TIMEOUT = 45
    SHAREPOINT_DEFAULT_SITE_ID = None # No hay default si no se carga
    SHAREPOINT_DEFAULT_DRIVE_ID = "Documents" # Default común
    logger.warning("No se pudo importar BASE_URL/TIMEOUT/Defaults desde el padre, usando defaults.")

# --- Constantes Específicas (si aplica) ---
# MAX_PAGINATION_SIZE = 999 # Se puede definir aquí o usar el default en llamadas

# --- Helper para obtener Site ID (Refactorizado) ---
# Ya no cachea globalmente, acepta headers
def _obtener_site_id_sp(headers: Dict[str, str], site_id_o_url: Optional[str] = None) -> str:
    """
    Obtiene el ID de un sitio de SharePoint.
    Prioridad: site_id_o_url (si es ID), SHAREPOINT_DEFAULT_SITE_ID (env var), lookup raíz.
    Requiere headers autenticados.
    """
    # 1. Si se pasa un ID directamente
    if site_id_o_url and ',' in site_id_o_url: # Formato común de Site ID: hostname,spsite-guid,spweb-guid
        logger.debug(f"Usando Site ID provisto directamente: {site_id_o_url}")
        return site_id_o_url
    # 2. Si se configuró un default en las variables de entorno
    if SHAREPOINT_DEFAULT_SITE_ID:
        logger.debug(f"Usando Site ID default de configuración: {SHAREPOINT_DEFAULT_SITE_ID}")
        return SHAREPOINT_DEFAULT_SITE_ID
    # 3. Si no hay default, intentar obtener el sitio raíz
    url = f"{BASE_URL}/sites/root?$select=id" # Seleccionar solo id
    logger.debug(f"No se proveyó Site ID ni default, obteniendo sitio raíz: GET {url}")
    response: Optional[requests.Response] = None
    try:
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        site_data = response.json()
        site_id = site_data.get('id')
        if not site_id:
            raise ValueError("Respuesta de sitio raíz inválida, falta 'id'.")
        logger.info(f"Site ID raíz obtenido: {site_id}")
        return site_id
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request obteniendo Site ID raíz: {req_ex}", exc_info=True)
        raise Exception(f"Error API obteniendo Site ID raíz: {req_ex}")
    except Exception as e:
        logger.error(f"Error inesperado obteniendo Site ID raíz: {e}", exc_info=True)
        raise

# --- Helpers para Endpoints (Refactorizados) ---
def _get_sp_drive_endpoint(headers: Dict[str, str], site_id: str, drive_id_or_name: Optional[str] = None) -> str:
    """Construye la URL base para un Drive de SharePoint."""
    target_drive = drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID
    # Podríamos verificar si target_drive es un nombre o ID y construir la URL adecuadamente
    # Por simplicidad, asumimos que es usable directamente
    return f"{BASE_URL}/sites/{site_id}/drives/{target_drive}"

def _get_sp_item_path_endpoint(headers: Dict[str, str], site_id: str, item_path: str, drive_id_or_name: Optional[str] = None) -> str:
    """Construye la URL a un item por ruta dentro de un Drive de SP."""
    drive_endpoint = _get_sp_drive_endpoint(headers, site_id, drive_id_or_name)
    safe_path = item_path.strip()
    if not safe_path.startswith('/'): safe_path = '/' + safe_path
    # Si la ruta es solo '/', apunta a la raíz; sino, usa formato /root:/path
    return f"{drive_endpoint}/root" if safe_path == '/' else f"{drive_endpoint}/root:{safe_path}"

def _get_drive_id(headers: Dict[str, str], site_id: str, drive_id_or_name: Optional[str] = None) -> str:
    """Obtiene el ID real de un Drive por su nombre o ID."""
    drive_endpoint = _get_sp_drive_endpoint(headers, site_id, drive_id_or_name)
    url = f"{drive_endpoint}?$select=id"
    response: Optional[requests.Response] = None
    try:
        logger.debug(f"Obteniendo ID real del drive: GET {url}")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        drive_data = response.json()
        actual_drive_id = drive_data.get('id')
        if not actual_drive_id:
            raise ValueError(f"No se pudo obtener 'id' del drive en {drive_endpoint}")
        logger.debug(f"Drive ID obtenido: {actual_drive_id}")
        return actual_drive_id
    except requests.exceptions.RequestException as req_ex:
        logger.error(f"Error Request obteniendo Drive ID para '{drive_id_or_name or SHAREPOINT_DEFAULT_DRIVE_ID}': {req_ex}", exc_info=True)
        raise Exception(f"Error API obteniendo Drive ID: {req_ex}")
    except Exception as e:
        logger.error(f"Error inesperado obteniendo Drive ID: {e}", exc_info=True)
        raise

# ---- FUNCIONES DE LISTAS (Refactorizadas) ----
# Aceptan 'headers' y 'site_id' opcional

def crear_lista(headers: Dict[str, str], nombre_lista: str, site_id: Optional[str] = None) -> dict:
    """Crea una nueva lista de SharePoint. Requiere headers autenticados."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists"
    # Usar columnas del archivo original sharepoint.py
    body = {
        "displayName": nombre_lista,
        "columns": [
            {"name": "Clave", "text": {}},
            {"name": "Valor", "text": {}}
            # Title se crea por defecto
        ],
        "list": {"template": "genericList"}
    }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando lista SP '{nombre_lista}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data=response.json()
        logger.info(f"Lista SP '{nombre_lista}' creada.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en crear_lista (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_lista (SP): {e}", exc_info=True)
        raise

def listar_listas(headers: Dict[str, str], site_id: Optional[str] = None) -> dict:
    """Lista todas las listas en un sitio de SharePoint. Requiere headers autenticados."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists"
    params = {'$select': 'id,name,displayName,webUrl'} # Campos útiles
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Listando listas SP sitio '{target_site_id}')")
        response = requests.get(url, headers=headers, params=params, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data=response.json()
        logger.info(f"Listadas {len(data.get('value',[]))} listas SP.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en listar_listas (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_listas (SP): {e}", exc_info=True)
        raise

def agregar_elemento(headers: Dict[str, str], nombre_lista: str, clave: str, valor: str, site_id: Optional[str] = None) -> dict:
    """Agrega un elemento a una lista de SharePoint (usando Clave/Valor). Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    # El endpoint puede usar nombre o ID de lista
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{nombre_lista}/items"
    body = {"fields": {"Clave": clave, "Valor": valor}} # Coincide con columnas del archivo original
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Agregando elemento SP a lista '{nombre_lista}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data=response.json()
        logger.info(f"Elemento SP agregado a lista '{nombre_lista}'. ID: {data.get('id')}")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en agregar_elemento (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en agregar_elemento (SP): {e}", exc_info=True)
        raise

def listar_elementos(headers: Dict[str, str], nombre_lista: str, site_id: Optional[str] = None, expand_fields: bool = True, top: int = 100) -> dict:
    """Lista elementos de una lista de SharePoint, manejando paginación. Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url_base = f"{BASE_URL}/sites/{target_site_id}/lists/{nombre_lista}/items"
    params = {'$top': min(int(top), 999)}
    if expand_fields:
        params['$expand'] = 'fields'

    all_items = []
    current_url: Optional[str] = url_base
    current_headers = headers.copy() # Copia por si acaso
    response: Optional[requests.Response] = None

    try:
        page_count = 0
        while current_url:
            page_count += 1
            logger.info(f"API Call: GET {current_url} Page: {page_count} (Listando elementos SP lista '{nombre_lista}')")
            # Params solo en la primera llamada
            current_params = params if page_count == 1 else None
            response = requests.get(current_url, headers=current_headers, params=current_params, timeout=GRAPH_API_TIMEOUT)
            response.raise_for_status()
            data = response.json()
            page_items = data.get('value', [])
            all_items.extend(page_items)
            logger.info(f"Página {page_count}: {len(page_items)} elementos SP obtenidos.")
            current_url = data.get('@odata.nextLink')
            # No necesita refrescar headers aquí
        logger.info(f"Total elementos SP lista '{nombre_lista}': {len(all_items)}")
        return {'value': all_items}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en listar_elementos (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_elementos (SP): {e}", exc_info=True)
        raise

def actualizar_elemento(headers: Dict[str, str], nombre_lista: str, item_id: str, nuevos_valores: dict, site_id: Optional[str] = None) -> dict:
    """Actualiza un elemento de lista. 'nuevos_valores' debe ser dict de campos. Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{nombre_lista}/items/{item_id}/fields"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Actualizando elemento SP '{item_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        # Añadir ETag si se proporciona en nuevos_valores
        etag = nuevos_valores.pop('@odata.etag', None)
        if etag: current_headers['If-Match'] = etag

        response = requests.patch(url, headers=current_headers, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Elemento SP '{item_id}' actualizado.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en actualizar_elemento (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en actualizar_elemento (SP): {e}", exc_info=True)
        raise

def eliminar_elemento(headers: Dict[str, str], nombre_lista: str, item_id: str, site_id: Optional[str] = None) -> dict:
    """Elimina un elemento de lista. Requiere headers autenticados."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    url = f"{BASE_URL}/sites/{target_site_id}/lists/{nombre_lista}/items/{item_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando elemento SP '{item_id}')")
        # Añadir ETag si se tiene/necesita
        # etag = obtener_etag_item(headers, site_id, nombre_lista, item_id)
        # current_headers = headers.copy()
        # if etag: current_headers['If-Match'] = etag
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT) # Usar headers originales
        response.raise_for_status() # 204 No Content
        logger.info(f"Elemento SP '{item_id}' eliminado.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en eliminar_elemento (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_elemento (SP): {e}", exc_info=True)
        raise


# ---- FUNCIONES DE DOCUMENTOS (Bibliotecas / Drives) ----

def listar_documentos_biblioteca(headers: Dict[str, str], biblioteca: Optional[str] = None, site_id: Optional[str] = None, top: int = 100, ruta_carpeta: str = '/') -> dict:
    """Lista documentos/carpetas en una ruta de una biblioteca. Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, ruta_carpeta, target_drive)
    url_base = f"{item_endpoint}/children"
    params = {'$top': min(int(top), 999)}

    all_files = []
    current_url: Optional[str] = url_base
    current_headers = headers.copy()
    response: Optional[requests.Response] = None

    try:
        page_count = 0
        while current_url:
            page_count += 1
            logger.info(f"API Call: GET {current_url} Page: {page_count} (Listando docs SP biblioteca '{target_drive}', ruta '{ruta_carpeta}')")
            current_params = params if page_count == 1 else None
            response = requests.get(current_url, headers=current_headers, params=current_params, timeout=GRAPH_API_TIMEOUT)
            response.raise_for_status()
            data = response.json()
            page_items = data.get('value', [])
            all_files.extend(page_items)
            current_url = data.get('@odata.nextLink')
        logger.info(f"Total docs SP biblioteca '{target_drive}', ruta '{ruta_carpeta}': {len(all_files)}")
        return {'value': all_files}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en listar_documentos_biblioteca (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_documentos_biblioteca (SP): {e}", exc_info=True)
        raise

def subir_documento(headers: Dict[str, str], nombre_archivo: str, contenido_bytes: bytes, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta_destino: str = '/', conflict_behavior: str = "rename") -> dict:
    """Sube un documento a una biblioteca. Espera contenido_bytes. Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    target_file_path = os.path.join(ruta_carpeta_destino, nombre_archivo).replace('\\', '/')
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, target_file_path, target_drive)
    url = f"{item_endpoint}:/content?@microsoft.graph.conflictBehavior={conflict_behavior}"

    upload_headers = headers.copy()
    upload_headers['Content-Type'] = 'application/octet-stream' # O el tipo MIME correcto si se conoce
    response: Optional[requests.Response] = None

    try:
        logger.info(f"API Call: PUT {item_endpoint}:/content (Subiendo doc SP '{nombre_archivo}' a '{ruta_carpeta_destino}')")
        if len(contenido_bytes) > 4*1024*1024: logger.warning(f"Archivo SP '{nombre_archivo}' > 4MB.")
        response = requests.put(url, headers=upload_headers, data=contenido_bytes, timeout=GRAPH_API_TIMEOUT * 3)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Doc SP '{nombre_archivo}' subido. ID: {data.get('id')}")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en subir_documento (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en subir_documento (SP): {e}", exc_info=True)
        raise

def eliminar_archivo(headers: Dict[str, str], nombre_archivo: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> dict:
    """Elimina un archivo o carpeta de una biblioteca. Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    item_path = os.path.join(ruta_carpeta, nombre_archivo).replace('\\', '/')
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, item_path, target_drive)
    url = item_endpoint # DELETE va directo al item
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando archivo/carpeta SP '{item_path}')")
        # Añadir If-Match con ETag si se necesita/tiene
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 204
        logger.info(f"Archivo/Carpeta SP '{item_path}' eliminado.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en eliminar_archivo (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_archivo (SP): {e}", exc_info=True)
        raise

# ---- FUNCIONES AVANZADAS DE ARCHIVOS (Refactorizadas) ----

def crear_carpeta_biblioteca(headers: Dict[str, str], nombre_carpeta: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta_padre: str = '/', conflict_behavior: str = "rename") -> dict:
    """Crea una carpeta en una biblioteca. Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    parent_folder_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, ruta_carpeta_padre, target_drive)
    url = f"{parent_folder_endpoint}/children"
    body = {
        "name": nombre_carpeta,
        "folder": {},
        "@microsoft.graph.conflictBehavior": conflict_behavior
    }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando carpeta SP '{nombre_carpeta}' en '{ruta_carpeta_padre}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Carpeta SP '{nombre_carpeta}' creada. ID: {data.get('id')}")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en crear_carpeta_biblioteca (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_carpeta_biblioteca (SP): {e}", exc_info=True)
        raise

def mover_archivo(headers: Dict[str, str], nombre_archivo: str, nueva_ubicacion: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta_origen: str = '/') -> dict:
    """Mueve un archivo o carpeta en una biblioteca. Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    item_path_origen = os.path.join(ruta_carpeta_origen, nombre_archivo).replace('\\', '/')
    item_endpoint_origen = _get_sp_item_path_endpoint(headers, target_site_id, item_path_origen, target_drive)
    url = item_endpoint_origen # PATCH se aplica al origen

    # Obtener ID del drive para parentReference.path
    try:
        actual_drive_id = _get_drive_id(headers, target_site_id, target_drive)
    except Exception as drive_err:
        raise Exception(f"Error obteniendo Drive ID para mover: {drive_err}")

    parent_path = f"/drives/{actual_drive_id}/root:{nueva_ubicacion.strip()}" if nueva_ubicacion != '/' else f"/drives/{actual_drive_id}/root"
    body = {
        "parentReference": {"path": parent_path},
        "name": nombre_archivo # Opcional: permitir renombrar al mover
    }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Moviendo SP '{item_path_origen}' a '{nueva_ubicacion}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.patch(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Archivo/Carpeta SP '{nombre_archivo}' movido a '{nueva_ubicacion}'.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en mover_archivo (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en mover_archivo (SP): {e}", exc_info=True)
        raise

def copiar_archivo(headers: Dict[str, str], nombre_archivo: str, nueva_ubicacion: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta_origen: str = '/', nuevo_nombre_copia: Optional[str] = None) -> dict:
    """Inicia la copia de un archivo o carpeta. Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    item_path_origen = os.path.join(ruta_carpeta_origen, nombre_archivo).replace('\\', '/')
    item_endpoint_origen = _get_sp_item_path_endpoint(headers, target_site_id, item_path_origen, target_drive)
    url = f"{item_endpoint_origen}/copy" # POST a la acción copy

    # Obtener ID del drive para parentReference.driveId
    try:
        actual_drive_id = _get_drive_id(headers, target_site_id, target_drive)
    except Exception as drive_err:
        raise Exception(f"Error obteniendo Drive ID para copiar: {drive_err}")

    parent_path = f"/drive/root:{nueva_ubicacion.strip()}" if nueva_ubicacion != '/' else "/drive/root"
    body = {
        "parentReference": {
            "driveId": actual_drive_id,
            "path": parent_path
            # Podría añadirse siteId si la copia es inter-sitio
        },
        "name": nuevo_nombre_copia or nombre_archivo # Permitir renombrar la copia
    }
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Iniciando copia SP '{item_path_origen}' a '{nueva_ubicacion}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # 202 Accepted
        monitor_url = response.headers.get('Location')
        logger.info(f"Copia SP '{nombre_archivo}' iniciada. Monitor: {monitor_url}")
        return {"status": "Copia Iniciada", "code": response.status_code, "monitorUrl": monitor_url}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en copiar_archivo (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en copiar_archivo (SP): {e}", exc_info=True)
        raise

def obtener_metadatos_archivo(headers: Dict[str, str], nombre_archivo: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> dict:
    """Obtiene metadatos de un archivo/carpeta. Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    item_path = os.path.join(ruta_carpeta, nombre_archivo).replace('\\', '/')
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, item_path, target_drive)
    url = item_endpoint
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo metadatos SP '{item_path}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Metadatos SP '{item_path}' obtenidos.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en obtener_metadatos_archivo (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_metadatos_archivo (SP): {e}", exc_info=True)
        raise

def actualizar_metadatos_archivo(headers: Dict[str, str], nombre_archivo: str, nuevos_valores: dict, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> dict:
    """Actualiza metadatos de un archivo/carpeta (ej: nombre). Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    item_path = os.path.join(ruta_carpeta, nombre_archivo).replace('\\', '/')
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, item_path, target_drive)
    url = item_endpoint
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Actualizando metadatos SP '{item_path}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        # ETag?
        etag = nuevos_valores.pop('@odata.etag', None)
        if etag: current_headers['If-Match'] = etag
        response = requests.patch(url, headers=current_headers, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Metadatos SP '{item_path}' actualizados.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en actualizar_metadatos_archivo (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en actualizar_metadatos_archivo (SP): {e}", exc_info=True)
        raise

def obtener_contenido_archivo(headers: Dict[str, str], nombre_archivo: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> bytes:
    """Obtiene contenido binario de un archivo. Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    item_path = os.path.join(ruta_carpeta, nombre_archivo).replace('\\', '/')
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, item_path, target_drive)
    url = f"{item_endpoint}/content"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo contenido SP '{item_path}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT * 2)
        response.raise_for_status()
        logger.info(f"Contenido SP '{item_path}' obtenido.")
        return response.content
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en obtener_contenido_archivo (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_contenido_archivo (SP): {e}", exc_info=True)
        raise

def actualizar_contenido_archivo(headers: Dict[str, str], nombre_archivo: str, nuevo_contenido: bytes, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/') -> dict:
    """Actualiza contenido binario de un archivo. Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    item_path = os.path.join(ruta_carpeta, nombre_archivo).replace('\\', '/')
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, item_path, target_drive)
    url = f"{item_endpoint}/content"

    upload_headers = headers.copy()
    upload_headers['Content-Type'] = 'application/octet-stream' # o el tipo correcto
    response: Optional[requests.Response] = None

    try:
        logger.info(f"API Call: PUT {url} (Actualizando contenido SP '{item_path}')")
        if len(nuevo_contenido) > 4*1024*1024: logger.warning(f"Archivo SP '{nombre_archivo}' > 4MB.")
        response = requests.put(url, headers=upload_headers, data=nuevo_contenido, timeout=GRAPH_API_TIMEOUT * 3)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Contenido SP '{item_path}' actualizado.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en actualizar_contenido_archivo (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en actualizar_contenido_archivo (SP): {e}", exc_info=True)
        raise

def crear_enlace_compartido_archivo(headers: Dict[str, str], nombre_archivo: str, biblioteca: Optional[str] = None, site_id: Optional[str] = None, ruta_carpeta: str = '/', tipo_enlace: str = "view", alcance: str = "anonymous") -> dict:
    """Crea un enlace compartido. Requiere headers."""
    target_site_id = _obtener_site_id_sp(headers, site_id)
    target_drive = biblioteca or SHAREPOINT_DEFAULT_DRIVE_ID
    item_path = os.path.join(ruta_carpeta, nombre_archivo).replace('\\', '/')
    item_endpoint = _get_sp_item_path_endpoint(headers, target_site_id, item_path, target_drive)
    url = f"{item_endpoint}/createLink"
    body = {"type": tipo_enlace, "scope": alcance}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando enlace SP '{item_path}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Enlace SP creado para '{item_path}'.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en crear_enlace_compartido_archivo (SP): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_enlace_compartido_archivo (SP): {e}", exc_info=True)
        raise
