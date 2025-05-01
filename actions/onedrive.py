# actions/onedrive.py (Refactorizado)

import logging
import requests
import os
from typing import Dict, Optional, Union, List

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

# ---- Helpers Locales (Usan BASE_URL global) ----
def _get_od_me_drive_endpoint() -> str:
    return f"{BASE_URL}/me/drive"

def _get_od_me_item_path_endpoint(ruta_relativa: str) -> str:
    drive_endpoint = _get_od_me_drive_endpoint()
    safe_path = ruta_relativa.strip()
    if not safe_path.startswith('/'): safe_path = '/' + safe_path
    return f"{drive_endpoint}/root" if safe_path == '/' else f"{drive_endpoint}/root:{safe_path}"

# ---- FUNCIONES ONEDRIVE (Refactorizadas) ----
# Aceptan 'headers' como parámetro obligatorio

def listar_archivos(headers: Dict[str, str], ruta: str = "/", top: int = 100) -> dict:
    """Lista archivos y carpetas en OneDrive (/me). Requiere headers autenticados."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    item_endpoint = _get_od_me_item_path_endpoint(ruta)
    url = f"{item_endpoint}/children"
    params = {'$top': min(int(top), 999)}
    all_items = []
    current_url: Optional[str] = url
    current_headers = headers.copy()

    try:
        page_count = 0
        while current_url:
            page_count += 1
            logger.info(f"API Call: GET {current_url} Page: {page_count} (Listando OD /me ruta '{ruta}')")
            current_params = params if page_count == 1 else None
            response = requests.get(current_url, headers=current_headers, params=current_params, timeout=GRAPH_API_TIMEOUT)
            response.raise_for_status()
            data = response.json()
            page_items = data.get('value', [])
            all_items.extend(page_items)
            current_url = data.get('@odata.nextLink')
        logger.info(f"Total items OD /me en '{ruta}': {len(all_items)}")
        return {'value': all_items}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en listar_archivos (OD): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_archivos (OD): {e}", exc_info=True)
        raise

def subir_archivo(headers: Dict[str, str], nombre_archivo: str, contenido_bytes: bytes, ruta: str = "/", conflict_behavior: str = "rename") -> dict:
    """Sube un archivo a OneDrive (/me). Requiere headers autenticados."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    # Usa 'nombre_archivo' directamente, asumiendo que es el nombre destino
    target_file_path = os.path.join(ruta, nombre_archivo).replace('\\', '/')
    item_endpoint = _get_od_me_item_path_endpoint(target_file_path)
    url = f"{item_endpoint}:/content?@microsoft.graph.conflictBehavior={conflict_behavior}"

    upload_headers = headers.copy()
    upload_headers['Content-Type'] = 'application/octet-stream'

    try:
        logger.info(f"API Call: PUT {item_endpoint}:/content (Subiendo OD /me '{nombre_archivo}' a ruta '{ruta}')")
        if len(contenido_bytes) > 4*1024*1024: logger.warning(f"Archivo OD '{nombre_archivo}' > 4MB.")
        response = requests.put(url, headers=upload_headers, data=contenido_bytes, timeout=GRAPH_API_TIMEOUT * 3)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Archivo OD '{nombre_archivo}' subido.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en subir_archivo (OD): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en subir_archivo (OD): {e}", exc_info=True)
        raise

def descargar_archivo(headers: Dict[str, str], nombre_archivo: str, ruta: str = "/") -> bytes:
    """Descarga un archivo de OneDrive (/me). Requiere headers autenticados."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    item_endpoint = _get_od_me_item_path_endpoint(os.path.join(ruta, nombre_archivo).replace('\\', '/'))
    url = f"{item_endpoint}/content"
    try:
        logger.info(f"API Call: GET {url} (Descargando OD /me '{nombre_archivo}' de ruta '{ruta}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT * 2)
        response.raise_for_status()
        logger.info(f"Archivo OD '{nombre_archivo}' descargado.")
        return response.content
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en descargar_archivo (OD): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en descargar_archivo (OD): {e}", exc_info=True)
        raise

def eliminar_archivo(headers: Dict[str, str], nombre_archivo: str, ruta: str = "/") -> dict:
    """Elimina un archivo o carpeta de OneDrive (/me). Requiere headers autenticados."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    item_endpoint = _get_od_me_item_path_endpoint(os.path.join(ruta, nombre_archivo).replace('\\', '/'))
    url = item_endpoint # DELETE va al item
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando OD /me '{nombre_archivo}' de ruta '{ruta}')")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        logger.info(f"Item OD '{nombre_archivo}' eliminado.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en eliminar_archivo (OD): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_archivo (OD): {e}", exc_info=True)
        raise

def crear_carpeta(headers: Dict[str, str], nombre_carpeta: str, ruta: str = "/") -> dict:
    """Crea una carpeta en OneDrive (/me). Requiere headers autenticados."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    parent_folder_endpoint = _get_od_me_item_path_endpoint(ruta)
    url = f"{parent_folder_endpoint}/children"
    body = {"name": nombre_carpeta, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"}
    post_headers = headers.copy()
    post_headers['Content-Type'] = 'application/json'
    try:
        logger.info(f"API Call: POST {url} (Creando OD /me carpeta '{nombre_carpeta}' en ruta '{ruta}')")
        response = requests.post(url, headers=post_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Carpeta OD '{nombre_carpeta}' creada.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en crear_carpeta (OD): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_carpeta (OD): {e}", exc_info=True)
        raise

def mover_archivo(headers: Dict[str, str], nombre_archivo: str, ruta_origen: str = "/", ruta_destino: str = "/NuevaCarpeta") -> dict:
    """Mueve un archivo o carpeta en OneDrive (/me). Requiere headers autenticados."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    item_origen_endpoint = _get_od_me_item_path_endpoint(os.path.join(ruta_origen, nombre_archivo).replace('\\', '/'))
    url = item_origen_endpoint # PATCH en el origen

    parent_path = f"/drive/root:{ruta_destino.strip()}" if ruta_destino != '/' else "/drive/root"
    body = {"parentReference": {"path": parent_path}, "name": nombre_archivo} # Mantener nombre original
    patch_headers = headers.copy()
    patch_headers['Content-Type'] = 'application/json'
    try:
        logger.info(f"API Call: PATCH {url} (Moviendo OD /me '{nombre_archivo}' de '{ruta_origen}' a '{ruta_destino}')")
        response = requests.patch(url, headers=patch_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Item OD '{nombre_archivo}' movido.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en mover_archivo (OD): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en mover_archivo (OD): {e}", exc_info=True)
        raise

def copiar_archivo(headers: Dict[str, str], nombre_archivo: str, ruta_origen: str = "/", ruta_destino: str = "/Copias") -> dict:
    """Inicia la copia de un archivo o carpeta en OneDrive (/me). Requiere headers autenticados."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    # Necesita ID del drive /me para parentReference
    drive_id = None
    try:
        drive_resp = requests.get(f"{BASE_URL}/me/drive?$select=id", headers=headers, timeout=GRAPH_API_TIMEOUT)
        drive_resp.raise_for_status()
        drive_id = drive_resp.json().get('id')
        if not drive_id: raise ValueError("No se pudo obtener ID del drive /me")
    except Exception as drive_err:
        logger.error(f"Error obteniendo ID drive /me para copia: {drive_err}", exc_info=True)
        raise Exception(f"Error obteniendo ID drive /me: {drive_err}")

    item_origen_endpoint = _get_od_me_item_path_endpoint(os.path.join(ruta_origen, nombre_archivo).replace('\\', '/'))
    url = f"{item_origen_endpoint}/copy"
    parent_path = f"/drive/root:{ruta_destino.strip()}" if ruta_destino != '/' else "/drive/root"
    body = {"parentReference": {"driveId": drive_id, "path": parent_path}, "name": f"Copia_{nombre_archivo}"}
    post_headers = headers.copy()
    post_headers['Content-Type'] = 'application/json'
    try:
        logger.info(f"API Call: POST {url} (Copiando OD /me '{nombre_archivo}' de '{ruta_origen}' a '{ruta_destino}')")
        response = requests.post(url, headers=post_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 202 Accepted
        monitor_url = response.headers.get('Location')
        logger.info(f"Copia OD '{nombre_archivo}' iniciada. Monitor: {monitor_url}")
        return {"status": "Copia Iniciada", "code": response.status_code, "monitorUrl": monitor_url}
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en copiar_archivo (OD): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en copiar_archivo (OD): {e}", exc_info=True)
        raise

def obtener_metadatos_archivo(headers: Dict[str, str], nombre_archivo: str, ruta: str = "/") -> dict:
    """Obtiene metadatos de un archivo/carpeta en OneDrive (/me). Requiere headers autenticados."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    item_endpoint = _get_od_me_item_path_endpoint(os.path.join(ruta, nombre_archivo).replace('\\', '/'))
    url = item_endpoint
    try:
        logger.info(f"API Call: GET {url} (Obteniendo metadatos OD /me '{nombre_archivo}' ruta '{ruta}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Metadatos OD '{nombre_archivo}' obtenidos.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en obtener_metadatos_archivo (OD): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_metadatos_archivo (OD): {e}", exc_info=True)
        raise

def actualizar_metadatos_archivo(headers: Dict[str, str], nombre_archivo: str, nuevos_valores: dict, ruta:str = "/") -> dict:
    """Actualiza metadatos de un archivo/carpeta en OneDrive (/me). Requiere headers autenticados."""
    if headers is None: raise ValueError("Headers autenticados requeridos.")
    item_endpoint = _get_od_me_item_path_endpoint(os.path.join(ruta, nombre_archivo).replace('\\', '/'))
    url = item_endpoint
    patch_headers = headers.copy()
    patch_headers['Content-Type'] = 'application/json'
    try:
        logger.info(f"API Call: PATCH {url} (Actualizando metadatos OD /me '{nombre_archivo}' ruta '{ruta}')")
        response = requests.patch(url, headers=patch_headers, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Metadatos OD '{nombre_archivo}' actualizados.")
        return data
    except requests.exceptions.RequestException as req_ex:
         logger.error(f"Error Request en actualizar_metadatos_archivo (OD): {req_ex}", exc_info=True)
         raise
    except Exception as e:
        logger.error(f"Error inesperado en actualizar_metadatos_archivo (OD): {e}", exc_info=True)
        raise
