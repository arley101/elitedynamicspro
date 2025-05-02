# actions/onedrive.py (Refactorizado)

import logging
import requests # Para tipos de excepción y llamadas directas donde el helper no aplica directamente
import os
import json
from typing import Dict, Optional, Union, List, Any

# Usar el logger estándar de Azure Functions
logger = logging.getLogger("azure.functions")

# Importar helper y constantes desde la estructura compartida
try:
    # Asume que shared está un nivel arriba de actions
    from ..shared.helpers.http_client import hacer_llamada_api
    from ..shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError as e:
    logging.critical(f"Error CRÍTICO importando helpers/constantes en OneDrive: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs):
        raise NotImplementedError("Dependencia 'hacer_llamada_api' no importada correctamente.")

# ---- Helpers Locales para Endpoints de OneDrive (/me/drive) ----
# Estos solo construyen URLs
def _get_od_me_drive_endpoint() -> str:
    """Devuelve el endpoint base para el drive principal del usuario."""
    return f"{BASE_URL}/me/drive"

def _get_od_me_item_path_endpoint(ruta_relativa: str) -> str:
    """Construye la URL para un item específico por path relativo a la raíz de /me/drive."""
    drive_endpoint = _get_od_me_drive_endpoint()
    # Limpiar y asegurar que el path empiece con '/' si no es vacío
    safe_path = ruta_relativa.strip()
    if not safe_path: # Si el path es vacío, apuntar a la raíz
        safe_path = '/'
    if not safe_path.startswith('/'):
        safe_path = '/' + safe_path

    # Si el path es solo '/', apunta a la raíz del drive
    if safe_path == '/':
        return f"{drive_endpoint}/root"
    else:
        # Para otros paths, se usa el formato /root:/path/to/item
        # Asegurarse de que no haya doble '//' si safe_path ya empezaba con /
        return f"{drive_endpoint}/root:{safe_path}"

# ---- FUNCIONES DE ACCIÓN PARA ONEDRIVE (/me/drive) ----
# Todas usan la firma (parametros: Dict[str, Any], headers: Dict[str, str])

def listar_archivos(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista archivos y carpetas en una ruta específica de OneDrive (/me/drive).

    Args:
        parametros (Dict[str, Any]): Opcional: 'ruta' (default '/'), 'top' (int, default 100).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Un diccionario {'value': [lista_completa_de_items]}.
    """
    ruta: str = parametros.get("ruta", "/")
    top: int = int(parametros.get("top", 100))

    item_endpoint = _get_od_me_item_path_endpoint(ruta)
    url_base = f"{item_endpoint}/children" # Endpoint para listar hijos
    params_query: Dict[str, Any] = {'$top': min(top, 999)} # Limitar top por llamada

    all_items: List[Dict[str, Any]] = []
    current_url: Optional[str] = url_base
    page_count = 0
    max_pages = 100 # Límite de seguridad

    try:
        while current_url and page_count < max_pages:
            page_count += 1
            logger.info(f"Listando OneDrive /me ruta '{ruta}', Página: {page_count}")

            current_params_page = params_query if page_count == 1 else None
            # Usar helper centralizado para cada página
            data = hacer_llamada_api("GET", current_url, headers, params=current_params_page)

            if data:
                page_items = data.get('value', [])
                all_items.extend(page_items)
                current_url = data.get('@odata.nextLink')
                if not current_url: break
            else:
                 logger.warning(f"Llamada a {current_url} para listar OneDrive devolvió None/vacío.")
                 break

        if page_count >= max_pages:
             logger.warning(f"Se alcanzó límite de {max_pages} páginas listando OneDrive en '{ruta}'.")

        logger.info(f"Total items OneDrive /me en '{ruta}': {len(all_items)}")
        return {'value': all_items}

    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_archivos (OneDrive) página {page_count}: {e}", exc_info=True)
        raise Exception(f"Error API listando archivos OneDrive: {e}") from e
    except Exception as e:
        logger.error(f"Error inesperado en listar_archivos (OneDrive) página {page_count}: {e}", exc_info=True)
        raise


def subir_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Sube un archivo a OneDrive (/me/drive). Maneja sesión de carga para >4MB.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo', 'contenido_bytes'.
                                     Opcional: 'ruta' (default '/'),
                                     'conflict_behavior' ('rename', 'replace', 'fail', default 'rename').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos del archivo subido.
    """
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo")
    contenido_bytes: Optional[bytes] = parametros.get("contenido_bytes")
    ruta: str = parametros.get("ruta", "/")
    conflict_behavior: str = parametros.get("conflict_behavior", "rename")

    if not nombre_archivo: raise ValueError("Parámetro 'nombre_archivo' es requerido.")
    if contenido_bytes is None or not isinstance(contenido_bytes, bytes):
        raise ValueError("Parámetro 'contenido_bytes' (bytes) es requerido.")

    # Construir path relativo al root de OneDrive
    target_folder_path = ruta.strip('/')
    target_file_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"

    # Endpoint para subir contenido por path
    item_endpoint = _get_od_me_item_path_endpoint(target_file_path)
    url_put_simple = f"{item_endpoint}:/content"
    params_query = {"@microsoft.graph.conflictBehavior": conflict_behavior}

    # Headers específicos para subida
    upload_headers = headers.copy()
    upload_headers['Content-Type'] = 'application/octet-stream' # Genérico para bytes

    file_size_mb = len(contenido_bytes) / (1024 * 1024)
    logger.info(f"Subiendo a OneDrive /me '{nombre_archivo}' ({file_size_mb:.2f} MB) a ruta '{ruta}' con conflict='{conflict_behavior}'")

    # --- Lógica de Subida ---
    if file_size_mb > 4.0:
        # --- INICIO: Lógica de Sesión de Carga ---
        create_session_url = f"{item_endpoint}:/createUploadSession"
        session_body = {"item": {"@microsoft.graph.conflictBehavior": conflict_behavior}}
        try:
            logger.info(f"Archivo > 4MB. Creando sesión de carga para '{nombre_archivo}'...")
            # Usar helper para crear sesión
            session_info = hacer_llamada_api("POST", create_session_url, headers, json_data=session_body)
            upload_url = session_info.get("uploadUrl")
            if not upload_url: raise ValueError("No se pudo obtener 'uploadUrl' de la sesión de carga.")
            logger.info(f"Sesión de carga creada. URL: {upload_url[:50]}...")

            # Subir fragmentos (usando requests directo por simplicidad aquí)
            chunk_size = 5 * 1024 * 1024 # 5 MB
            start_byte = 0
            total_bytes = len(contenido_bytes)
            last_response_json = {}
            while start_byte < total_bytes:
                end_byte = min(start_byte + chunk_size - 1, total_bytes - 1)
                chunk_data = contenido_bytes[start_byte : end_byte + 1]
                content_range = f"bytes {start_byte}-{end_byte}/{total_bytes}"
                chunk_headers = {'Content-Length': str(len(chunk_data)), 'Content-Range': content_range}
                logger.debug(f"Subiendo chunk OneDrive: {content_range}")
                chunk_timeout = max(GRAPH_API_TIMEOUT, int(file_size_mb * 5))
                # PUT a uploadUrl no necesita Auth header
                chunk_response = requests.put(upload_url, headers=chunk_headers, data=chunk_data, timeout=chunk_timeout)
                chunk_response.raise_for_status()
                start_byte = end_byte + 1
                # Guardar la última respuesta JSON (contiene metadatos al final)
                if chunk_response.content: # Solo si hay cuerpo en la respuesta
                    try: last_response_json = chunk_response.json()
                    except json.JSONDecodeError: pass # Ignorar si no es JSON

            logger.info(f"Archivo OneDrive '{nombre_archivo}' subido exitosamente mediante sesión.")
            return last_response_json # Devolver metadatos de la última respuesta

        except requests.exceptions.RequestException as e:
            logger.error(f"Error Request durante sesión de carga OneDrive para '{nombre_archivo}': {e}", exc_info=True)
            raise Exception(f"Error API durante sesión de carga OneDrive: {e}") from e
        except Exception as e:
            logger.error(f"Error inesperado durante sesión de carga OneDrive para '{nombre_archivo}': {e}", exc_info=True)
            raise
        # --- FIN: Lógica de Sesión de Carga ---
    else:
        # --- Subida Simple (<= 4MB) ---
        try:
             simple_upload_timeout = max(GRAPH_API_TIMEOUT, int(file_size_mb * 10))
             # Usar helper con 'data'
             resultado = hacer_llamada_api(
                 metodo="PUT",
                 url=url_put_simple,
                 headers=upload_headers,
                 params=params_query,
                 data=contenido_bytes,
                 timeout=simple_upload_timeout,
                 expect_json=True
             )
             logger.info(f"Archivo OneDrive '{nombre_archivo}' subido (subida simple).")
             return resultado
        except requests.exceptions.RequestException as e:
            logger.error(f"Error Request en subida simple OneDrive de '{nombre_archivo}': {e}", exc_info=True)
            raise Exception(f"Error API subiendo archivo OneDrive (simple): {e}") from e
        except Exception as e:
            logger.error(f"Error inesperado en subida simple OneDrive de '{nombre_archivo}': {e}", exc_info=True)
            raise


def descargar_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> bytes:
    """
    Descarga el contenido de un archivo de OneDrive (/me/drive).

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo'. Opcional: 'ruta' (default '/').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        bytes: El contenido binario del archivo.
    """
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo")
    ruta: str = parametros.get("ruta", "/")

    if not nombre_archivo: raise ValueError("Parámetro 'nombre_archivo' es requerido.")

    # Construir path y endpoint
    target_folder_path = ruta.strip('/')
    target_file_path = f"/{nombre_archivo}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo}"
    item_endpoint = _get_od_me_item_path_endpoint(target_file_path)
    url = f"{item_endpoint}/content" # Endpoint de contenido

    logger.info(f"Descargando archivo OneDrive /me '{nombre_archivo}' de ruta '{ruta}'")

    # Usar helper con expect_json=False para obtener objeto Response
    download_timeout = max(GRAPH_API_TIMEOUT, 60)
    response = hacer_llamada_api("GET", url, headers, timeout=download_timeout, expect_json=False)

    if isinstance(response, requests.Response):
        logger.info(f"Archivo OneDrive '{nombre_archivo}' descargado ({len(response.content)} bytes).")
        return response.content
    else:
        logger.error(f"Respuesta inesperada del helper al descargar archivo OneDrive: {type(response)}")
        raise Exception("Error interno al descargar archivo OneDrive.")


def eliminar_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina un archivo o carpeta de OneDrive (/me/drive).

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo_o_carpeta'. Opcional: 'ruta' (default '/').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de eliminación.
    """
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta")
    ruta: str = parametros.get("ruta", "/")

    if not nombre_archivo_o_carpeta: raise ValueError("Parámetro 'nombre_archivo_o_carpeta' es requerido.")

    # Construir path y endpoint
    target_folder_path = ruta.strip('/')
    target_file_path = f"/{nombre_archivo_o_carpeta}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo_o_carpeta}"
    item_endpoint = _get_od_me_item_path_endpoint(target_file_path)
    url = item_endpoint # DELETE en el endpoint del item

    logger.info(f"Eliminando archivo/carpeta OneDrive /me '{nombre_archivo_o_carpeta}' de ruta '{ruta}'")
    # Helper devuelve None en éxito 204
    hacer_llamada_api("DELETE", url, headers)
    return {"status": "Eliminado", "path": target_file_path}


def crear_carpeta(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea una nueva carpeta en OneDrive (/me/drive).

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_carpeta'.
                                     Opcional: 'ruta' (carpeta padre, default '/'),
                                     'conflict_behavior' ('rename', 'replace', 'fail', default 'rename').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos de la carpeta creada.
    """
    nombre_carpeta: Optional[str] = parametros.get("nombre_carpeta")
    ruta: str = parametros.get("ruta", "/") # Carpeta padre
    conflict_behavior: str = parametros.get("conflict_behavior", "rename")

    if not nombre_carpeta: raise ValueError("Parámetro 'nombre_carpeta' es requerido.")

    # Endpoint de la carpeta padre
    parent_folder_endpoint = _get_od_me_item_path_endpoint(ruta)
    url = f"{parent_folder_endpoint}/children" # POST a /children crea item

    body: Dict[str, Any] = {
        "name": nombre_carpeta,
        "folder": {}, # Indica que es una carpeta
        "@microsoft.graph.conflictBehavior": conflict_behavior
    }
    logger.info(f"Creando carpeta OneDrive /me '{nombre_carpeta}' en ruta '{ruta}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)


def mover_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Mueve un archivo o carpeta a una nueva ubicación en OneDrive (/me/drive).

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo_o_carpeta', 'nueva_ruta_carpeta_padre'.
                                     Opcional: 'ruta_origen' (default '/'),
                                     'nuevo_nombre' (para renombrar al mover).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos del item movido/renombrado.
    """
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta")
    ruta_origen: str = parametros.get("ruta_origen", "/")
    nueva_ruta_carpeta_padre: Optional[str] = parametros.get("nueva_ruta_carpeta_padre")
    nuevo_nombre: Optional[str] = parametros.get("nuevo_nombre")

    if not nombre_archivo_o_carpeta: raise ValueError("Parámetro 'nombre_archivo_o_carpeta' es requerido.")
    if nueva_ruta_carpeta_padre is None: raise ValueError("Parámetro 'nueva_ruta_carpeta_padre' es requerido.")

    # Path de origen
    target_folder_path_origen = ruta_origen.strip('/')
    item_path_origen = f"/{nombre_archivo_o_carpeta}" if not target_folder_path_origen else f"/{target_folder_path_origen}/{nombre_archivo_o_carpeta}"
    item_origen_endpoint = _get_od_me_item_path_endpoint(item_path_origen)
    url = item_origen_endpoint # PATCH sobre el item de origen

    # Construir referencia a la carpeta padre de destino
    parent_dest_path = nueva_ruta_carpeta_padre.strip()
    if not parent_dest_path.startswith('/'): parent_dest_path = '/' + parent_dest_path
    # La referencia al padre usa /drive/root:/path/to/parent
    parent_reference_path = "/drive/root" if parent_dest_path == '/' else f"/drive/root:{parent_dest_path}"

    body: Dict[str, Any] = {
        "parentReference": {
            "path": parent_reference_path
            # Se podría usar 'id' de la carpeta destino si se tiene
        }
    }
    # Añadir nuevo nombre si se especifica
    body["name"] = nuevo_nombre if nuevo_nombre is not None else nombre_archivo_o_carpeta

    logger.info(f"Moviendo OneDrive /me '{item_path_origen}' a '{nueva_ruta_carpeta_padre}' (nuevo nombre: {body['name']})")
    return hacer_llamada_api("PATCH", url, headers, json_data=body)


def copiar_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Copia un archivo a una nueva ubicación en OneDrive (/me/drive). Operación asíncrona.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo', 'nueva_ruta_carpeta_padre'.
                                     Opcional: 'ruta_origen' (default '/'), 'nuevo_nombre_copia'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Respuesta 202 Accepted con URL de monitorización.
    """
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo")
    ruta_origen: str = parametros.get("ruta_origen", "/")
    nueva_ruta_carpeta_padre: Optional[str] = parametros.get("nueva_ruta_carpeta_padre")
    nuevo_nombre_copia: Optional[str] = parametros.get("nuevo_nombre_copia")

    if not nombre_archivo: raise ValueError("Parámetro 'nombre_archivo' es requerido.")
    if nueva_ruta_carpeta_padre is None: raise ValueError("Parámetro 'nueva_ruta_carpeta_padre' es requerido.")

    # Obtener ID del drive /me/drive (necesario para parentReference)
    drive_endpoint = _get_od_me_drive_endpoint()
    try:
        drive_data = hacer_llamada_api("GET", f"{drive_endpoint}?$select=id", headers)
        actual_drive_id = drive_data.get('id')
        if not actual_drive_id: raise ValueError("No se pudo obtener ID del drive /me.")
    except Exception as drive_err:
        logger.error(f"Error obteniendo ID drive /me para copiar: {drive_err}", exc_info=True)
        raise Exception(f"Error obteniendo ID drive /me para copia: {drive_err}") from drive_err

    # Path de origen
    target_folder_path_origen = ruta_origen.strip('/')
    item_path_origen = f"/{nombre_archivo}" if not target_folder_path_origen else f"/{target_folder_path_origen}/{nombre_archivo}"
    item_origen_endpoint = _get_od_me_item_path_endpoint(item_path_origen)
    url = f"{item_origen_endpoint}/copy" # Endpoint de copia

    # Referencia a carpeta padre destino
    parent_dest_path = nueva_ruta_carpeta_padre.strip()
    if not parent_dest_path.startswith('/'): parent_dest_path = '/' + parent_dest_path
    # La referencia usa driveId y path relativo a la raíz de ESE drive
    parent_reference_path = "/drive/root" if parent_dest_path == '/' else f"/drive/root:{parent_dest_path}"

    body: Dict[str, Any] = {
        "parentReference": {
            "driveId": actual_drive_id, # Asume copia dentro del mismo drive
            "path": parent_reference_path
        }
    }
    # Nombre opcional para la copia
    body["name"] = nuevo_nombre_copia if nuevo_nombre_copia is not None else f"Copia de {nombre_archivo}"

    logger.info(f"Iniciando copia asíncrona OneDrive /me de '{item_path_origen}' a '{nueva_ruta_carpeta_padre}'")

    # La copia devuelve 202 Accepted. Usar helper con expect_json=False.
    response = hacer_llamada_api("POST", url, headers, json_data=body, expect_json=False)

    if isinstance(response, requests.Response) and response.status_code == 202:
        monitor_url = response.headers.get('Location')
        logger.info(f"Copia OneDrive '{nombre_archivo}' iniciada. Monitor URL: {monitor_url}")
        return {
            "status": "Copia Iniciada",
            "status_code": response.status_code,
            "monitorUrl": monitor_url,
            "detail": "La copia se realiza en segundo plano. Usa la URL de monitorización."
        }
    elif isinstance(response, requests.Response):
         logger.error(f"Respuesta inesperada al iniciar copia OneDrive: {response.status_code} {response.reason}.")
         raise Exception(f"Respuesta inesperada al iniciar copia OneDrive: {response.status_code}")
    else:
         logger.error(f"Respuesta inesperada del helper al iniciar copia OneDrive: {type(response)}")
         raise Exception("Error interno al procesar la solicitud de copia OneDrive.")


def obtener_metadatos_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene los metadatos de un archivo o carpeta en OneDrive (/me/drive).

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo_o_carpeta'. Opcional: 'ruta' (default '/').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos del item.
    """
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta")
    ruta: str = parametros.get("ruta", "/")

    if not nombre_archivo_o_carpeta: raise ValueError("Parámetro 'nombre_archivo_o_carpeta' es requerido.")

    # Construir path y endpoint
    target_folder_path = ruta.strip('/')
    item_path = f"/{nombre_archivo_o_carpeta}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo_o_carpeta}"
    item_endpoint = _get_od_me_item_path_endpoint(item_path)
    url = item_endpoint # GET en el endpoint del item

    logger.info(f"Obteniendo metadatos OneDrive /me '{item_path}'")
    return hacer_llamada_api("GET", url, headers)


def actualizar_metadatos_archivo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Actualiza metadatos de un archivo o carpeta en OneDrive (/me/drive). Soporta ETag.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo_o_carpeta', 'nuevos_valores' (dict).
                                     Opcional: 'ruta' (default '/'), '@odata.etag' dentro de nuevos_valores.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos actualizados.
    """
    nombre_archivo_o_carpeta: Optional[str] = parametros.get("nombre_archivo_o_carpeta")
    ruta: str = parametros.get("ruta", "/")
    nuevos_valores: Optional[Dict[str, Any]] = parametros.get("nuevos_valores")

    if not nombre_archivo_o_carpeta: raise ValueError("Parámetro 'nombre_archivo_o_carpeta' es requerido.")
    if not nuevos_valores or not isinstance(nuevos_valores, dict):
        raise ValueError("Parámetro 'nuevos_valores' (diccionario) es requerido.")

    # Construir path y endpoint
    target_folder_path = ruta.strip('/')
    item_path = f"/{nombre_archivo_o_carpeta}" if not target_folder_path else f"/{target_folder_path}/{nombre_archivo_o_carpeta}"
    item_endpoint = _get_od_me_item_path_endpoint(item_path)
    url = item_endpoint # PATCH en el endpoint del item

    # Manejar ETag
    current_headers = headers.copy()
    body_data = nuevos_valores.copy()
    etag = body_data.pop('@odata.etag', None)
    if etag:
        current_headers['If-Match'] = etag
        logger.debug("Usando ETag para actualización de metadatos OneDrive.")

    logger.info(f"Actualizando metadatos OneDrive /me '{item_path}'")
    return hacer_llamada_api("PATCH", url, current_headers, json_data=body_data)

# --- FIN DEL MÓDULO actions/onedrive.py ---
