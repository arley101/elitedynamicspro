# actions/planner_todo.py (Refactorizado)

import logging
import requests # Solo para tipos de excepción
import json
from typing import Dict, List, Optional, Union, Any
from datetime import datetime, timezone

# Usar logger estándar de Azure Functions
logger = logging.getLogger("azure.functions")

# Importar helper y constantes desde la estructura compartida
try:
    # Asume que shared está un nivel arriba de actions
    from ..shared.helpers.http_client import hacer_llamada_api
    from ..shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError as e:
    logging.critical(f"Error CRÍTICO importando helpers/constantes en Planner/ToDo: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs):
        raise NotImplementedError("Dependencia 'hacer_llamada_api' no importada correctamente.")

# ==================================
# ==== FUNCIONES ACCIÓN PLANNER ====
# ==================================
# Todas usan la firma (parametros: Dict[str, Any], headers: Dict[str, str])

def listar_planes(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista los planes de Planner asociados a un Grupo de Microsoft 365.

    Args:
        parametros (Dict[str, Any]): Debe contener 'grupo_id'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Respuesta de Graph API, usualmente {'value': [...]}.
    """
    grupo_id: Optional[str] = parametros.get("grupo_id")
    if not grupo_id: raise ValueError("Parámetro 'grupo_id' es requerido.")

    url = f"{BASE_URL}/groups/{grupo_id}/planner/plans"
    logger.info(f"Listando planes de Planner para grupo '{grupo_id}'")
    return hacer_llamada_api("GET", url, headers)


def obtener_plan(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Obtiene los detalles de un Plan de Planner específico.

    Args:
        parametros (Dict[str, Any]): Debe contener 'plan_id'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El objeto del plan de Graph API.
    """
    plan_id: Optional[str] = parametros.get("plan_id")
    if not plan_id: raise ValueError("Parámetro 'plan_id' es requerido.")

    url = f"{BASE_URL}/planner/plans/{plan_id}"
    logger.info(f"Obteniendo detalles del plan de Planner '{plan_id}'")
    return hacer_llamada_api("GET", url, headers)


def crear_plan(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea un nuevo Plan de Planner asociado a un Grupo.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_plan', 'grupo_id' (owner).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El objeto del plan creado.
    """
    nombre_plan: Optional[str] = parametros.get("nombre_plan")
    grupo_id: Optional[str] = parametros.get("grupo_id") # El grupo es el 'owner'

    if not nombre_plan: raise ValueError("Parámetro 'nombre_plan' es requerido.")
    if not grupo_id: raise ValueError("Parámetro 'grupo_id' (owner) es requerido.")

    url = f"{BASE_URL}/planner/plans"
    body = {"owner": grupo_id, "title": nombre_plan}
    logger.info(f"Creando plan de Planner '{nombre_plan}' para grupo '{grupo_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)


def actualizar_plan(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Actualiza los detalles de un Plan de Planner. Soporta ETag.

    Args:
        parametros (Dict[str, Any]): Debe contener 'plan_id', 'nuevos_valores' (dict).
                                     Opcional: '@odata.etag' dentro de nuevos_valores.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El objeto del plan actualizado (o re-obtenido si la API devuelve 204).
    """
    plan_id: Optional[str] = parametros.get("plan_id")
    nuevos_valores: Optional[Dict[str, Any]] = parametros.get("nuevos_valores")

    if not plan_id: raise ValueError("Parámetro 'plan_id' es requerido.")
    if not nuevos_valores or not isinstance(nuevos_valores, dict):
        raise ValueError("Parámetro 'nuevos_valores' (diccionario) es requerido.")

    url = f"{BASE_URL}/planner/plans/{plan_id}"
    current_headers = headers.copy()
    body_data = nuevos_valores.copy()
    etag = body_data.pop('@odata.etag', None) # Extraer ETag del cuerpo si existe

    if etag:
        current_headers['If-Match'] = etag
        logger.info(f"Usando ETag '{etag}' para actualizar plan {plan_id}")
    else:
        logger.warning(f"Actualizando plan {plan_id} sin ETag. Podría haber conflictos.")

    logger.info(f"Actualizando plan de Planner '{plan_id}'")
    # PATCH en Planner puede devolver 204 No Content o 200 OK con el objeto actualizado.
    result = hacer_llamada_api("PATCH", url, current_headers, json_data=body_data)

    if result is None: # Implica 204 No Content
         logger.warning(f"Actualizar plan {plan_id} devolvió 204 No Content. Re-obteniendo plan para devolver estado actual.")
         # Llamar a obtener_plan para devolver el estado actualizado
         params_get = {"plan_id": plan_id}
         return obtener_plan(params_get, headers)
    else:
         # Devolver el resultado si la API devolvió 200 OK con cuerpo
         return result


def eliminar_plan(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina un Plan de Planner. Soporta ETag.

    Args:
        parametros (Dict[str, Any]): Debe contener 'plan_id'. Opcional: 'etag'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de eliminación.
    """
    plan_id: Optional[str] = parametros.get("plan_id")
    etag: Optional[str] = parametros.get("etag") # ETag como parámetro separado

    if not plan_id: raise ValueError("Parámetro 'plan_id' es requerido.")

    url = f"{BASE_URL}/planner/plans/{plan_id}"
    current_headers = headers.copy()
    if etag:
        current_headers['If-Match'] = etag
        logger.info(f"Eliminando plan {plan_id} con ETag '{etag}'.")
    else:
        logger.warning(f"Eliminando plan {plan_id} sin ETag.")

    # DELETE devuelve 204 No Content (None del helper).
    hacer_llamada_api("DELETE", url, current_headers)
    return {"status": "Eliminado", "id": plan_id}


def listar_tareas_planner(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista las tareas de un Plan de Planner específico, manejando paginación.

    Args:
        parametros (Dict[str, Any]): Debe contener 'plan_id'. Opcional: 'top' (int, default 100).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Un diccionario {'value': [lista_completa_de_tareas]}.
    """
    plan_id: Optional[str] = parametros.get("plan_id")
    top: int = int(parametros.get("top", 100))

    if not plan_id: raise ValueError("Parámetro 'plan_id' es requerido.")

    url_base = f"{BASE_URL}/planner/plans/{plan_id}/tasks"
    params_query: Dict[str, Any] = {'$top': min(top, 999)} # Limitar top por llamada

    all_tasks: List[Dict[str, Any]] = []
    current_url: Optional[str] = url_base
    page_count = 0
    max_pages = 100 # Límite de seguridad

    try:
        while current_url and page_count < max_pages:
            page_count += 1
            logger.info(f"Listando tareas Planner plan '{plan_id}', Página: {page_count}")

            current_params_page = params_query if page_count == 1 else None
            # Usar helper centralizado
            data = hacer_llamada_api("GET", current_url, headers, params=current_params_page)

            if data:
                tasks_in_page = data.get('value', [])
                all_tasks.extend(tasks_in_page)
                current_url = data.get('@odata.nextLink')
                if not current_url: break
            else:
                 logger.warning(f"Llamada a {current_url} para listar tareas Planner devolvió None/vacío.")
                 break

        if page_count >= max_pages:
             logger.warning(f"Se alcanzó límite de {max_pages} páginas listando tareas Planner del plan '{plan_id}'.")

        logger.info(f"Total tareas Planner listadas para plan '{plan_id}': {len(all_tasks)}")
        return {'value': all_tasks}

    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_tareas_planner (página {page_count}): {e}", exc_info=True)
        raise Exception(f"Error API listando tareas Planner: {e}") from e
    except Exception as e:
        logger.error(f"Error inesperado en listar_tareas_planner (página {page_count}): {e}", exc_info=True)
        raise


def crear_tarea_planner(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea una nueva tarea en un Plan de Planner.

    Args:
        parametros (Dict[str, Any]): Debe contener 'plan_id', 'titulo_tarea'.
                                     Opcional: 'bucket_id', 'detalles' (dict con campos adicionales como assignments).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El objeto de la tarea creada.
    """
    plan_id: Optional[str] = parametros.get("plan_id")
    titulo_tarea: Optional[str] = parametros.get("titulo_tarea")
    bucket_id: Optional[str] = parametros.get("bucket_id")
    # 'detalles' puede contener asignaciones, fechas, etc.
    detalles: Optional[Dict[str, Any]] = parametros.get("detalles")

    if not plan_id: raise ValueError("Parámetro 'plan_id' es requerido.")
    if not titulo_tarea: raise ValueError("Parámetro 'titulo_tarea' es requerido.")

    url = f"{BASE_URL}/planner/tasks"
    body: Dict[str, Any] = {"planId": plan_id, "title": titulo_tarea}
    if bucket_id: body["bucketId"] = bucket_id
    # Fusionar detalles si se proporcionan
    if detalles and isinstance(detalles, dict):
        body.update(detalles) # Añade/sobrescribe campos del body base

    logger.info(f"Creando tarea Planner '{titulo_tarea}' en plan '{plan_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)


def actualizar_tarea_planner(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Actualiza una tarea de Planner existente. Soporta ETag.

    Args:
        parametros (Dict[str, Any]): Debe contener 'tarea_id', 'nuevos_valores' (dict).
                                     Opcional: '@odata.etag' dentro de nuevos_valores.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El objeto de la tarea actualizada (o estado si devuelve 204).
    """
    tarea_id: Optional[str] = parametros.get("tarea_id")
    nuevos_valores: Optional[Dict[str, Any]] = parametros.get("nuevos_valores")

    if not tarea_id: raise ValueError("Parámetro 'tarea_id' es requerido.")
    if not nuevos_valores or not isinstance(nuevos_valores, dict):
        raise ValueError("Parámetro 'nuevos_valores' (diccionario) es requerido.")

    url = f"{BASE_URL}/planner/tasks/{tarea_id}"
    current_headers = headers.copy()
    body_data = nuevos_valores.copy()
    etag = body_data.pop('@odata.etag', None)

    if etag:
        current_headers['If-Match'] = etag
        logger.info(f"Usando ETag '{etag}' para actualizar tarea Planner {tarea_id}")
    else:
        logger.warning(f"Actualizando tarea Planner {tarea_id} sin ETag.")

    logger.info(f"Actualizando tarea Planner '{tarea_id}'")
    # PATCH puede devolver 204 o 200 OK
    result = hacer_llamada_api("PATCH", url, current_headers, json_data=body_data)

    if result is None:
         logger.warning(f"Actualizar tarea Planner {tarea_id} devolvió 204 No Content.")
         # Podríamos re-obtener la tarea, pero por ahora devolvemos status
         return {"status": "Actualizado (No Content)", "id": tarea_id}
    else:
         return result # Devolver el cuerpo si hubo 200 OK


def eliminar_tarea_planner(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina una tarea de Planner. Soporta ETag.

    Args:
        parametros (Dict[str, Any]): Debe contener 'tarea_id'. Opcional: 'etag'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de eliminación.
    """
    tarea_id: Optional[str] = parametros.get("tarea_id")
    etag: Optional[str] = parametros.get("etag")

    if not tarea_id: raise ValueError("Parámetro 'tarea_id' es requerido.")

    url = f"{BASE_URL}/planner/tasks/{tarea_id}"
    current_headers = headers.copy()
    if etag:
        current_headers['If-Match'] = etag
        logger.info(f"Eliminando tarea Planner {tarea_id} con ETag.")
    else:
        logger.warning(f"Eliminando tarea Planner {tarea_id} sin ETag.")

    # DELETE devuelve 204 No Content (None del helper).
    hacer_llamada_api("DELETE", url, current_headers)
    return {"status": "Eliminado", "id": tarea_id}


# =================================
# ==== FUNCIONES ACCIÓN TO-DO ====
# =================================
# Operan sobre /me/todo

def listar_listas_todo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista las listas de tareas de Microsoft To Do del usuario actual (/me).

    Args:
        parametros (Dict[str, Any]): Vacío o ignorado por ahora.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Respuesta de Graph API, usualmente {'value': [...]}.
    """
    url = f"{BASE_URL}/me/todo/lists"
    logger.info("Listando listas de ToDo para /me")
    # Podría añadirse paginación si se esperan muchas listas
    return hacer_llamada_api("GET", url, headers)


def crear_lista_todo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea una nueva lista de tareas en Microsoft To Do para el usuario actual.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_lista'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El objeto de la lista creada.
    """
    nombre_lista: Optional[str] = parametros.get("nombre_lista")
    if not nombre_lista: raise ValueError("Parámetro 'nombre_lista' es requerido.")

    url = f"{BASE_URL}/me/todo/lists"
    body = {"displayName": nombre_lista}
    logger.info(f"Creando lista de ToDo '{nombre_lista}' para /me")
    return hacer_llamada_api("POST", url, headers, json_data=body)


def actualizar_lista_todo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Actualiza una lista de To Do existente (ej. renombrar).

    Args:
        parametros (Dict[str, Any]): Debe contener 'lista_id', 'nuevos_valores' (dict).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El objeto de la lista actualizada.
    """
    lista_id: Optional[str] = parametros.get("lista_id")
    nuevos_valores: Optional[Dict[str, Any]] = parametros.get("nuevos_valores")

    if not lista_id: raise ValueError("Parámetro 'lista_id' es requerido.")
    if not nuevos_valores or not isinstance(nuevos_valores, dict):
        raise ValueError("Parámetro 'nuevos_valores' (diccionario) es requerido.")

    url = f"{BASE_URL}/me/todo/lists/{lista_id}"
    logger.info(f"Actualizando lista de ToDo '{lista_id}'")
    # Podría requerir ETag, pero no está documentado explícitamente para listas ToDo
    return hacer_llamada_api("PATCH", url, headers, json_data=nuevos_valores)


def eliminar_lista_todo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina una lista de To Do.

    Args:
        parametros (Dict[str, Any]): Debe contener 'lista_id'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de eliminación.
    """
    lista_id: Optional[str] = parametros.get("lista_id")
    if not lista_id: raise ValueError("Parámetro 'lista_id' es requerido.")

    url = f"{BASE_URL}/me/todo/lists/{lista_id}"
    logger.info(f"Eliminando lista de ToDo '{lista_id}'")
    # DELETE devuelve 204 No Content (None del helper).
    hacer_llamada_api("DELETE", url, headers)
    return {"status": "Eliminado", "id": lista_id}


def listar_tareas_todo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lista las tareas de una lista de To Do específica, manejando paginación.

    Args:
        parametros (Dict[str, Any]): Debe contener 'lista_id'. Opcional: 'top' (int, default 100).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Un diccionario {'value': [lista_completa_de_tareas]}.
    """
    lista_id: Optional[str] = parametros.get("lista_id")
    top: int = int(parametros.get("top", 100))

    if not lista_id: raise ValueError("Parámetro 'lista_id' es requerido.")

    url_base = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks"
    params_query: Dict[str, Any] = {'$top': min(top, 999)}

    all_tasks: List[Dict[str, Any]] = []
    current_url: Optional[str] = url_base
    page_count = 0
    max_pages = 100

    try:
        while current_url and page_count < max_pages:
            page_count += 1
            logger.info(f"Listando tareas ToDo lista '{lista_id}', Página: {page_count}")

            current_params_page = params_query if page_count == 1 else None
            data = hacer_llamada_api("GET", current_url, headers, params=current_params_page)

            if data:
                tasks_in_page = data.get('value', [])
                if isinstance(tasks_in_page, list): # Verificar que sea lista
                    all_tasks.extend(tasks_in_page)
                else:
                    logger.warning(f"Respuesta inesperada al listar tareas ToDo (no es lista): {type(tasks_in_page)}")
                current_url = data.get('@odata.nextLink')
                if not current_url: break
            else:
                 logger.warning(f"Llamada a {current_url} para listar tareas ToDo devolvió None/vacío.")
                 break

        if page_count >= max_pages:
             logger.warning(f"Se alcanzó límite de {max_pages} páginas listando tareas ToDo de lista '{lista_id}'.")

        logger.info(f"Total tareas ToDo listadas para lista '{lista_id}': {len(all_tasks)}")
        return {'value': all_tasks}

    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_tareas_todo (página {page_count}): {e}", exc_info=True)
        raise Exception(f"Error API listando tareas ToDo: {e}") from e
    except Exception as e:
        logger.error(f"Error inesperado en listar_tareas_todo (página {page_count}): {e}", exc_info=True)
        raise


def crear_tarea_todo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea una nueva tarea en una lista de To Do.

    Args:
        parametros (Dict[str, Any]): Debe contener 'lista_id', 'titulo_tarea'.
                                     Opcional: 'detalles' (str o dict con 'content' y 'contentType').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El objeto de la tarea creada.
    """
    lista_id: Optional[str] = parametros.get("lista_id")
    titulo_tarea: Optional[str] = parametros.get("titulo_tarea")
    detalles: Optional[Any] = parametros.get("detalles") # Puede ser string o dict

    if not lista_id: raise ValueError("Parámetro 'lista_id' es requerido.")
    if not titulo_tarea: raise ValueError("Parámetro 'titulo_tarea' es requerido.")

    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks"
    body: Dict[str, Any] = {"title": titulo_tarea}

    # Añadir cuerpo/detalles si se proporcionan
    if detalles is not None:
         if isinstance(detalles, str):
             # Si es string, asumir texto plano
             body['body'] = {"content": detalles, "contentType": "text"}
         elif isinstance(detalles, dict) and 'content' in detalles and 'contentType' in detalles:
             # Si es dict con formato correcto, usarlo
             body['body'] = detalles
         else:
             logger.warning(f"Formato inesperado para 'detalles' en crear_tarea_todo: {detalles}. Se ignorará el cuerpo.")

    logger.info(f"Creando tarea ToDo '{titulo_tarea}' en lista '{lista_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)


def actualizar_tarea_todo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Actualiza una tarea de To Do existente.

    Args:
        parametros (Dict[str, Any]): Debe contener 'lista_id', 'tarea_id', 'nuevos_valores' (dict).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El objeto de la tarea actualizada.
    """
    lista_id: Optional[str] = parametros.get("lista_id")
    tarea_id: Optional[str] = parametros.get("tarea_id")
    nuevos_valores: Optional[Dict[str, Any]] = parametros.get("nuevos_valores")

    if not lista_id: raise ValueError("Parámetro 'lista_id' es requerido.")
    if not tarea_id: raise ValueError("Parámetro 'tarea_id' es requerido.")
    if not nuevos_valores or not isinstance(nuevos_valores, dict):
        raise ValueError("Parámetro 'nuevos_valores' (diccionario) es requerido.")

    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks/{tarea_id}"
    logger.info(f"Actualizando tarea ToDo '{tarea_id}' en lista '{lista_id}'")
    # Podría requerir ETag, pero no documentado explícitamente
    return hacer_llamada_api("PATCH", url, headers, json_data=nuevos_valores)


def eliminar_tarea_todo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Elimina una tarea de To Do.

    Args:
        parametros (Dict[str, Any]): Debe contener 'lista_id', 'tarea_id'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Confirmación de eliminación.
    """
    lista_id: Optional[str] = parametros.get("lista_id")
    tarea_id: Optional[str] = parametros.get("tarea_id")

    if not lista_id: raise ValueError("Parámetro 'lista_id' es requerido.")
    if not tarea_id: raise ValueError("Parámetro 'tarea_id' es requerido.")

    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks/{tarea_id}"
    logger.info(f"Eliminando tarea ToDo '{tarea_id}' de lista '{lista_id}'")
    # DELETE devuelve 204 No Content (None del helper).
    hacer_llamada_api("DELETE", url, headers)
    return {"status": "Eliminado", "id": tarea_id}


def completar_tarea_todo(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Marca una tarea de To Do como completada.

    Args:
        parametros (Dict[str, Any]): Debe contener 'lista_id', 'tarea_id'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: El objeto de la tarea actualizada (marcada como completada).
    """
    lista_id: Optional[str] = parametros.get("lista_id")
    tarea_id: Optional[str] = parametros.get("tarea_id")

    if not lista_id: raise ValueError("Parámetro 'lista_id' es requerido.")
    if not tarea_id: raise ValueError("Parámetro 'tarea_id' es requerido.")

    logger.info(f"Marcando tarea ToDo '{tarea_id}' en lista '{lista_id}' como completada.")
    # Para completar, se actualiza el estado a 'completed'
    payload = {"status": "completed"}
    # Reutilizar la función de actualizar
    params_actualizar = {
        "lista_id": lista_id,
        "tarea_id": tarea_id,
        "nuevos_valores": payload
    }
    return actualizar_tarea_todo(params_actualizar, headers)

# --- FIN DEL MÓDULO actions/planner_todo.py ---
