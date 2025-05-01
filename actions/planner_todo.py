# actions/planner_todo.py (Refactorizado v2 con Helper)

import logging
import requests
import json
# Corregido: Añadir Any
from typing import Dict, List, Optional, Union, Any
from datetime import datetime, timezone

# Usar logger principal
logger = logging.getLogger("azure.functions")

# Importar helper y constantes
try:
    from helpers.http_client import hacer_llamada_api
    from shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    logger.error("Error importando helpers/constantes en Planner/ToDo.")
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs): raise NotImplementedError("Helper no importado")

# ---- PLANNER ----
def listar_planes(headers: Dict[str, str], grupo_id: str) -> Dict[str, Any]:
    """Lista planes de un Grupo."""
    url = f"{BASE_URL}/groups/{grupo_id}/planner/plans"; logger.info(f"Listando planes grupo '{grupo_id}'")
    return hacer_llamada_api("GET", url, headers)

def obtener_plan(headers: Dict[str, str], plan_id: str) -> Dict[str, Any]:
    """Obtiene un Plan."""
    url = f"{BASE_URL}/planner/plans/{plan_id}"; logger.info(f"Obteniendo plan '{plan_id}'")
    return hacer_llamada_api("GET", url, headers)

def crear_plan(headers: Dict[str, str], nombre_plan: str, grupo_id: str) -> Dict[str, Any]:
    """Crea un Plan en un Grupo."""
    url = f"{BASE_URL}/planner/plans"; body = {"owner": grupo_id, "title": nombre_plan}; logger.info(f"Creando plan '{nombre_plan}' para grupo '{grupo_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def actualizar_plan(headers: Dict[str, str], plan_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]:
    """Actualiza un Plan."""
    url = f"{BASE_URL}/planner/plans/{plan_id}"; etag = nuevos_valores.pop('@odata.etag', None); current_headers = headers.copy()
    if etag: current_headers['If-Match'] = etag; logger.info(f"Usando ETag para plan {plan_id}")
    else: logger.warning(f"Actualizando plan {plan_id} sin ETag.")
    logger.info(f"Actualizando plan '{plan_id}'")
    # El helper devuelve JSON o None (si 204). Si es None, re-obtener.
    result = hacer_llamada_api("PATCH", url, current_headers, json_data=nuevos_valores)
    if result is None: # Implica 204 No Content
         logger.warning(f"Actualizar plan {plan_id} devolvió 204. Re-obteniendo plan.")
         return obtener_plan(headers=headers, plan_id=plan_id)
    return result

def eliminar_plan(headers: Dict[str, str], plan_id: str, etag: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """Elimina un Plan."""
    url = f"{BASE_URL}/planner/plans/{plan_id}"; current_headers = headers.copy()
    if etag: current_headers['If-Match'] = etag; logger.info(f"Eliminando plan {plan_id} con ETag.")
    else: logger.warning(f"Eliminando plan {plan_id} sin ETag.")
    hacer_llamada_api("DELETE", url, current_headers) # Devuelve None si 204
    return {"status": "Eliminado", "id": plan_id}

def listar_tareas_planner(headers: Dict[str, str], plan_id: str, top: int = 100) -> Dict[str, Any]:
    """Lista tareas de un Plan (con paginación)."""
    url_base = f"{BASE_URL}/planner/plans/{plan_id}/tasks"; params: Dict[str, Any] = {'$top': min(int(top), 999)}; all_tasks: List[Dict[str, Any]] = []; current_url: Optional[str] = url_base; current_headers = headers.copy(); response: Optional[requests.Response] = None
    try: # Paginación con requests directo
        page_count = 0
        while current_url:
            page_count += 1; logger.info(f"API Call: GET {current_url} Page: {page_count} (Listando tareas planner plan '{plan_id}')")
            current_params = params if page_count == 1 else None; assert current_url is not None
            response = requests.get(current_url, headers=current_headers, params=current_params, timeout=GRAPH_API_TIMEOUT)
            response.raise_for_status(); data = response.json(); tasks_in_page = data.get('value', []); all_tasks.extend(tasks_in_page)
            current_url = data.get('@odata.nextLink')
        logger.info(f"Listadas tareas del plan ID '{plan_id}'. Total: {len(all_tasks)}"); return {'value': all_tasks}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en listar_tareas_planner: {e}", exc_info=True); raise Exception(f"Error API listando tareas planner: {e}")
    except Exception as e: logger.error(f"Error inesperado en listar_tareas_planner: {e}", exc_info=True); raise

def crear_tarea_planner(headers: Dict[str, str], plan_id: str, titulo_tarea: str, bucket_id: Optional[str] = None, detalles: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Crea una tarea de Planner."""
    url = f"{BASE_URL}/planner/tasks"; body: Dict[str, Any] = {"planId": plan_id, "title": titulo_tarea};
    if bucket_id: body["bucketId"] = bucket_id
    if detalles and isinstance(detalles, dict): body.update(detalles)
    logger.info(f"Creando tarea planner '{titulo_tarea}' en plan '{plan_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def actualizar_tarea_planner(headers: Dict[str, str], tarea_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]:
    """Actualiza una tarea de Planner."""
    url = f"{BASE_URL}/planner/tasks/{tarea_id}"; etag = nuevos_valores.pop('@odata.etag', None); current_headers = headers.copy();
    if etag: current_headers['If-Match'] = etag; logger.info(f"Usando ETag Tarea {tarea_id}")
    else: logger.warning(f"Actualizando tarea planner {tarea_id} sin ETag.")
    logger.info(f"Actualizando tarea planner '{tarea_id}'")
    result = hacer_llamada_api("PATCH", url, current_headers, json_data=nuevos_valores)
    if result is None: logger.warning(f"Actualizar tarea {tarea_id} devolvió 204."); return {"status": "Actualizado (No Content)", "id": tarea_id}
    else: return result # Asume 200 OK con body

def eliminar_tarea_planner(headers: Dict[str, str], tarea_id: str, etag: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """Elimina una tarea de Planner."""
    url = f"{BASE_URL}/planner/tasks/{tarea_id}"; current_headers = headers.copy();
    if etag: current_headers['If-Match'] = etag; logger.info(f"Eliminando tarea {tarea_id} con ETag.")
    else: logger.warning(f"Eliminando tarea planner {tarea_id} sin ETag.")
    hacer_llamada_api("DELETE", url, current_headers) # Devuelve None si 204
    return {"status": "Eliminado", "id": tarea_id}

# ---- TO DO ----
def listar_listas_todo(headers: Dict[str, str]) -> Dict[str, Any]:
    """Lista las listas de ToDo del usuario (/me)."""
    url = f"{BASE_URL}/me/todo/lists"; logger.info("Listando listas ToDo /me")
    return hacer_llamada_api("GET", url, headers)

def crear_lista_todo(headers: Dict[str, str], nombre_lista: str) -> Dict[str, Any]:
    """Crea una lista de To Do para el usuario (/me)."""
    url = f"{BASE_URL}/me/todo/lists"; body = {"displayName": nombre_lista}; logger.info(f"Creando lista ToDo '{nombre_lista}' para /me")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def actualizar_lista_todo(headers: Dict[str, str], lista_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]:
    """Actualiza una lista de To Do (/me)."""
    url = f"{BASE_URL}/me/todo/lists/{lista_id}"; logger.info(f"Actualizando lista ToDo '{lista_id}'")
    # Podría requerir ETag
    return hacer_llamada_api("PATCH", url, headers, json_data=nuevos_valores)

def eliminar_lista_todo(headers: Dict[str, str], lista_id: str) -> Optional[Dict[str, Any]]:
    """Elimina una lista de To Do (/me)."""
    url = f"{BASE_URL}/me/todo/lists/{lista_id}"; logger.info(f"Eliminando lista ToDo '{lista_id}'")
    hacer_llamada_api("DELETE", url, headers) # Devuelve None si 204
    return {"status": "Eliminado", "id": lista_id}

def listar_tareas_todo(headers: Dict[str, str], lista_id: str, top: int = 100) -> Dict[str, Any]:
    """Lista las tareas de una lista de To Do específica (/me)."""
    url_base = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks"; params: Dict[str, Any] = {'$top': min(int(top), 999)}; all_tasks: List[Dict[str, Any]] = []; current_url: Optional[str] = url_base; response: Optional[requests.Response] = None; current_headers = headers.copy()
    try: # Paginación con requests directo
        page_count = 0
        while current_url:
            page_count += 1; logger.info(f"API Call: GET {current_url} Page: {page_count} (Listando tareas ToDo lista '{lista_id}')")
            current_params = params if page_count == 1 else None; assert current_url is not None
            response = requests.get(current_url, headers=current_headers, params=current_params, timeout=GRAPH_API_TIMEOUT)
            response.raise_for_status(); data = response.json(); tasks_in_page = data.get('value', []);
            if isinstance(tasks_in_page, list): all_tasks.extend(tasks_in_page)
            else: logger.warning(f"Se esperaba lista en 'value' (tareas todo): {type(tasks_in_page)}")
            current_url = data.get('@odata.nextLink')
        logger.info(f"Listadas tareas lista To Do ID '{lista_id}'. Total: {len(all_tasks)}"); return {'value': all_tasks}
    except requests.exceptions.RequestException as e: logger.error(f"Error Request en listar_tareas_todo: {e}", exc_info=True); raise Exception(f"Error API listando tareas ToDo: {e}")
    except Exception as e: logger.error(f"Error inesperado en listar_tareas_todo: {e}", exc_info=True); raise

def crear_tarea_todo(headers: Dict[str, str], lista_id: str, titulo_tarea: str, detalles: Optional[Any] = None) -> Dict[str, Any]:
    """Crea una tarea de To Do (/me)."""
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks"; body: Dict[str, Any] = {"title": titulo_tarea};
    if detalles is not None:
         if isinstance(detalles, str): body['body'] = {"content": detalles, "contentType": "text"}
         elif isinstance(detalles, dict) and 'content' in detalles: body['body'] = detalles
         else: logger.warning(f"Formato inesperado para 'detalles' en crear_tarea_todo: {detalles}")
    logger.info(f"Creando tarea ToDo '{titulo_tarea}' en lista '{lista_id}'")
    return hacer_llamada_api("POST", url, headers, json_data=body)

def actualizar_tarea_todo(headers: Dict[str, str], lista_id: str, tarea_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]:
    """Actualiza una tarea de To Do (/me)."""
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks/{tarea_id}"; logger.info(f"Actualizando tarea ToDo '{tarea_id}'")
    # Podría requerir ETag
    return hacer_llamada_api("PATCH", url, headers, json_data=nuevos_valores)

def eliminar_tarea_todo(headers: Dict[str, str], lista_id: str, tarea_id: str) -> Optional[Dict[str, Any]]:
    """Elimina una tarea de To Do (/me)."""
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks/{tarea_id}"; logger.info(f"Eliminando tarea ToDo '{tarea_id}'")
    hacer_llamada_api("DELETE", url, headers) # Devuelve None si 204
    return {"status": "Eliminado", "id": tarea_id}

def completar_tarea_todo(headers: Dict[str, str], lista_id: str, tarea_id: str) -> dict:
    """Marca una tarea de To Do como completada (/me)."""
    logger.info(f"Completando tarea ToDo ID '{tarea_id}'.")
    payload = {"status": "completed"}
    return actualizar_tarea_todo(headers=headers, lista_id=lista_id, tarea_id=tarea_id, nuevos_valores=payload)
