import logging
import requests
import json # Necesario para algunos manejos de error
from typing import Dict, List, Optional, Union, Any
from datetime import datetime, timezone

# Usar el logger de la función principal
logger = logging.getLogger("azure.functions")

# Importar constantes globales desde __init__.py
try:
    from .. import BASE_URL, GRAPH_API_TIMEOUT
except ImportError:
    # Fallback por si se ejecuta standalone (poco probable en Azure Functions)
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_TIMEOUT = 45
    logger.warning("No se pudo importar BASE_URL/GRAPH_API_TIMEOUT desde el padre, usando defaults.")

# ---- PLANNER ----
# Todas las funciones ahora aceptan 'headers' como primer argumento

def listar_planes(headers: Dict[str, str], grupo_id: str) -> Dict[str, Any]:
    """Lista los planes de Planner de un grupo. Requiere headers autenticados."""
    url = f"{BASE_URL}/groups/{grupo_id}/planner/plans"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Listando planes grupo '{grupo_id}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Listados planes del grupo '{grupo_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_planes: {e}", exc_info=True)
        raise # Re-lanzar para que main() lo maneje
    except Exception as e:
        logger.error(f"Error inesperado en listar_planes: {e}", exc_info=True)
        raise

def obtener_plan(headers: Dict[str, str], plan_id: str) -> Dict[str, Any]:
    """Obtiene un plan de Planner específico. Requiere headers autenticados."""
    url = f"{BASE_URL}/planner/plans/{plan_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Obteniendo plan '{plan_id}')")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Plan ID '{plan_id}' obtenido.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en obtener_plan: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en obtener_plan: {e}", exc_info=True)
        raise

def crear_plan(headers: Dict[str, str], nombre_plan: str, grupo_id: str) -> Dict[str, Any]:
    """Crea un nuevo plan de Planner. Requiere headers autenticados."""
    url = f"{BASE_URL}/planner/plans"
    body = {"owner": grupo_id, "title": nombre_plan}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando plan '{nombre_plan}' para grupo '{grupo_id}')")
        # Asegurar Content-Type correcto si no está en headers (main debería ponerlo)
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        plan_id = data.get('id')
        logger.info(f"Plan '{nombre_plan}' creado ID: {plan_id}.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en crear_plan: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_plan: {e}", exc_info=True)
        raise

def actualizar_plan(headers: Dict[str, str], plan_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]:
    """Actualiza un plan de Planner existente. Requiere headers autenticados."""
    url = f"{BASE_URL}/planner/plans/{plan_id}"
    response: Optional[requests.Response] = None
    try:
        etag = nuevos_valores.pop('@odata.etag', None)
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        if etag:
            current_headers['If-Match'] = etag
            logger.info(f"Usando ETag para plan {plan_id}")

        logger.info(f"API Call: PATCH {url} (Actualizando plan '{plan_id}')")
        response = requests.patch(url, headers=current_headers, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        logger.info(f"Plan ID '{plan_id}' actualizado.")
        if response.status_code == 204:
             logger.warning(f"Actualizar plan {plan_id} devolvió 204 No Content. Re-obteniendo plan.")
             # Para devolver el plan actualizado, necesitamos llamar a obtener_plan
             # Pasamos los mismos headers originales
             return obtener_plan(headers=headers, plan_id=plan_id)
        else:
             # Asumiendo 200 OK con cuerpo
             return response.json()
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en actualizar_plan: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en actualizar_plan: {e}", exc_info=True)
        raise

def eliminar_plan(headers: Dict[str, str], plan_id: str) -> Dict[str, Any]:
    """Elimina un plan de Planner. Requiere headers autenticados."""
    url = f"{BASE_URL}/planner/plans/{plan_id}"
    response: Optional[requests.Response] = None
    try:
        # Añadir ETag si se tiene para eliminación segura
        # etag = obtener_etag_plan(headers, plan_id) # Necesitaría una función helper
        # current_headers = headers.copy()
        # if etag: current_headers['If-Match'] = etag
        # else: logger.warning(f"Eliminando plan {plan_id} sin ETag.")

        logger.warning(f"Eliminando plan {plan_id} sin ETag (refactorización simple).")
        logger.info(f"API Call: DELETE {url} (Eliminando plan '{plan_id}')")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 204
        logger.info(f"Plan ID '{plan_id}' eliminado.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en eliminar_plan: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_plan: {e}", exc_info=True)
        raise

def listar_tareas_planner(headers: Dict[str, str], plan_id: str) -> Dict[str, Any]:
    """Lista las tareas de un plan de Planner. Requiere headers autenticados."""
    url: Optional[str] = f"{BASE_URL}/planner/plans/{plan_id}/tasks"
    all_tasks: List[Dict[str, Any]] = []
    response: Optional[requests.Response] = None
    current_headers = headers.copy() # Copia por si se modifica en bucle (aunque no debería aquí)

    try:
        page_count = 0
        while url:
            page_count += 1
            logger.info(f"API Call: GET {url} Page: {page_count} (Listando tareas planner plan '{plan_id}')")
            response = requests.get(url, headers=current_headers, timeout=GRAPH_API_TIMEOUT)
            response.raise_for_status()
            data = response.json()
            tasks_in_page = data.get('value', [])
            if isinstance(tasks_in_page, list): all_tasks.extend(tasks_in_page)
            else: logger.warning(f"Se esperaba lista en 'value' (tareas planner): {type(tasks_in_page)}")
            url = data.get('@odata.nextLink')
            # No es necesario actualizar headers aquí para Graph API
        logger.info(f"Listadas tareas del plan ID '{plan_id}'. Total: {len(all_tasks)}")
        return {'value': all_tasks}
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_tareas_planner: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_tareas_planner: {e}", exc_info=True)
        raise

def crear_tarea_planner(headers: Dict[str, str], plan_id: str, titulo_tarea: str, bucket_id: Optional[str] = None, detalles: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Crea una tarea de Planner. Requiere headers autenticados."""
    url = f"{BASE_URL}/planner/tasks"
    body: Dict[str, Any] = {"planId": plan_id, "title": titulo_tarea}
    if bucket_id: body["bucketId"] = bucket_id
    if detalles and isinstance(detalles, dict): body.update(detalles) # Asignar otros campos como 'assignments'
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando tarea planner '{titulo_tarea}' en plan '{plan_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        task_id = data.get('id')
        logger.info(f"Tarea Planner '{titulo_tarea}' creada ID: {task_id}.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en crear_tarea_planner: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_tarea_planner: {e}", exc_info=True)
        raise

def actualizar_tarea_planner(headers: Dict[str, str], tarea_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]:
    """Actualiza una tarea de Planner. Requiere headers autenticados."""
    url = f"{BASE_URL}/planner/tasks/{tarea_id}"
    response: Optional[requests.Response] = None
    try:
        etag = nuevos_valores.pop('@odata.etag', None) # Usar ETag de la tarea si se tiene
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        if etag:
             current_headers['If-Match'] = etag
             logger.info(f"Usando ETag Tarea {tarea_id}")
        else:
             logger.warning(f"Actualizando tarea planner {tarea_id} sin ETag.")

        logger.info(f"API Call: PATCH {url} (Actualizando tarea planner '{tarea_id}')")
        response = requests.patch(url, headers=current_headers, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        logger.info(f"Tarea ID '{tarea_id}' actualizada.")
        if response.status_code == 204:
             logger.warning(f"Actualizar tarea {tarea_id} devolvió 204.")
             # Necesitaríamos obtener_tarea_planner(headers, tarea_id) para devolverla
             return {"status": "Actualizado (No Content)", "id": tarea_id}
        else:
             return response.json() # Asume 200 OK con cuerpo
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en actualizar_tarea_planner: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en actualizar_tarea_planner: {e}", exc_info=True)
        raise

def eliminar_tarea_planner(headers: Dict[str, str], tarea_id: str) -> Dict[str, Any]:
    """Elimina una tarea de Planner. Requiere headers autenticados."""
    url = f"{BASE_URL}/planner/tasks/{tarea_id}"
    response: Optional[requests.Response] = None
    try:
        # Añadir ETag si se tiene
        # etag = obtener_etag_tarea(headers, tarea_id)
        # current_headers = headers.copy()
        # if etag: current_headers['If-Match'] = etag
        # else: logger.warning(f"Eliminando tarea planner {tarea_id} sin ETag.")
        logger.warning(f"Eliminando tarea planner {tarea_id} sin ETag (refactorización simple).")
        logger.info(f"API Call: DELETE {url} (Eliminando tarea planner '{tarea_id}')")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 204
        logger.info(f"Tarea ID '{tarea_id}' eliminada.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en eliminar_tarea_planner: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_tarea_planner: {e}", exc_info=True)
        raise

# ---- TO DO ----
# Todas las funciones ahora aceptan 'headers' y usan /me

def listar_listas_todo(headers: Dict[str, str]) -> Dict[str, Any]:
    """Lista las listas de To Do del usuario (/me). Requiere headers autenticados."""
    url = f"{BASE_URL}/me/todo/lists"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: GET {url} (Listando listas ToDo /me)")
        response = requests.get(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info("Listadas listas To Do.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_listas_todo: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_listas_todo: {e}", exc_info=True)
        raise

def crear_lista_todo(headers: Dict[str, str], nombre_lista: str) -> Dict[str, Any]:
    """Crea una lista de To Do para el usuario (/me). Requiere headers autenticados."""
    url = f"{BASE_URL}/me/todo/lists"
    body = {"displayName": nombre_lista}
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando lista ToDo '{nombre_lista}' para /me)")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        list_id = data.get('id')
        logger.info(f"Lista ToDo '{nombre_lista}' creada ID: {list_id}.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en crear_lista_todo: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_lista_todo: {e}", exc_info=True)
        raise

def actualizar_lista_todo(headers: Dict[str, str], lista_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]:
    """Actualiza una lista de To Do (/me). Requiere headers autenticados."""
    url = f"{BASE_URL}/me/todo/lists/{lista_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Actualizando lista ToDo '{lista_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        # ToDo API no usa ETags para listas/tareas que yo sepa
        response = requests.patch(url, headers=current_headers, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Lista To Do ID '{lista_id}' actualizada.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en actualizar_lista_todo: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en actualizar_lista_todo: {e}", exc_info=True)
        raise

def eliminar_lista_todo(headers: Dict[str, str], lista_id: str) -> Dict[str, Any]:
    """Elimina una lista de To Do (/me). Requiere headers autenticados."""
    url = f"{BASE_URL}/me/todo/lists/{lista_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando lista ToDo '{lista_id}')")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 204
        logger.info(f"Lista To Do ID '{lista_id}' eliminada.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en eliminar_lista_todo: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_lista_todo: {e}", exc_info=True)
        raise

def listar_tareas_todo(headers: Dict[str, str], lista_id: str) -> Dict[str, Any]:
    """Lista las tareas de una lista de To Do específica (/me). Requiere headers autenticados."""
    url: Optional[str] = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks"
    all_tasks: List[Dict[str, Any]] = []
    response: Optional[requests.Response] = None
    current_headers = headers.copy()

    try:
        page_count = 0
        while url:
            page_count += 1
            logger.info(f"API Call: GET {url} Page: {page_count} (Listando tareas ToDo lista '{lista_id}')")
            response = requests.get(url, headers=current_headers, timeout=GRAPH_API_TIMEOUT)
            response.raise_for_status()
            data = response.json()
            tasks_in_page = data.get('value', [])
            if isinstance(tasks_in_page, list): all_tasks.extend(tasks_in_page)
            else: logger.warning(f"Se esperaba lista en 'value' (tareas todo): {type(tasks_in_page)}")
            url = data.get('@odata.nextLink')
            # No es necesario refrescar token para Graph paginación
        logger.info(f"Listadas tareas lista To Do ID '{lista_id}'. Total: {len(all_tasks)}")
        return {'value': all_tasks}
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en listar_tareas_todo: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en listar_tareas_todo: {e}", exc_info=True)
        raise

def crear_tarea_todo(headers: Dict[str, str], lista_id: str, titulo_tarea: str, detalles: Optional[Any] = None) -> Dict[str, Any]:
    """Crea una tarea de To Do (/me). Requiere headers autenticados."""
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks"
    body: Dict[str, Any] = {"title": titulo_tarea}
    if detalles is not None:
         # Asumimos que 'detalles' es el contenido del cuerpo de la nota
         if isinstance(detalles, str):
             body['body'] = {"content": detalles, "contentType": "text"} # o html si se pasa html
         elif isinstance(detalles, dict) and 'content' in detalles: # Si ya viene el objeto body
             body['body'] = detalles
         else: logger.warning(f"Formato inesperado para 'detalles' en crear_tarea_todo: {detalles}")
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: POST {url} (Creando tarea ToDo '{titulo_tarea}' en lista '{lista_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.post(url, headers=current_headers, json=body, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        task_id = data.get('id')
        logger.info(f"Tarea ToDo '{titulo_tarea}' creada ID: {task_id}.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en crear_tarea_todo: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en crear_tarea_todo: {e}", exc_info=True)
        raise

def actualizar_tarea_todo(headers: Dict[str, str], lista_id: str, tarea_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]:
    """Actualiza una tarea de To Do (/me). Requiere headers autenticados."""
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks/{tarea_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: PATCH {url} (Actualizando tarea ToDo '{tarea_id}')")
        current_headers = headers.copy()
        current_headers.setdefault('Content-Type', 'application/json')
        response = requests.patch(url, headers=current_headers, json=nuevos_valores, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        logger.info(f"Tarea ToDo ID '{tarea_id}' actualizada.")
        return data
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en actualizar_tarea_todo: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en actualizar_tarea_todo: {e}", exc_info=True)
        raise

def eliminar_tarea_todo(headers: Dict[str, str], lista_id: str, tarea_id: str) -> Dict[str, Any]:
    """Elimina una tarea de To Do (/me). Requiere headers autenticados."""
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks/{tarea_id}"
    response: Optional[requests.Response] = None
    try:
        logger.info(f"API Call: DELETE {url} (Eliminando tarea ToDo '{tarea_id}')")
        response = requests.delete(url, headers=headers, timeout=GRAPH_API_TIMEOUT)
        response.raise_for_status() # Espera 204
        logger.info(f"Tarea ToDo ID '{tarea_id}' eliminada.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logger.error(f"Error Request en eliminar_tarea_todo: {e}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error inesperado en eliminar_tarea_todo: {e}", exc_info=True)
        raise

def completar_tarea_todo(headers: Dict[str, str], lista_id: str, tarea_id: str) -> dict:
    """Marca una tarea de To Do como completada (/me). Requiere headers autenticados."""
    logger.info(f"Completando tarea ToDo ID '{tarea_id}'.")
    payload = {"status": "completed"}
    # Llama a la función refactorizada de actualizar
    return actualizar_tarea_todo(headers=headers, lista_id=lista_id, tarea_id=tarea_id, nuevos_valores=payload)

# --- FIN: Funciones Auxiliares ---
