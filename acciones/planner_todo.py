import logging
import os
import requests
from auth import obtener_token # Asume que auth.py está en la raíz
from typing import Dict, List, Optional, Union, Any
from datetime import datetime, timezone # Añadido timezone
import json

# Configuración básica de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- INICIO: Configuración Redundante ---
# (Código omitido por brevedad - igual que antes)
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
GRAPH_SCOPE = os.getenv('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')
if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE]):
    logging.error("❌ Faltan variables de entorno en planner_todo.")
    # Considerar no lanzar excepción aquí

BASE_URL = "https://graph.microsoft.com/v1.0"
HEADERS: Dict[str, Optional[str]] = {
    'Authorization': None,
    'Content-Type': 'application/json'
}

def _actualizar_headers() -> None:
    try:
        token = obtener_token()
        HEADERS['Authorization'] = f'Bearer {token}'
        logging.info("Headers actualizados en planner_todo.")
    except Exception as e:
        logging.error(f"❌ Error al obtener el token en planner_todo: {e}")
        raise Exception(f"Error al obtener el token en planner_todo: {e}")
# --- FIN: Configuración Redundante ---


# ---- PLANNER ----

def listar_planes(grupo_id: str) -> Dict[str, Any]:
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{grupo_id}/planner/plans"; response: Optional[requests.Response] = None
    try:
        response = requests.get(url, headers=HEADERS); response.raise_for_status(); data = response.json()
        logging.info(f"Listados planes del grupo '{grupo_id}'.")
        return data
    except requests.exceptions.RequestException as e: logging.error(f"❌ Error listar planes: {e}"); raise Exception(f"Error listar planes: {e}")
    except json.JSONDecodeError as e: logging.error(f"❌ Error JSON (listar planes): {e}"); raise Exception(f"Error JSON (listar planes): {e}")

# --- CORRECCIÓN: 'obtener_plan' definida ANTES de 'actualizar_plan' ---
def obtener_plan(plan_id: str) -> Dict[str, Any]:
    """Obtiene un plan de Planner específico."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/plans/{plan_id}"
    response: Optional[requests.Response] = None
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        # Añadir log
        logging.info(f"Plan ID '{plan_id}' obtenido.")
        return response.json()
    except requests.exceptions.RequestException as e:
        error_details = getattr(e.response, 'text', str(e))
        logging.error(f"❌ Error obtener plan {plan_id}: {e}. Detalles: {error_details}")
        raise Exception(f"Error al obtener plan con ID {plan_id}: {e}")
    except json.JSONDecodeError as e:
        response_text = getattr(response, 'text', 'No response object available')
        logging.error(f"❌ Error JSON (obtener plan): {e}. Respuesta: {response_text}")
        raise Exception(f"Error al decodificar JSON (obtener plan): {e}")

def crear_plan(nombre_plan: str, grupo_id: str) -> Dict[str, Any]:
    _actualizar_headers()
    url = f"{BASE_URL}/planner/plans"; body = {"owner": grupo_id, "title": nombre_plan}; response: Optional[requests.Response] = None
    try:
        response = requests.post(url, headers=HEADERS, json=body); response.raise_for_status(); data = response.json()
        plan_id = data.get('id'); logging.info(f"Plan '{nombre_plan}' creado ID: {plan_id}."); return data
    except requests.exceptions.RequestException as e: logging.error(f"❌ Error crear plan: {e}"); raise Exception(f"Error crear plan: {e}")
    except json.JSONDecodeError as e: logging.error(f"❌ Error JSON (crear plan): {e}"); raise Exception(f"Error JSON (crear plan): {e}")

def actualizar_plan(plan_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]:
    """Actualiza un plan de Planner existente."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/plans/{plan_id}"
    response: Optional[requests.Response] = None
    try:
        etag = nuevos_valores.pop('@odata.etag', None); current_headers = HEADERS.copy()
        if etag: current_headers['If-Match'] = etag; logging.info(f"Usando ETag para plan {plan_id}")

        response = requests.patch(url, headers=current_headers, json=nuevos_valores); response.raise_for_status()
        logging.info(f"Plan ID '{plan_id}' actualizado.")
        if response.status_code == 204:
             logging.warning(f"Actualizar plan {plan_id} devolvió 204 No Content. Re-obteniendo plan.")
             # Ahora la llamada a obtener_plan es válida porque está definida antes
             return obtener_plan(plan_id)
        else:
             return response.json()
    except requests.exceptions.RequestException as e: logging.error(f"❌ Error actualizar plan {plan_id}: {e}"); raise Exception(f"Error actualizar plan: {e}")
    except json.JSONDecodeError as e: logging.error(f"❌ Error JSON (actualizar plan): {e}"); raise Exception(f"Error JSON (actualizar plan): {e}")

def eliminar_plan(plan_id: str) -> Dict[str, Any]:
     _actualizar_headers(); url = f"{BASE_URL}/planner/plans/{plan_id}"
     response: Optional[requests.Response] = None
     try:
         logging.warning(f"Eliminando plan {plan_id} sin ETag."); response = requests.delete(url, headers=HEADERS); response.raise_for_status()
         logging.info(f"Plan ID '{plan_id}' eliminado."); return {"status": "Eliminado", "code": response.status_code}
     except requests.exceptions.RequestException as e: logging.error(f"❌ Error eliminar plan {plan_id}: {e}"); raise Exception(f"Error eliminar plan: {e}")


def listar_tareas_planner(plan_id: str) -> Dict[str, Any]:
    """Lista las tareas de un plan de Planner específico, manejando paginación."""
    _actualizar_headers()
    url: Optional[str] = f"{BASE_URL}/planner/plans/{plan_id}/tasks"
    all_tasks: List[Dict[str, Any]] = []
    response: Optional[requests.Response] = None
    try:
        while url:
            logging.info(f"Obteniendo página de tareas planner desde: {url}") # Corregido a logging.
            response = requests.get(url, headers=HEADERS); response.raise_for_status()
            data: Dict[str, Any] = response.json()
            tasks_in_page = data.get('value', [])
            if isinstance(tasks_in_page, list): all_tasks.extend(tasks_in_page)
            else: logging.warning(f"Se esperaba lista en 'value' (tareas planner): {type(tasks_in_page)}") # Corregido a logging.
            url = data.get('@odata.nextLink')
            if url: _actualizar_headers()
        logging.info(f"Listadas tareas del plan ID '{plan_id}'. Total: {len(all_tasks)}") # Corregido a logging.
        return {'value': all_tasks}
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error listar tareas planner {plan_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error listar tareas planner: {e}") # logging.error
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (listar tareas planner): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar tareas planner): {e}") # logging.error


def crear_tarea_planner(plan_id: str, titulo_tarea: str, bucket_id: Optional[str] = None, detalles: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/planner/tasks"; body: Dict[str, Any] = {"planId": plan_id, "title": titulo_tarea}
    if bucket_id: body["bucketId"] = bucket_id
    if detalles and isinstance(detalles, dict): body.update(detalles)
    response: Optional[requests.Response] = None
    try:
        response = requests.post(url, headers=HEADERS, json=body); response.raise_for_status(); data = response.json()
        task_id = data.get('id'); logging.info(f"Tarea Planner '{titulo_tarea}' creada ID: {task_id}."); return data # logging.info
    except requests.exceptions.RequestException as e: logging.error(f"❌ Error crear tarea planner: {e}"); raise Exception(f"Error crear tarea planner: {e}")
    except json.JSONDecodeError as e: logging.error(f"❌ Error JSON (crear tarea planner): {e}"); raise Exception(f"Error JSON (crear tarea planner): {e}")


def actualizar_tarea_planner(tarea_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]:
     _actualizar_headers(); url = f"{BASE_URL}/planner/tasks/{tarea_id}"
     response: Optional[requests.Response] = None
     try:
         etag = nuevos_valores.pop('@odata.etag', None); current_headers = HEADERS.copy()
         if etag: current_headers['If-Match'] = etag; logging.info(f"Usando ETag Tarea {tarea_id}") # logging.info
         response = requests.patch(url, headers=current_headers, json=nuevos_valores); response.raise_for_status()
         logging.info(f"Tarea ID '{tarea_id}' actualizada.") # logging.info
         if response.status_code == 204:
             logging.warning(f"Tarea {tarea_id} devolvió 204."); # logging.warning
             # Necesitaríamos una función obtener_tarea_planner(tarea_id) para devolverla actualizada
             # return obtener_tarea_planner(tarea_id)
             return {"status": "Actualizado (No Content)", "id": tarea_id}
         else: return response.json()
     except requests.exceptions.RequestException as e: logging.error(f"❌ Error actualizar tarea planner {tarea_id}: {e}"); raise Exception(f"Error actualizar tarea planner: {e}")
     except json.JSONDecodeError as e: logging.error(f"❌ Error JSON (actualizar tarea planner): {e}"); raise Exception(f"Error JSON (actualizar tarea planner): {e}")


def eliminar_tarea_planner(tarea_id: str) -> Dict[str, Any]:
     _actualizar_headers(); url = f"{BASE_URL}/planner/tasks/{tarea_id}"
     response: Optional[requests.Response] = None
     try:
         logging.warning(f"Eliminando tarea planner {tarea_id} sin ETag.") # logging.warning
         response = requests.delete(url, headers=HEADERS); response.raise_for_status()
         logging.info(f"Tarea ID '{tarea_id}' eliminada.") # logging.info
         return {"status": "Eliminado", "code": response.status_code}
     except requests.exceptions.RequestException as e: logging.error(f"❌ Error eliminar tarea planner {tarea_id}: {e}"); raise Exception(f"Error eliminar tarea planner: {e}")


# ---- TO DO ----
# (Funciones listar_listas_todo, crear_lista_todo, etc. usando logging.)
def listar_listas_todo() -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/me/todo/lists"; response: Optional[requests.Response] = None
    try:
        response = requests.get(url, headers=HEADERS); response.raise_for_status(); data = response.json()
        logging.info("Listadas listas To Do.") # logging.info
        return data
    except requests.exceptions.RequestException as e: logging.error(f"❌ Error listar listas todo: {e}"); raise Exception(f"Error listar listas todo: {e}")
    except json.JSONDecodeError as e: logging.error(f"❌ Error JSON (list todo): {e}"); raise Exception(f"Error JSON (list todo): {e}")

def crear_lista_todo(nombre_lista: str) -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/me/todo/lists"; body = {"displayName": nombre_lista}; response: Optional[requests.Response] = None
    try:
        response = requests.post(url, headers=HEADERS, json=body); response.raise_for_status(); data = response.json()
        list_id = data.get('id'); logging.info(f"Lista ToDo '{nombre_lista}' creada ID: {list_id}."); return data # logging.info
    except requests.exceptions.RequestException as e: logging.error(f"❌ Error crear lista todo: {e}"); raise Exception(f"Error crear lista todo: {e}")
    except json.JSONDecodeError as e: logging.error(f"❌ Error JSON (crear lista todo): {e}"); raise Exception(f"Error JSON (crear lista todo): {e}")

# ... (Actualizar, Eliminar, Listar Tareas, Crear Tarea, etc. para ToDo - sin cambios estructurales, solo verificar logging) ...

def actualizar_lista_todo(lista_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]:
     _actualizar_headers(); url = f"{BASE_URL}/me/todo/lists/{lista_id}"; response: Optional[requests.Response] = None
     try:
         response = requests.patch(url, headers=HEADERS, json=nuevos_valores); response.raise_for_status(); data = response.json()
         logging.info(f"Lista To Do ID '{lista_id}' actualizada."); return data # logging.info
     except requests.exceptions.RequestException as e: logging.error(f"❌ Error actualizar lista todo {lista_id}: {e}"); raise Exception(f"Error actualizar lista todo: {e}")
     except json.JSONDecodeError as e: logging.error(f"❌ Error JSON (actualizar lista todo): {e}"); raise Exception(f"Error JSON (actualizar lista todo): {e}")

def eliminar_lista_todo(lista_id: str) -> Dict[str, Any]:
     _actualizar_headers(); url = f"{BASE_URL}/me/todo/lists/{lista_id}"; response: Optional[requests.Response] = None
     try:
         response = requests.delete(url, headers=HEADERS); response.raise_for_status()
         logging.info(f"Lista To Do ID '{lista_id}' eliminada."); return {"status": "Eliminado", "code": response.status_code} # logging.info
     except requests.exceptions.RequestException as e: logging.error(f"❌ Error eliminar lista todo {lista_id}: {e}"); raise Exception(f"Error eliminar lista todo: {e}")


def listar_tareas_todo(lista_id: str) -> Dict[str, Any]:
    """Lista las tareas de una lista de tareas de To Do específica, manejando paginación."""
    _actualizar_headers()
    url: Optional[str] = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks"
    all_tasks: List[Dict[str, Any]] = []
    response: Optional[requests.Response] = None
    try:
        while url:
            logging.info(f"Obteniendo página de tareas todo desde: {url}") # Corregido a logging.
            response = requests.get(url, headers=HEADERS); response.raise_for_status()
            data: Dict[str, Any] = response.json()
            tasks_in_page = data.get('value', [])
            if isinstance(tasks_in_page, list): all_tasks.extend(tasks_in_page)
            else: logging.warning(f"Se esperaba lista en 'value' (tareas todo): {type(tasks_in_page)}") # Corregido a logging.
            url = data.get('@odata.nextLink')
            if url: _actualizar_headers()
        logging.info(f"Listadas tareas lista To Do ID '{lista_id}'. Total: {len(all_tasks)}") # Corregido a logging.
        return {'value': all_tasks}
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error listar tareas todo {lista_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error listar tareas todo: {e}") # logging.error
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (listar tareas todo): {e}. Respuesta: {response_text}"); raise Exception(f"Error al decodificar JSON (listar tareas todo): {e}") # logging.error


def crear_tarea_todo(lista_id: str, titulo_tarea: str, detalles: Optional[Any] = None) -> Dict[str, Any]:
    _actualizar_headers(); url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks"; body: Dict[str, Any] = {"title": titulo_tarea}
    if detalles is not None:
         if isinstance(detalles, dict) and 'content' in detalles and isinstance(detalles['content'], str): body['body'] = detalles
         elif isinstance(detalles, str): body['body'] = {"content": detalles, "contentType": "text"}
         else: logging.warning(f"Formato inesperado para 'detalles' en crear_tarea_todo: {detalles}") # logging.warning
    response: Optional[requests.Response] = None
    try:
        response = requests.post(url, headers=HEADERS, json=body); response.raise_for_status(); data = response.json()
        task_id = data.get('id'); logging.info(f"Tarea ToDo '{titulo_tarea}' creada ID: {task_id}."); return data # logging.info
    except requests.exceptions.RequestException as e: logging.error(f"❌ Error crear tarea todo: {e}"); raise Exception(f"Error crear tarea todo: {e}")
    except json.JSONDecodeError as e: logging.error(f"❌ Error JSON (crear tarea todo): {e}"); raise Exception(f"Error JSON (crear tarea todo): {e}")


def actualizar_tarea_todo(lista_id: str, tarea_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]:
     _actualizar_headers(); url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks/{tarea_id}"; response: Optional[requests.Response] = None
     try:
         response = requests.patch(url, headers=HEADERS, json=nuevos_valores); response.raise_for_status(); data = response.json()
         logging.info(f"Tarea ToDo ID '{tarea_id}' actualizada."); return data # logging.info
     except requests.exceptions.RequestException as e: logging.error(f"❌ Error actualizar tarea todo {tarea_id}: {e}"); raise Exception(f"Error actualizar tarea todo: {e}")
     except json.JSONDecodeError as e: logging.error(f"❌ Error JSON (actualizar tarea todo): {e}"); raise Exception(f"Error JSON (actualizar tarea todo): {e}")

def eliminar_tarea_todo(lista_id: str, tarea_id: str) -> Dict[str, Any]:
     _actualizar_headers(); url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks/{tarea_id}"; response: Optional[requests.Response] = None
     try:
         response = requests.delete(url, headers=HEADERS); response.raise_for_status()
         logging.info(f"Tarea ToDo ID '{tarea_id}' eliminada."); return {"status": "Eliminado", "code": response.status_code} # logging.info
     except requests.exceptions.RequestException as e: logging.error(f"❌ Error eliminar tarea todo {tarea_id}: {e}"); raise Exception(f"Error eliminar tarea todo: {e}")


def completar_tarea_todo(lista_id: str, tarea_id: str) -> dict:
    logging.info(f"Completando tarea ToDo ID '{tarea_id}'.") # logging.info
    payload = {"status": "completed"}
    return actualizar_tarea_todo(lista_id, tarea_id, payload)

# --- FIN: Funciones Auxiliares ---
