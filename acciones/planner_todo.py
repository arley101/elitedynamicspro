import logging
import os
import requests
# CORRECCION: Asegurarse que la importación de auth es correcta si auth.py está en la raíz
# Si auth.py está en la raíz y acciones es una subcarpeta, esta importación podría necesitar ajuste
# dependiendo de cómo se ejecute el código. Para MyPy y ejecución simple, podría funcionar.
# Si falla en runtime, podrías necesitar ajustar PYTHONPATH o usar importación relativa/absoluta.
from auth import obtener_token
# CORRECCION: Añadir Any para tipos de diccionarios
from typing import Dict, List, Optional, Union, Any
from datetime import datetime
import json # Importar json para los except blocks

# Configuración básica de logging
# Considera usar el logger de Azure Functions si es aplicable
# logger = logging.getLogger("azure.functions")
# logger.setLevel(logging.INFO)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- INICIO: Configuración Redundante - Considerar eliminar ---
# Esta sección es redundante si __init__.py o auth.py ya cargan y validan esto.
# Si se elimina, HEADERS y _actualizar_headers también deben importarse o pasarse.
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
GRAPH_SCOPE = os.getenv('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')

if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE]):
    logging.error("❌ Faltan variables de entorno (CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE) en planner_todo.")
    # No debería lanzar Exception aquí, quizás solo loguear y permitir que falle después si se usa
    # raise Exception("Faltan variables de entorno.")

BASE_URL = "https://graph.microsoft.com/v1.0"
HEADERS: Dict[str, Optional[str]] = { # Tipado explícito
    'Authorization': None,
    'Content-Type': 'application/json'
}

def _actualizar_headers() -> None:
    """Obtiene un nuevo token de acceso y actualiza el diccionario HEADERS local."""
    # Advertencia: Esta función actualiza los HEADERS definidos LOCALMENTE en este archivo.
    try:
        # Llama a la función importada de auth.py
        token = obtener_token() # Asume flujo 'aplicacion' por defecto
        HEADERS['Authorization'] = f'Bearer {token}'
        logging.info("Headers actualizados en planner_todo.")
    except Exception as e:
        logging.error(f"❌ Error al obtener el token en planner_todo: {e}")
        raise Exception(f"Error al obtener el token en planner_todo: {e}")
# --- FIN: Configuración Redundante ---


# ---- PLANNER ----

def listar_planes(grupo_id: str) -> Dict[str, Any]: # Tipo de retorno
    """Lista los planes de Planner en un grupo específico."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{grupo_id}/planner/plans"
    response: Optional[requests.Response] = None
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data: Dict[str, Any] = response.json() # Tipado
        logging.info(f"Listados planes del grupo '{grupo_id}'.")
        return data
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error listar planes: {e}. Detalles: {error_details}"); raise Exception(f"Error listar planes: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (listar planes): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar planes): {e}")


def obtener_plan(plan_id: str) -> Dict[str, Any]: # Tipo de retorno
    """Obtiene un plan de Planner específico."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/plans/{plan_id}"
    response: Optional[requests.Response] = None
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        return response.json()
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error obtener plan {plan_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error obtener plan {plan_id}: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (obtener plan): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (obtener plan): {e}")

def crear_plan(nombre_plan: str, grupo_id: str) -> Dict[str, Any]: # Tipo de retorno
    """Crea un nuevo plan de Planner en un grupo."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/plans"
    body = {"owner": grupo_id, "title": nombre_plan} # Simplificado, Graph ahora permite owner ID directamente
    response: Optional[requests.Response] = None
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data: Dict[str, Any] = response.json()
        plan_id = data.get('id'); logging.info(f"Plan '{nombre_plan}' creado ID: {plan_id}."); return data
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error crear plan: {e}. Detalles: {error_details}"); raise Exception(f"Error crear plan: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (crear plan): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (crear plan): {e}")


def actualizar_plan(plan_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]: # Tipo de retorno
    """Actualiza un plan de Planner existente."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/plans/{plan_id}"
    # Nota: Planner API a menudo requiere If-Match header con ETag para PATCH.
    # Si falla con 412 Precondition Failed, necesitas obtener el ETag del plan y añadirlo a HEADERS.
    # Ejemplo: HEADERS['If-Match'] = etag_obtenido
    response: Optional[requests.Response] = None
    try:
        # Añadir ETag si está en nuevos_valores (buena práctica pasarlo)
        etag = nuevos_valores.pop('@odata.etag', None) # Quitar etag del body si viene
        current_headers = HEADERS.copy() # Copiar para no modificar global permanentemente
        if etag:
            current_headers['If-Match'] = etag
            logging.info(f"Usando ETag: {etag} para actualizar plan {plan_id}")

        response = requests.patch(url, headers=current_headers, json=nuevos_valores)
        response.raise_for_status()
        logging.info(f"Plan ID '{plan_id}' actualizado.")
        # PATCH exitoso usualmente devuelve 200 OK con el objeto actualizado o 204 No Content.
        # Si es 204, response.json() fallará.
        if response.status_code == 204:
            # Si no hay contenido, re-obtener el plan para devolver el estado actualizado
             logging.warning(f"Actualizar plan {plan_id} devolvió 204 No Content. Re-obteniendo plan.")
             return obtener_plan(plan_id)
        else:
             return response.json()
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error actualizar plan {plan_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error actualizar plan {plan_id}: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (actualizar plan): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (actualizar plan): {e}")


def eliminar_plan(plan_id: str) -> Dict[str, Any]: # Tipo de retorno
    """Elimina un plan de Planner."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/plans/{plan_id}"
    # Nota: Planner API requiere If-Match header con ETag para DELETE.
    # Necesitarías obtener el ETag primero. Aquí asumimos que no se necesita o falla.
    response: Optional[requests.Response] = None
    try:
        # TODO: Considerar obtener ETag y añadirlo a Headers con If-Match
        logging.warning(f"Intentando eliminar plan {plan_id} sin ETag (If-Match). Podría fallar si es requerido.")
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status() # Espera 204 No Content
        logging.info(f"Plan ID '{plan_id}' eliminado.")
        return {"status": "Eliminado", "code": response.status_code}
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error eliminar plan {plan_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error eliminar plan {plan_id}: {e}")


def listar_tareas_planner(plan_id: str) -> Dict[str, Any]: # Tipo de retorno
    """Lista las tareas de un plan de Planner específico, manejando paginación."""
    _actualizar_headers()
    # CORRECCIÓN: Inicializar url como Optional[str]
    url: Optional[str] = f"{BASE_URL}/planner/plans/{plan_id}/tasks"
    all_tasks: List[Dict[str, Any]] = [] # CORRECCIÓN: Tipado explícito
    response: Optional[requests.Response] = None
    try:
        while url:
            logger.info(f"Obteniendo página de tareas planner desde: {url}")
            response = requests.get(url, headers=HEADERS)
            response.raise_for_status()
            data: Dict[str, Any] = response.json()
            tasks_in_page = data.get('value', [])
            if isinstance(tasks_in_page, list): all_tasks.extend(tasks_in_page)
            else: logger.warning(f"Se esperaba lista en 'value' (tareas planner): {type(tasks_in_page)}")

            # Línea ~154 (equivalente): Asigna Optional[str] a url (ahora tipado como Optional[str])
            url = data.get('@odata.nextLink')
            if url:
                _actualizar_headers() # Actualiza headers para la siguiente página

        logging.info(f"Listadas tareas del plan ID '{plan_id}'. Total: {len(all_tasks)}")
        return {'value': all_tasks}
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error listar tareas planner {plan_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error listar tareas planner: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (listar tareas planner): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar tareas planner): {e}")


def crear_tarea_planner(plan_id: str, titulo_tarea: str, bucket_id: Optional[str] = None, detalles: Optional[Dict[str, Any]] = None) -> Dict[str, Any]: # Tipo de retorno
    """Crea una nueva tarea de Planner en un plan."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/tasks"
    body: Dict[str, Any] = { # Tipado explícito
        "planId": plan_id,
        "title": titulo_tarea,
    }
    if bucket_id: body["bucketId"] = bucket_id
    # 'detalles' usualmente se actualiza con PATCH después de crear la tarea,
    # pero si lo pasas aquí, debe ser el objeto details completo.
    # MyPy podría quejarse de la estructura de 'detalles' si no la defines mejor.
    if detalles and isinstance(detalles, dict): body.update(detalles) # Simple merge, might need refinement

    response: Optional[requests.Response] = None
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data: Dict[str, Any] = response.json()
        task_id = data.get('id'); logging.info(f"Tarea '{titulo_tarea}' creada ID: {task_id}."); return data
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error crear tarea planner: {e}. Detalles: {error_details}"); raise Exception(f"Error crear tarea planner: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (crear tarea planner): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (crear tarea planner): {e}")


def actualizar_tarea_planner(tarea_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]: # Tipo de retorno
    """Actualiza una tarea de Planner existente."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/tasks/{tarea_id}"
    # Nota: Planner API requiere If-Match header con ETag para PATCH.
    response: Optional[requests.Response] = None
    try:
        etag = nuevos_valores.pop('@odata.etag', None)
        current_headers = HEADERS.copy()
        if etag: current_headers['If-Match'] = etag; logging.info(f"Usando ETag para actualizar tarea {tarea_id}")

        response = requests.patch(url, headers=current_headers, json=nuevos_valores)
        response.raise_for_status()
        logging.info(f"Tarea ID '{tarea_id}' actualizada.")
        if response.status_code == 204:
             # Re-obtener si no hay contenido
             logging.warning(f"Actualizar tarea {tarea_id} devolvió 204 No Content. Re-obteniendo tarea.")
             # Necesitaríamos una función obtener_tarea_planner(tarea_id)
             # return obtener_tarea_planner(tarea_id)
             return {"status": "Actualizado (No Content)", "id": tarea_id} # Placeholder
        else:
             return response.json()
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error actualizar tarea planner {tarea_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error actualizar tarea planner: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (actualizar tarea planner): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (actualizar tarea planner): {e}")


def eliminar_tarea_planner(tarea_id: str) -> Dict[str, Any]: # Tipo de retorno
    """Elimina una tarea de Planner."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/tasks/{tarea_id}"
    # Nota: Planner API requiere If-Match header con ETag para DELETE.
    response: Optional[requests.Response] = None
    try:
        logging.warning(f"Intentando eliminar tarea planner {tarea_id} sin ETag.")
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status() # Espera 204
        logging.info(f"Tarea ID '{tarea_id}' eliminada.")
        return {"status": "Eliminado", "code": response.status_code}
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error eliminar tarea planner {tarea_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error eliminar tarea planner: {e}")


# ---- TO DO ----

def listar_listas_todo() -> Dict[str, Any]: # Tipo de retorno
    """Lista las listas de tareas de To Do del usuario."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists"
    response: Optional[requests.Response] = None
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status(); data = response.json(); logging.info("Listadas listas To Do."); return data
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error listar listas todo: {e}. Detalles: {error_details}"); raise Exception(f"Error listar listas todo: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (listar listas todo): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar listas todo): {e}")


def crear_lista_todo(nombre_lista: str) -> Dict[str, Any]: # Tipo de retorno
    """Crea una nueva lista de tareas de To Do."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists"
    body = {"displayName": nombre_lista}
    response: Optional[requests.Response] = None
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status(); data = response.json(); list_id = data.get('id'); logging.info(f"Lista To Do '{nombre_lista}' creada ID: {list_id}."); return data
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error crear lista todo: {e}. Detalles: {error_details}"); raise Exception(f"Error crear lista todo: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (crear lista todo): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (crear lista todo): {e}")


def actualizar_lista_todo(lista_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]: # Tipo de retorno
    """Actualiza una lista de tareas de To Do existente."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists/{lista_id}"
    response: Optional[requests.Response] = None
    try:
        # Asume que no se necesita ETag para listas ToDo (verificar docs si falla)
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status(); data = response.json(); logging.info(f"Lista To Do ID '{lista_id}' actualizada."); return data
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error actualizar lista todo {lista_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error actualizar lista todo: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (actualizar lista todo): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (actualizar lista todo): {e}")


def eliminar_lista_todo(lista_id: str) -> Dict[str, Any]: # Tipo de retorno
    """Elimina una lista de tareas de To Do."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists/{lista_id}"
    response: Optional[requests.Response] = None
    try:
        # Asume que no se necesita ETag (verificar docs si falla)
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status(); logging.info(f"Lista To Do ID '{lista_id}' eliminada."); return {"status": "Eliminado", "code": response.status_code}
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error eliminar lista todo {lista_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error eliminar lista todo: {e}")


def listar_tareas_todo(lista_id: str) -> Dict[str, Any]: # Tipo de retorno
    """Lista las tareas de una lista de tareas de To Do específica, manejando paginación."""
    _actualizar_headers()
    # CORRECCIÓN: Inicializar url como Optional[str]
    url: Optional[str] = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks"
    all_tasks: List[Dict[str, Any]] = [] # Tipado explícito
    response: Optional[requests.Response] = None
    try:
        while url:
            logger.info(f"Obteniendo página de tareas todo desde: {url}")
            response = requests.get(url, headers=HEADERS)
            response.raise_for_status()
            data: Dict[str, Any] = response.json()
            tasks_in_page = data.get('value', [])
            if isinstance(tasks_in_page, list): all_tasks.extend(tasks_in_page)
            else: logger.warning(f"Se esperaba lista en 'value' (tareas todo): {type(tasks_in_page)}")

            # Línea ~295 (equivalente): Asigna Optional[str] a url
            url = data.get('@odata.nextLink')
            if url: _actualizar_headers() # Actualiza para siguiente página

        logging.info(f"Listadas tareas lista To Do ID '{lista_id}'. Total: {len(all_tasks)}")
        return {'value': all_tasks}
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error listar tareas todo {lista_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error listar tareas todo: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (listar tareas todo): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (listar tareas todo): {e}")


def crear_tarea_todo(lista_id: str, titulo_tarea: str, detalles: Optional[Dict[str, Any]] = None) -> Dict[str, Any]: # Tipo de retorno
    """Crea una nueva tarea en una lista de tareas de To Do."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks"
    body: Dict[str, Any] = {"title": titulo_tarea} # Tipado
    # 'detalles' en ToDo usualmente se refiere al cuerpo/descripción, esperaríamos un string
    # o un objeto específico si queremos setear fechas, etc. El API es diferente a Planner.
    # Asumiré que 'detalles' aquí podría ser el cuerpo como texto.
    if detalles is not None:
         if isinstance(detalles, dict) and 'content' in detalles and isinstance(detalles['content'], str):
             # Si pasan un dict como {"content": "descripcion", "contentType": "text"}
             body['body'] = detalles
         elif isinstance(detalles, str):
             # Si pasan solo el string del cuerpo
             body['body'] = {"content": detalles, "contentType": "text"}
         else:
             logger.warning(f"Formato inesperado para 'detalles' en crear_tarea_todo: {detalles}")

    response: Optional[requests.Response] = None
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status(); data = response.json(); task_id = data.get('id'); logging.info(f"Tarea ToDo '{titulo_tarea}' creada ID: {task_id}."); return data
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error crear tarea todo: {e}. Detalles: {error_details}"); raise Exception(f"Error crear tarea todo: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (crear tarea todo): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (crear tarea todo): {e}")


def actualizar_tarea_todo(lista_id: str, tarea_id: str, nuevos_valores: Dict[str, Any]) -> Dict[str, Any]: # Tipo de retorno
    """Actualiza una tarea de To Do existente."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks/{tarea_id}"
    response: Optional[requests.Response] = None
    try:
        # Asume que no necesita ETag (verificar docs si falla)
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status(); data = response.json(); logging.info(f"Tarea ToDo ID '{tarea_id}' actualizada."); return data
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error actualizar tarea todo {tarea_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error actualizar tarea todo: {e}")
    except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (actualizar tarea todo): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (actualizar tarea todo): {e}")


def eliminar_tarea_todo(lista_id: str, tarea_id: str) -> Dict[str, Any]: # Tipo de retorno
    """Elimina una tarea de una lista de tareas de To Do."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks/{tarea_id}"
    response: Optional[requests.Response] = None
    try:
        # Asume que no necesita ETag (verificar docs si falla)
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status(); logging.info(f"Tarea ToDo ID '{tarea_id}' eliminada."); return {"status": "Eliminado", "code": response.status_code}
    # ... (except blocks) ...
    except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error eliminar tarea todo {tarea_id}: {e}. Detalles: {error_details}"); raise Exception(f"Error eliminar tarea todo: {e}")


def completar_tarea_todo(lista_id: str, tarea_id: str) -> dict:
    """Completa una tarea de To Do."""
    # Completar es actualizar el estado
    logging.info(f"Intentando completar tarea ToDo ID '{tarea_id}' en lista '{lista_id}'.")
    # El cuerpo para completar puede variar, a veces es {"status": "completed"}
    # A veces requiere actualizar 'completedDateTime'. Consultar docs de ToDo Task API.
    # Por simplicidad, intentamos con status.
    payload = {"status": "completed"}
    return actualizar_tarea_todo(lista_id, tarea_id, payload)

# --- FIN: Funciones Auxiliares ---
