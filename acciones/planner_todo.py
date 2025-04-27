import logging
import os
import requests
from auth import obtener_token  # Importante: Importar la función obtener_token
from typing import Dict, List, Optional, Union
from datetime import datetime

# Configuración básica de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Variables de entorno (¡CRUCIALES!)
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
GRAPH_SCOPE = os.getenv('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')  # Valor por defecto

# Verificar variables de entorno (¡CRUCIAL!)
if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE]):
    logging.error("❌ Faltan variables de entorno (CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE). La función no puede funcionar.")
    raise Exception("Faltan variables de entorno.")

BASE_URL = "https://graph.microsoft.com/v1.0"
HEADERS = {
    'Authorization': None,  # Inicialmente None, se actualiza con cada request
    'Content-Type': 'application/json'
}


# Función para obtener el token y actualizar los HEADERS
def _actualizar_headers() -> None:
    """Obtiene un nuevo token de acceso y actualiza el diccionario HEADERS."""
    try:
        HEADERS['Authorization'] = f'Bearer {obtener_token()}'
    except Exception as e:  # Captura la excepción de obtener_token
        logging.error(f"❌ Error al obtener el token: {e}")
        raise Exception(f"Error al obtener el token: {e}")



# ---- PLANNER ----

def listar_planes(grupo_id: str) -> dict:
    """Lista los planes de Planner en un grupo específico."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{grupo_id}/planner/plans"  # Usa el ID del grupo
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Listados planes del grupo '{grupo_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar planes del grupo '{grupo_id}': {e}")
        raise Exception(f"Error al listar planes del grupo '{grupo_id}': {e}")



def obtener_plan(plan_id: str) -> dict:
    """Obtiene un plan de Planner específico."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/plans/{plan_id}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener plan con ID {plan_id}: {e}")
        raise Exception(f"Error al obtener plan con ID {plan_id}: {e}")


def crear_plan(nombre_plan: str, grupo_id: str) -> dict:
    """Crea un nuevo plan de Planner en un grupo."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/plans"
    body = {
        "owner@odata.bind": f"https://graph.microsoft.com/v1.0/groups/{grupo_id}",  # Referencia al grupo
        "title": nombre_plan
    }
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        plan_id = data.get('id')
        logging.info(f"Plan '{nombre_plan}' creado en el grupo '{grupo_id}' con ID: {plan_id}.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear plan '{nombre_plan}' en el grupo '{grupo_id}': {e}")
        raise Exception(f"Error al crear plan '{nombre_plan}' en el grupo '{grupo_id}': {e}")



def actualizar_plan(plan_id: str, nuevos_valores: Dict) -> dict:
    """Actualiza un plan de Planner existente."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/plans/{plan_id}"
    # La API de Planner requiere el ETag para actualizaciones.  Asegúrate de incluirlo en nuevos_valores si es necesario.
    try:
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status()
        logging.info(f"Plan con ID '{plan_id}' actualizado. Nuevos Valores: {nuevos_valores}")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al actualizar el plan con ID '{plan_id}': {e}")
        raise Exception(f"Error al actualizar el plan con ID '{plan_id}': {e}")


def eliminar_plan(plan_id: str) -> dict:
    """Elimina un plan de Planner."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/plans/{plan_id}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Plan con ID '{plan_id}' eliminado.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar el plan con ID '{plan_id}': {e}")
        raise Exception(f"Error al eliminar el plan con ID '{plan_id}': {e}")



def listar_tareas_planner(plan_id: str) -> dict:
    """Lista las tareas de un plan de Planner específico, manejando paginación."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/plans/{plan_id}/tasks"
    try:
        all_tasks = []
        while url:
            response = requests.get(url, headers=HEADERS)
            response.raise_for_status()
            data = response.json()
            all_tasks.extend(data.get('value', []))
            url = data.get('@odata.nextLink')
            if url:
                _actualizar_headers()
        logging.info(f"Listadas tareas del plan con ID '{plan_id}'. Total tareas: {len(all_tasks)}")
        return {'value': all_tasks}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar las tareas del plan con ID '{plan_id}': {e}")
        raise Exception(f"Error al listar las tareas del plan con ID '{plan_id}': {e}")


def crear_tarea_planner(plan_id: str, titulo_tarea: str, bucket_id: Optional[str] = None, detalles: Optional[dict] = None) -> dict:
    """Crea una nueva tarea de Planner en un plan."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/tasks"
    body = {
        "planId": plan_id,
        "title": titulo_tarea,
    }
    if bucket_id:
        body["bucketId"] = bucket_id
    if detalles:
        body["details"] = detalles

    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        task_id = data.get('id')
        logging.info(f"Tarea '{titulo_tarea}' creada en el plan '{plan_id}' con ID: {task_id}.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear la tarea '{titulo_tarea}' en el plan con ID '{plan_id}': {e}")
        raise Exception(f"Error al crear la tarea '{titulo_tarea}' en el plan con ID '{plan_id}': {e}")



def actualizar_tarea_planner(tarea_id: str, nuevos_valores: dict) -> dict:
    """Actualiza una tarea de Planner existente."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/tasks/{tarea_id}"
    # La API de Planner requiere el ETag para actualizaciones.  Asegúrate de incluirlo en nuevos_valores si es necesario.
    try:
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status()
        logging.info(f"Tarea con ID '{tarea_id}' actualizada. Nuevos Valores: {nuevos_valores}")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al actualizar la tarea con ID '{tarea_id}': {e}")
        raise Exception(f"Error al actualizar la tarea con ID '{tarea_id}': {e}")



def eliminar_tarea_planner(tarea_id: str) -> dict:
    """Elimina una tarea de Planner."""
    _actualizar_headers()
    url = f"{BASE_URL}/planner/tasks/{tarea_id}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Tarea con ID '{tarea_id}' eliminada.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar la tarea con ID '{tarea_id}': {e}")
        raise Exception(f"Error al eliminar la tarea con ID '{tarea_id}': {e}")


# ---- TO DO ----

def listar_listas_todo() -> dict:
    """Lista las listas de tareas de To Do del usuario."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info("Listadas las listas de To Do.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar las listas de To Do: {e}")
        raise Exception(f"Error al listar las listas de To Do: {e}")



def crear_lista_todo(nombre_lista: str) -> dict:
    """Crea una nueva lista de tareas de To Do."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists"
    body = {"displayName": nombre_lista}
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        list_id = data.get('id')
        logging.info(f"Lista de To Do '{nombre_lista}' creada con ID: {list_id}.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear la lista de To Do '{nombre_lista}': {e}")
        raise Exception(f"Error al crear la lista de To Do '{nombre_lista}': {e}")



def actualizar_lista_todo(lista_id: str, nuevos_valores: dict) -> dict:
    """Actualiza una lista de tareas de To Do existente."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists/{lista_id}"
    try:
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status()
        logging.info(f"Lista de To Do con ID '{lista_id}' actualizada. Nuevos valores: {nuevos_valores}")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al actualizar la lista de To Do con ID '{lista_id}': {e}")
        raise Exception(f"Error al actualizar la lista de To Do con ID '{lista_id}': {e}")



def eliminar_lista_todo(lista_id: str) -> dict:
    """Elimina una lista de tareas de To Do."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists/{lista_id}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Lista de To Do con ID '{lista_id}' eliminada.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar la lista de To Do con ID '{lista_id}': {e}")
        raise Exception(f"Error al eliminar la lista de To Do con ID '{lista_id}': {e}")



def listar_tareas_todo(lista_id: str) -> dict:
    """Lista las tareas de una lista de tareas de To Do específica, manejando paginación."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks"
    try:
        all_tasks = []
        while url:
            response = requests.get(url, headers=HEADERS)
            response.raise_for_status()
            data = response.json()
            all_tasks.extend(data.get('value', []))
            url = data.get('@odata.nextLink')
            if url:
                _actualizar_headers()
        logging.info(f"Listadas tareas de la lista de To Do con ID '{lista_id}'. Total tareas: {len(all_tasks)}")
        return {'value': all_tasks}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar las tareas de la lista de To Do con ID '{lista_id}': {e}")
        raise Exception(f"Error al listar las tareas de la lista de To Do con ID '{lista_id}': {e}")



def crear_tarea_todo(lista_id: str, titulo_tarea: str, detalles: Optional[dict] = None) -> dict:
    """Crea una nueva tarea en una lista de tareas de To Do."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks"
    body = {
        "title": titulo_tarea,
    }
    if detalles:
        body["body"] = detalles
    try:
        response = requests.post(url, headers=HEADERS, json=body)
        response.raise_for_status()
        data = response.json()
        task_id = data.get('id')
        logging.info(f"Tarea '{titulo_tarea}' creada en la lista de To Do '{lista_id}' con ID: {task_id}.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear la tarea '{titulo_tarea}' en la lista de To Do con ID '{lista_id}': {e}")
        raise Exception(f"Error al crear la tarea '{titulo_tarea}' en la lista de To Do con ID '{lista_id}': {e}")



def actualizar_tarea_todo(lista_id: str, tarea_id: str, nuevos_valores: dict) -> dict:
    """Actualiza una tarea de To Do existente."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks/{tarea_id}"
    try:
        response = requests.patch(url, headers=HEADERS, json=nuevos_valores)
        response.raise_for_status()
        logging.info(f"Tarea con ID '{tarea_id}' actualizada en la lista de To Do con ID '{lista_id}'.  Nuevos Valores: {nuevos_valores}")
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al actualizar la tarea con ID '{tarea_id}' en la lista de To Do con ID '{lista_id}': {e}")
        raise Exception(f"Error al actualizar la tarea con ID '{lista_id}' en la lista de To Do con ID '{lista_id}': {e}")



def eliminar_tarea_todo(lista_id: str, tarea_id: str) -> dict:
    """Elimina una tarea de una lista de tareas de To Do."""
    _actualizar_headers()
    url = f"{BASE_URL}/me/todo/lists/{lista_id}/tasks/{tarea_id}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Tarea con ID '{tarea_id}' eliminada de la lista de To Do con ID '{lista_id}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar la tarea con ID '{tarea_id}' de la lista de To Do con ID '{lista_id}': {e}")
        raise Exception(f"Error al eliminar la tarea con ID '{tarea_id}' de la lista de To Do con ID '{lista_id}': {e}")



def completar_tarea_todo(lista_id: str, tarea_id: str) -> dict:
    """Completa una tarea de To Do."""
    return actualizar_tarea_todo(lista_id, tarea_id, {"status": "completed"})
