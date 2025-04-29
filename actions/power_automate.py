import logging
import os
import requests
from auth import obtener_token  # Importante: Importar la función obtener_token
from typing import Dict, List, Optional, Union

# Configuración básica de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Variables de entorno (¡CRUCIALES!)
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
GRAPH_SCOPE = os.getenv('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')  # Valor por defecto
AZURE_SUBSCRIPTION_ID = os.getenv('AZURE_SUBSCRIPTION_ID')  # Nuevo: ID de la suscripción de Azure
AZURE_RESOURCE_GROUP = os.getenv('AZURE_RESOURCE_GROUP')  # Nuevo: Nombre del grupo de recursos de Azure
AZURE_LOCATION = os.getenv('AZURE_LOCATION') # Nueva variable para la ubicación de los recursos de Azure

# Verificar variables de entorno (¡CRUCIALES!)
if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE, AZURE_SUBSCRIPTION_ID, AZURE_RESOURCE_GROUP]):
    logging.error("❌ Faltan variables de entorno (CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE, AZURE_SUBSCRIPTION_ID, AZURE_RESOURCE_GROUP). La función no puede funcionar.")
    raise Exception("Faltan variables de entorno.")

BASE_URL = "https://management.azure.com"
RESOURCE = "https://service.flow.microsoft.com/"
API_VERSION = "2016-11-01"  # Esta versión es bastante antigua, ¡verificar!
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



# ---- POWER AUTOMATE (Flows) ----

def listar_flows(suscripcion_id: str, grupo_recurso: str) -> dict:
    """Lista los flujos de Power Automate en un grupo de recursos."""
    _actualizar_headers()
    url = f"{BASE_URL}/subscriptions/{suscripcion_id}/resourceGroups/{grupo_recurso}/providers/Microsoft.Logic/workflows?api-version={API_VERSION}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Listados flows en el grupo de recursos '{grupo_recurso}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar flows: {e}")
        raise Exception(f"Error al listar flows: {e}")



def obtener_flow(suscripcion_id: str, grupo_recurso: str, nombre_flow: str) -> dict:
    """Obtiene un flujo de Power Automate específico."""
    _actualizar_headers()
    url = f"{BASE_URL}/subscriptions/{suscripcion_id}/resourceGroups/{grupo_recurso}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={API_VERSION}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Obtenido flow '{nombre_flow}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener el flow '{nombre_flow}': {e}")
        raise Exception(f"Error al obtener el flow '{nombre_flow}': {e}")



def crear_flow(suscripcion_id: str, grupo_recurso: str, nombre_flow: str, definicion_flow: dict, ubicacion:str) -> dict:
    """Crea un nuevo flujo de Power Automate."""
    _actualizar_headers()
    url = f"{BASE_URL}/subscriptions/{suscripcion_id}/resourceGroups/{grupo_recurso}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={API_VERSION}"
    body = {
        "location": ubicacion,  # Debes proporcionar la ubicación del recurso
        "properties": {
            "definition": definicion_flow  # La definición del flujo (JSON)
        }
    }
    try:
        response = requests.put(url, headers=HEADERS, json=body)  # Usar PUT para crear
        response.raise_for_status()
        data = response.json()
        logging.info(f"Flujo '{nombre_flow}' creado en el grupo de recursos '{grupo_recurso}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al crear el flujo '{nombre_flow}': {e}")
        raise Exception(f"Error al crear el flujo '{nombre_flow}': {e}")



def actualizar_flow(suscripcion_id: str, grupo_recurso: str, nombre_flow: str, definicion_flow: dict) -> dict:
    """Actualiza un flujo de Power Automate existente."""
    _actualizar_headers()
    url = f"{BASE_URL}/subscriptions/{suscripcion_id}/resourceGroups/{grupo_recurso}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={API_VERSION}"
    body = {
        "properties": {
            "definition": definicion_flow
        }
    }
    try:
        response = requests.put(url, headers=HEADERS, json=body)  # Usar PUT para actualizar
        response.raise_for_status()
        data = response.json()
        logging.info(f"Flujo '{nombre_flow}' actualizado en el grupo de recursos '{grupo_recurso}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al actualizar el flujo '{nombre_flow}': {e}")
        raise Exception(f"Error al actualizar el flujo '{nombre_flow}': {e}")



def eliminar_flow(suscripcion_id: str, grupo_recurso: str, nombre_flow: str) -> dict:
    """Elimina un flujo de Power Automate."""
    _actualizar_headers()
    url = f"{BASE_URL}/subscriptions/{suscripcion_id}/resourceGroups/{grupo_recurso}/providers/Microsoft.Logic/workflows/{nombre_flow}?api-version={API_VERSION}"
    try:
        response = requests.delete(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Flujo '{nombre_flow}' eliminado del grupo de recursos '{grupo_recurso}'.")
        return {"status": "Eliminado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al eliminar el flujo '{nombre_flow}': {e}")
        raise Exception(f"Error al eliminar el flujo '{nombre_flow}': {e}")



def ejecutar_flow(flow_url: str, parametros: Optional[dict] = None) -> dict:
    """Ejecuta un flujo de Power Automate a través de su URL de desencadenador."""
    _actualizar_headers()  # En este caso, es posible que no sea necesario, dependiendo de cómo esté autenticado el trigger del Flow
    headers = {'Content-Type': 'application/json'}
    try:
        response = requests.post(flow_url, headers=headers, json=parametros if parametros else {})
        response.raise_for_status()
        logging.info(f"Flujo en la URL '{flow_url}' ejecutado.")
        return {"status": "Ejecutado", "code": response.status_code, "response": response.json()} #returns the json
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al ejecutar el flujo en la URL '{flow_url}': {e}")
        raise Exception(f"Error al ejecutar el flujo en la URL '{flow_url}': {e}")



def obtener_estado_ejecucion_flow(run_id: str, suscripcion_id: str, grupo_recurso: str, nombre_flow:str) -> dict:
    """
    Obtiene el estado de una ejecución específica de un flujo de Power Automate.

    Args:
        run_id: El ID de la ejecución del flujo.
        suscripcion_id: El ID de la suscripción de Azure.
        grupo_recurso: El nombre del grupo de recursos de Azure.
        nombre_flow: El nombre del flow.

    Returns:
        Un diccionario con el estado de la ejecución del flujo.
    """
    _actualizar_headers()
    url = f"{BASE_URL}/subscriptions/{suscripcion_id}/resourceGroups/{grupo_recurso}/providers/Microsoft.Logic/workflows/{nombre_flow}/runs/{run_id}?api-version={API_VERSION}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Obtenido estado de ejecución '{run_id}' del flujo '{nombre_flow}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener el estado de ejecución '{run_id}' del flujo '{nombre_flow}': {e}")
        raise Exception(f"Error al obtener el estado de ejecución '{run_id}' del flujo '{nombre_flow}': {e}")
