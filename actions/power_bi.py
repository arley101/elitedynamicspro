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

# Verificar variables de entorno (¡CRUCIALES!)
if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE]):
    logging.error("❌ Faltan variables de entorno (CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE). La función no puede funcionar.")
    raise Exception("Faltan variables de entorno.")

BASE_URL = "https://api.powerbi.com/v1.0/myorg"
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



# ---- POWER BI ----

def listar_workspaces(expand: Optional[List[str]] = None) -> dict:
    """Lista los workspaces de Power BI a los que tiene acceso el usuario."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups"
    if expand:
        url += f"?$expand={','.join(expand)}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Listados workspaces de Power BI.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar workspaces: {e}")
        raise Exception(f"Error al listar workspaces: {e}")



def obtener_workspace(workspace_id: str) -> dict:
    """Obtiene un workspace de Power BI específico."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{workspace_id}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener el workspace {workspace_id}: {e}")
        raise Exception(f"Error al obtener el workspace {workspace_id}: {e}")



def listar_dashboards(workspace_id: str) -> dict:
    """Lista los dashboards en un workspace de Power BI."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{workspace_id}/dashboards"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Listados dashboards del workspace '{workspace_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar dashboards del workspace '{workspace_id}': {e}")
        raise Exception(f"Error al listar dashboards del workspace '{workspace_id}': {e}")



def obtener_dashboard(workspace_id: str, dashboard_id: str) -> dict:
    """Obtiene un dashboard de Power BI específico."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{workspace_id}/dashboards/{dashboard_id}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener el dashboard {dashboard_id} del workspace {workspace_id}: {e}")
        raise Exception(f"Error al obtener el dashboard {dashboard_id} del workspace {workspace_id}: {e}")



def listar_reports(workspace_id: str) -> dict:
    """Lista los informes en un workspace de Power BI."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{workspace_id}/reports"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Listados informes del workspace '{workspace_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar informes del workspace '{workspace_id}': {e}")
        raise Exception(f"Error al listar informes del workspace '{workspace_id}': {e}")


def obtener_reporte(workspace_id: str, report_id: str) -> dict:
    """Obtiene un reporte de Power BI específico."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{workspace_id}/reports/{report_id}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener el reporte {report_id} del workspace {workspace_id}: {e}")
        raise Exception(f"Error al obtener el reporte {report_id} del workspace {workspace_id}: {e}")



def listar_datasets(workspace_id: str) -> dict:
    """Lista los datasets en un workspace de Power BI."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{workspace_id}/datasets"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Listados datasets del workspace '{workspace_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al listar datasets del workspace '{workspace_id}': {e}")
        raise Exception(f"Error al listar datasets del workspace '{workspace_id}': {e}")


def obtener_dataset(workspace_id: str, dataset_id: str) -> dict:
    """Obtiene un dataset de Power BI específico."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener el dataset {dataset_id} del workspace {workspace_id}: {e}")
        raise Exception(f"Error al obtener el dataset {dataset_id} del workspace {workspace_id}: {e}")



def refrescar_dataset(workspace_id: str, dataset_id: str) -> dict:
    """Refresca un dataset de Power BI."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}/refreshes"
    try:
        response = requests.post(url, headers=HEADERS)
        response.raise_for_status()
        logging.info(f"Refresco del dataset '{dataset_id}' en el workspace '{workspace_id}' iniciado.")
        return {"status": "Refresh iniciado", "code": response.status_code}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al refrescar el dataset '{dataset_id}' en el workspace '{workspace_id}': {e}")
        raise Exception(f"Error al refrescar el dataset '{dataset_id}' en el workspace '{workspace_id}': {e}")



def obtener_estado_refresco_dataset(workspace_id: str, dataset_id: str) -> dict:
    """Obtiene el estado de los refrescos de un dataset de Power BI."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{workspace_id}/datasets/{dataset_id}/refreshes"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        logging.info(f"Obtenido el estado de refresco del dataset '{dataset_id}' en el workspace '{workspace_id}'.")
        return data
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener el estado de refresco del dataset '{dataset_id}' en el workspace '{workspace_id}': {e}")
        raise Exception(f"Error al obtener el estado de refresco del dataset '{dataset_id}' en el workspace '{workspace_id}': {e}")



def obtener_embed_url(workspace_id: str, report_id: str) -> dict:
    """Obtiene la URL de inserción de un informe de Power BI."""
    _actualizar_headers()
    url = f"{BASE_URL}/groups/{workspace_id}/reports/{report_id}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json()
        embed_url = data.get("embedUrl")
        if embed_url:
            logging.info(f"Obtenida URL de inserción para el informe '{report_id}' del workspace '{workspace_id}': {embed_url}")
            return {"embedUrl": embed_url}
        else:
            logging.warning(f"⚠ No se encontró la URL de inserción para el informe '{report_id}' del workspace '{workspace_id}'.")
            return {"error": "No se encontró la URL de inserción"}
    except requests.exceptions.RequestException as e:
        logging.error(f"❌ Error al obtener la URL de inserción del informe '{report_id}' del workspace '{workspace_id}': {e}")
        raise Exception(f"Error al obtener la URL de inserción del informe '{report_id}' del workspace '{workspace_id}': {e}")
