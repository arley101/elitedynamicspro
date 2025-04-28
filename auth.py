import os
import requests
import logging
import json # Módulo json importado
from typing import Optional, Dict, Any, Union, List # Añadido Union, List

# Configuración básica de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# Considera usar el logger de Azure Functions si este código corre dentro de una función
# logger = logging.getLogger("azure.functions")
# logger.setLevel(logging.INFO)

class AuthManager:
    """
    Clase para gestionar la autenticación con Microsoft Identity, soportando flujos de aplicación y delegado.
    """
    def __init__(self):
        self.client_id = os.getenv('CLIENT_ID')
        self.client_secret = os.getenv('CLIENT_SECRET')
        self.tenant_id = os.getenv('TENANT_ID')
        self.scope_app = os.getenv('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')
        self.user = os.getenv('DELEGATED_USER')
        self.password = os.getenv('DELEGATED_PASS')
        self.scope_delegated = os.getenv('DELEGATED_SCOPE', 'offline_access https://graph.microsoft.com/User.Read')

        if not all([self.client_id, self.tenant_id]):
            logging.error("❌ Faltan CLIENT_ID o TENANT_ID. La autenticación no puede funcionar.")
            raise ValueError("Faltan variables de entorno críticas: CLIENT_ID o TENANT_ID.")

    def obtener_token(self, flujo: str = "aplicacion", servicio: str = "graph") -> str:
        # (Aquí va toda la lógica interna de obtener_token que ya tenías en tu clase)
        # ... (código de obtener_token de la clase omitido por brevedad, pero debe estar aquí) ...
        logging.info(f"Intentando obtener token para servicio: {servicio}, flujo: {flujo}")
        token = None

        if servicio == "graph":
            if flujo == "aplicacion":
                if not self.client_secret:
                     logging.error("❌ Falta CLIENT_SECRET para el flujo de aplicación.")
                     raise ValueError("Falta la variable de entorno: CLIENT_SECRET para el flujo de aplicación.")
                token = self._obtener_token_aplicacion_graph()
            elif flujo == "delegado":
                if not all([self.user, self.password]):
                    logging.error("❌ Faltan DELEGATED_USER o DELEGATED_PASS para el flujo delegado.")
                    raise ValueError("Faltan variables de entorno: DELEGATED_USER o DELEGATED_PASS para el flujo delegado.")
                logging.warning("Intentando obtener token delegado usando flujo ROPC (grant_type=password). ¡Flujo no recomendado!")
                token = self._obtener_token_delegado_graph()
            else:
                logging.error(f"❌ Flujo no reconocido para Graph API: {flujo}")
                raise ValueError(f"Flujo no reconocido para Graph API: {flujo}")
        else:
            logging.error(f"❌ Servicio no implementado: {servicio}")
            raise NotImplementedError(f"Servicio {servicio} aún no implementado.")

        if not token:
             logging.error(f"❌ No se pudo obtener el token para flujo '{flujo}', servicio '{servicio}'. Revisar logs anteriores.")
             raise Exception(f"Fallo inesperado al obtener token para flujo '{flujo}', servicio '{servicio}'.")
        return token


    def _obtener_token_aplicacion_graph(self) -> Optional[str]:
        # (Aquí va toda la lógica interna de _obtener_token_aplicacion_graph que ya tenías)
        # ... (código omitido por brevedad) ...
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {'client_id': self.client_id,'client_secret': self.client_secret,'scope': self.scope_app,'grant_type': 'client_credentials'}
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        response = None
        try:
            response = requests.post(url, data=data, headers=headers)
            response.raise_for_status()
            token_data = response.json()
            token = token_data.get('access_token')
            if not token: logging.error(f"❌ No se encontró 'access_token' (app). Respuesta: {token_data}"); raise Exception("No 'access_token' (app).")
            logging.info("Token de aplicación obtenido exitosamente.")
            return token
        except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error red/HTTP (token app): {e}. Detalles: {error_details}"); raise Exception(f"Error red/HTTP (token app): {e}")
        except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (token app): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (token app): {e}")


    def _obtener_token_delegado_graph(self) -> Optional[str]:
        # (Aquí va toda la lógica interna de _obtener_token_delegado_graph que ya tenías)
        # --- ADVERTENCIA --- (ROPC no recomendado)
        # ... (código omitido por brevedad) ...
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {'client_id': self.client_id,'scope': self.scope_delegated,'grant_type': 'password','username': self.user,'password': self.password}
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        response = None
        try:
            response = requests.post(url, data=data, headers=headers)
            response.raise_for_status()
            token_data = response.json()
            token = token_data.get('access_token'); refresh_token = token_data.get('refresh_token')
            if not token: logging.error(f"❌ No se encontró 'access_token' (delegado). Respuesta: {token_data}"); raise Exception("No 'access_token' (delegado).")
            logging.info(f"Token delegado obtenido. Refresh token {'obtenido' if refresh_token else 'NO obtenido'}.")
            return token
        except requests.exceptions.RequestException as e: error_details = getattr(e.response, 'text', str(e)); logging.error(f"❌ Error red/HTTP (token delegado): {e}. Detalles: {error_details}"); raise Exception(f"Error red/HTTP (token delegado): {e}")
        except json.JSONDecodeError as e: response_text = getattr(response, 'text', 'No response object available'); logging.error(f"❌ Error JSON (token delegado): {e}. Respuesta: {response_text}"); raise Exception(f"Error JSON (token delegado): {e}")

# --- Instancia Global ---
# Se crea una instancia de AuthManager que puede ser usada por las funciones globales
auth_manager = AuthManager()

# --- Funciones Globales de Conveniencia ---

# !!!!! ESTA ES LA FUNCIÓN QUE FALTABA EN TU VERSIÓN !!!!!
def obtener_token(flujo: str = "aplicacion", servicio: str = "graph") -> str:
     """Función global simple para obtener solo el token crudo.
        Llama al método de la instancia global auth_manager.
        Necesaria para que otros módulos puedan hacer 'from auth import obtener_token'.
     """
     return auth_manager.obtener_token(flujo=flujo, servicio=servicio)

def obtener_headers_auth(flujo: str = "aplicacion", servicio: str = "graph") -> Dict[str, str]:
    """Obtiene las cabeceras de autorización con un token fresco."""
    # Llama a la función global obtener_token (que usa el manager)
    token = obtener_token(flujo=flujo, servicio=servicio)
    return {'Authorization': f'Bearer {token}','Content-Type': 'application/json'}
