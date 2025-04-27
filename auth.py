import os
import requests
import logging
from typing import Optional, Dict, Any

# Configuración básica de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class AuthManager:
    """
    Clase para gestionar la autenticación con Microsoft Identity, soportando flujos de aplicación y delegado.
    """

    def __init__(self):
        """
        Inicializa el AuthManager obteniendo las credenciales de las variables de entorno.
        """
        self.client_id = os.getenv('CLIENT_ID')
        self.client_secret = os.getenv('CLIENT_SECRET')
        self.tenant_id = os.getenv('TENANT_ID')
        self.scope_app = os.getenv('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')  # Valor por defecto para la aplicación
        self.user = os.getenv('DELEGATED_USER')
        self.password = os.getenv('DELEGATED_PASS')
        self.scope_delegated = os.getenv('DELEGATED_SCOPE', 'offline_access https://graph.microsoft.com/User.Read') #Valor por defecto para delegado

        if not all([self.client_id, self.tenant_id]):
            logging.error("❌ Faltan CLIENT_ID o TENANT_ID. La autenticación no puede funcionar.")
            raise Exception("Faltan CLIENT_ID o TENANT_ID.")

    def obtener_token(self, flujo: str = "aplicacion", servicio: str = "graph") -> str:
        """
        Obtiene un token de acceso de Microsoft Identity.

        Args:
            flujo: El tipo de flujo de autenticación a utilizar: "aplicacion" o "delegado".
            servicio: El servicio para el que se solicita el token: "graph" (por defecto).

        Returns:
            El token de acceso como una cadena.

        Raises:
            ValueError: Si el flujo o el servicio no son reconocidos.
            Exception: Si ocurre un error al obtener el token.
        """
        logging.info(f"Obteniendo token para servicio: {servicio}, flujo: {flujo}")
        if servicio == "graph":
            if flujo == "aplicacion":
                return self._obtener_token_aplicacion_graph()
            elif flujo == "delegado":
                if not all([self.user, self.password]):
                    logging.error("❌ Faltan DELEGATED_USER o DELEGATED_PASS para el flujo delegado.")
                    raise ValueError("Faltan DELEGATED_USER o DELEGATED_PASS para el flujo delegado.")
                return self._obtener_token_delegado_graph()
            else:
                logging.error(f"❌ Flujo no reconocido para Graph API: {flujo}")
                raise ValueError(f"Flujo no reconocido para Graph API: {flujo}")
        else:
            logging.error(f"❌ Servicio no implementado: {servicio}")
            raise NotImplementedError(f"Servicio {servicio} aún no implementado.")

    def _obtener_token_aplicacion_graph(self) -> str:
        """
        Obtiene un token de acceso para la API de Microsoft Graph usando el flujo de credenciales de aplicación.
        """
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'scope': self.scope_app,
            'grant_type': 'client_credentials'
        }
        try:
            response = requests.post(url, data=data)
            response.raise_for_status()  # Lanza una excepción para códigos de error
            token_data = response.json()
            token = token_data.get('access_token')
            logging.info("Token de aplicación obtenido exitosamente.")
            return token
        except requests.exceptions.RequestException as e:
            logging.error(f"❌ Error al obtener el token de aplicación para Graph: {e}")
            raise Exception(f"Error al obtener el token de aplicación para Graph: {e}")
        except json.JSONDecodeError as e:
            logging.error(f"❌ Error al decodificar la respuesta JSON: {e}. Respuesta: {response.text}")
            raise Exception(f"Error al decodificar la respuesta JSON: {e}")

    def _obtener_token_delegado_graph(self) -> str:
        """
        Obtiene un token de acceso para la API de Microsoft Graph usando el flujo de contraseña de usuario (delegado).
        """
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            'client_id': self.client_id,
            'scope': self.scope_delegated,
            'grant_type': 'password',
            'username': self.user,
            'password': self.password
        }
        try:
            response = requests.post(url, data=data)
            response.raise_for_status()  # Lanza una excepción para códigos de error
            token_data = response.json()
            token = token_data.get('access_token')
            logging.info("Token de delegado obtenido exitosamente.")
            return token
        except requests.exceptions.RequestException as e:
            logging.error(f"❌ Error al obtener el token de delegado para Graph: {e}")
            raise Exception(f"Error al obtener el token de delegado para Graph: {e}")
        except json.JSONDecodeError as e:
            logging.error(f"❌ Error al decodificar la respuesta JSON: {e}. Respuesta: {response.text}")
            raise Exception(f"Error al decodificar la respuesta JSON: {e}")


# Instancia global para que todos los módulos usen
auth_manager = AuthManager()

def obtener_token(flujo: str = "aplicacion", servicio: str = "graph") -> str:
    """
    Función global para obtener el token de acceso.  Esta es la función que deben llamar tus otros módulos.

    Args:
        flujo:  "aplicacion" o "delegado"
        servicio: "graph"

    Returns:
        El token de acceso
    """
    return auth_manager.obtener_token(flujo=flujo, servicio=servicio)
