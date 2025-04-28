import os
import requests
import logging
import json # <--- CORRECCIÓN: Módulo json importado
from typing import Optional, Dict, Any

# Configuración básica de logging
# Considera usar el logger de Azure Functions si este código corre dentro de una función
# logger = logging.getLogger("azure.functions") 
# logger.setLevel(logging.INFO)
# Si es un módulo independiente, la configuración básica está bien:
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
        self.scope_delegated = os.getenv('DELEGATED_SCOPE', 'offline_access https://graph.microsoft.com/User.Read') # Valor por defecto para delegado

        if not all([self.client_id, self.tenant_id]):
            logging.error("❌ Faltan CLIENT_ID o TENANT_ID. La autenticación no puede funcionar.")
            # En una Azure Function, quizás quieras devolver una respuesta HTTP de error aquí
            raise ValueError("Faltan variables de entorno críticas: CLIENT_ID o TENANT_ID.")

    def obtener_token(self, flujo: str = "aplicacion", servicio: str = "graph") -> str:
        """
        Obtiene un token de acceso de Microsoft Identity.

        Args:
            flujo: El tipo de flujo de autenticación a utilizar: "aplicacion" o "delegado".
            servicio: El servicio para el que se solicita el token: "graph" (por defecto).

        Returns:
            El token de acceso como una cadena.

        Raises:
            ValueError: Si el flujo o el servicio no son reconocidos, o faltan credenciales delegadas.
            NotImplementedError: Si el servicio no está implementado.
            Exception: Si ocurre un error de red o API al obtener el token.
        """
        logging.info(f"Intentando obtener token para servicio: {servicio}, flujo: {flujo}")
        token = None # Inicializar token

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
                token = self._obtener_token_delegado_graph()
            else:
                logging.error(f"❌ Flujo no reconocido para Graph API: {flujo}")
                raise ValueError(f"Flujo no reconocido para Graph API: {flujo}")
        else:
            logging.error(f"❌ Servicio no implementado: {servicio}")
            raise NotImplementedError(f"Servicio {servicio} aún no implementado.")

        if not token:
             # Si llegamos aquí, algo falló en las funciones internas pero no lanzaron excepción (no debería pasar con raise_for_status)
             logging.error(f"❌ No se pudo obtener el token para flujo '{flujo}', servicio '{servicio}'. Revisar logs anteriores.")
             raise Exception(f"Fallo inesperado al obtener token para flujo '{flujo}', servicio '{servicio}'.")
        
        return token


    def _obtener_token_aplicacion_graph(self) -> Optional[str]:
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
        headers = {'Content-Type': 'application/x-www-form-urlencoded'} # Importante para este grant_type
        response = None # Inicializar response fuera del try
        try:
            response = requests.post(url, data=data, headers=headers)
            response.raise_for_status()  # Lanza una excepción para códigos de error HTTP (4xx o 5xx)
            token_data = response.json()
            token = token_data.get('access_token')
            if not token:
                 logging.error(f"❌ No se encontró 'access_token' en la respuesta de token de aplicación. Respuesta: {token_data}")
                 raise Exception("No se pudo obtener el 'access_token' de la respuesta.")
            logging.info("Token de aplicación obtenido exitosamente.")
            return token
        except requests.exceptions.RequestException as e:
            error_details = getattr(e.response, 'text', str(e))
            logging.error(f"❌ Error de red/HTTP al obtener el token de aplicación para Graph: {e}. Detalles: {error_details}")
            raise Exception(f"Error de red/HTTP al obtener el token de aplicación para Graph: {e}")
        except json.JSONDecodeError as e:
            # Si response existe y tuvo éxito (2xx) pero el JSON es inválido (raro)
            # O si raise_for_status falló y la respuesta de error no es JSON válido
            response_text = getattr(response, 'text', 'No response object available')
            logging.error(f"❌ Error al decodificar la respuesta JSON (token aplicación): {e}. Respuesta: {response_text}")
            raise Exception(f"Error al decodificar la respuesta JSON (token aplicación): {e}")

    # --- ADVERTENCIA ---
    # El flujo Resource Owner Password Credentials (ROPC) ('grant_type': 'password')
    # NO es recomendado por Microsoft por razones de seguridad.
    # Evítalo si es posible. Considera flujos interactivos o de credenciales de dispositivo.
    # Puede ser bloqueado por políticas de Acceso Condicional.
    # Referencia: https://docs.microsoft.com/es-es/azure/active-directory/develop/v2-oauth-ropc
    def _obtener_token_delegado_graph(self) -> Optional[str]:
        """
        Obtiene un token de acceso para la API de Microsoft Graph usando el flujo de contraseña de usuario (delegado - ROPC).
        ¡¡ Flujo NO RECOMENDADO por seguridad !!
        """
        logging.warning("Intentando obtener token delegado usando flujo ROPC (grant_type=password). ¡Flujo no recomendado!")
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            'client_id': self.client_id,
            'scope': self.scope_delegated,
            'grant_type': 'password',
            'username': self.user,
            'password': self.password
            # 'client_secret' a veces es necesario para clientes confidenciales incluso con ROPC, verifica tu registro de app
            # 'client_secret': self.client_secret 
        }
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        response = None
        try:
            response = requests.post(url, data=data, headers=headers)
            response.raise_for_status()  # Lanza una excepción para códigos de error
            token_data = response.json()
            token = token_data.get('access_token')
            refresh_token = token_data.get('refresh_token') # Podrías querer guardar y usar el refresh token
            if not token:
                 logging.error(f"❌ No se encontró 'access_token' en la respuesta de token delegado. Respuesta: {token_data}")
                 raise Exception("No se pudo obtener el 'access_token' de la respuesta delegada.")
            logging.info(f"Token de delegado obtenido exitosamente. Refresh token {'obtenido' if refresh_token else 'NO obtenido'}.")
            return token
        except requests.exceptions.RequestException as e:
            error_details = getattr(e.response, 'text', str(e))
            logging.error(f"❌ Error de red/HTTP al obtener el token de delegado para Graph: {e}. Detalles: {error_details}")
            # Errores comunes aquí: AADSTS70002 (credenciales inválidas), AADSTS50126 (usuario/pass incorrecto),
            # AADSTS65001 (permisos no concedidos), políticas de Acceso Condicional bloqueando.
            raise Exception(f"Error de red/HTTP al obtener el token de delegado para Graph: {e}")
        except json.JSONDecodeError as e:
            response_text = getattr(response, 'text', 'No response object available')
            logging.error(f"❌ Error al decodificar la respuesta JSON (token delegado): {e}. Respuesta: {response_text}")
            raise Exception(f"Error al decodificar la respuesta JSON (token delegado): {e}")


# Instancia global para que otros módulos puedan importarla y usarla fácilmente
# Considera patrones de inyección de dependencias para aplicaciones más grandes.
auth_manager = AuthManager()

# Función global de conveniencia para simplificar llamadas desde otros módulos
def obtener_headers_auth(flujo: str = "aplicacion", servicio: str = "graph") -> Dict[str, str]:
    """
    Obtiene las cabeceras de autorización con un token fresco.

    Args:
        flujo: "aplicacion" o "delegado"
        servicio: "graph"

    Returns:
        Un diccionario con la cabecera 'Authorization' y 'Content-Type'.
    """
    token = auth_manager.obtener_token(flujo=flujo, servicio=servicio)
    return {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json' # Por defecto para Graph API
    }

# Nota: La función `obtener_token` original que tenías al final parecía incompleta.
# La he reemplazado por `obtener_headers_auth` que es más útil generalmente,
# ya que devuelve el diccionario de cabeceras listo para usar con `requests`.
# Si realmente solo necesitas el token crudo, puedes volver a poner:
# def obtener_token(flujo: str = "aplicacion", servicio: str = "graph") -> str:
#     return auth_manager.obtener_token(flujo=flujo, servicio=servicio)
