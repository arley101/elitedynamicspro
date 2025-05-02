"""
HttpTrigger/__init__.py

Gestión de solicitudes HTTP para Azure Functions. Este módulo procesa acciones
específicas basadas en solicitudes entrantes mediante Graph API y servicios relacionados.

Correcciones:
- Se corrigió la anotación de tipo para el retorno de `extraer_accion_y_parametros`.
- Se importó `Tuple` desde `typing`.
"""

import json
import logging
# CORRECCIÓN: Importar Tuple
from typing import Dict, Any, Optional, Tuple
import azure.functions as func

# Asumiendo que estos helpers y constantes están en las rutas correctas
# Si 'shared' está al mismo nivel que 'HttpTrigger', el import debería ser:
# from ..shared.constants import BASE_URL, GRAPH_API_TIMEOUT
# Si 'helpers' está dentro de 'HttpTrigger', el import es correcto:
from .helpers.validadores import validar_parametros
from .helpers.ejecutor import ejecutar_accion
# Si 'mapeo_acciones.py' está al mismo nivel que 'HttpTrigger':
# from ..mapeo_acciones import acciones_disponibles
# Si está en la raíz del proyecto y la raíz está en PYTHONPATH:
from mapeo_acciones import acciones_disponibles
# Ajusta los imports según tu estructura final

# --- Configuración del Logger ---
logger = logging.getLogger("azure.functions") # Usar el logger estándar de Azure Functions
logger.setLevel(logging.INFO)

# --- Cargar Acciones Disponibles ---
# (El código de carga de mapeo_acciones ya maneja errores de importación)
ALL_ACTIONS_LOADED = bool(acciones_disponibles) # Verificar si el diccionario tiene acciones

# --- Función Principal ---
def main(req: func.HttpRequest) -> func.HttpResponse:
    """
    Función principal que procesa la solicitud HTTP entrante.
    """
    logger.info("Procesando solicitud HTTP...")
    # Obtener ID de invocación para trazabilidad
    invocation_id = req.headers.get('X-Azure-Functions-InvocationId', 'N/A')
    logger.info(f"Invocation ID: {invocation_id}")

    # Validar si las acciones están cargadas
    if not ALL_ACTIONS_LOADED:
        logger.error("Error crítico: No hay acciones disponibles cargadas.")
        return func.HttpResponse(
            "Error interno del servidor: No se pudieron cargar las acciones.",
            status_code=500
        )

    try:
        # Extraer acción y parámetros
        accion, parametros = extraer_accion_y_parametros(req)

        # Validar que se proporcionó una acción
        if not accion:
            logger.warning("Solicitud recibida sin parámetro 'accion'.")
            return func.HttpResponse(
                "Parámetro 'accion' es requerido.",
                status_code=400
            )

        logger.info(f"Acción solicitada: '{accion}' con parámetros: {parametros}")

        # Validar si la acción existe en el mapeo
        if accion not in acciones_disponibles:
            acciones_validas = list(acciones_disponibles.keys())
            logger.warning(f"Acción '{accion}' no reconocida. Acciones válidas: {acciones_validas}")
            return func.HttpResponse(
                f"Acción '{accion}' no reconocida. Las acciones válidas son: {acciones_validas}",
                status_code=400
            )

        # Obtener la función correspondiente
        funcion_a_ejecutar = acciones_disponibles[accion]

        # Validar y convertir parámetros usando type hints de la función destino
        # Usar get_type_hints para resolver forward references si las hubiera
        try:
            from typing import get_type_hints
            type_hints = get_type_hints(funcion_a_ejecutar)
            # Excluir 'headers' de la validación de parámetros de entrada si existe
            type_hints_params = {k: v for k, v in type_hints.items() if k != 'headers' and k != 'return'}
            if type_hints_params:
                 parametros = validar_parametros(parametros, type_hints_params)
                 logger.info(f"Parámetros validados y convertidos para '{accion}': {parametros}")
        except NameError: # Si la función no tiene type hints o hay error al obtenerlos
             logger.warning(f"No se pudieron obtener type hints para la función '{accion}'. Se usarán los parámetros tal cual.")
        except ValueError as val_err: # Capturar errores específicos de validación
            logger.warning(f"Error de validación de parámetros para '{accion}': {val_err}")
            return func.HttpResponse(f"Error en parámetros para '{accion}': {val_err}", status_code=400)


        # Ejecutar la acción usando el ejecutor centralizado
        logger.info(f"Ejecutando acción '{accion}'...")
        resultado = ejecutar_accion(funcion_a_ejecutar, parametros, req.headers)
        logger.info(f"Acción '{accion}' ejecutada exitosamente.")

        # Devolver el resultado
        return preparar_respuesta(resultado)

    # Manejo específico para errores de validación/valor
    except ValueError as ve:
        logger.warning(f"Error de valor durante el procesamiento de '{accion or 'acción desconocida'}': {ve}")
        return func.HttpResponse(f"Error en los datos proporcionados: {ve}", status_code=400)
    # Manejo general de excepciones
    except Exception as e:
        logger.exception(f"Error inesperado procesando la acción '{accion or 'acción desconocida'}': {e}")
        return func.HttpResponse(
            "Error interno del servidor durante la ejecución de la acción.",
            status_code=500
        )


# --- Funciones Auxiliares ---
# CORRECCIÓN: Anotación de tipo de retorno corregida
def extraer_accion_y_parametros(req: func.HttpRequest) -> Tuple[Optional[str], Dict[str, Any]]:
    """
    Extrae la acción y los parámetros de la solicitud HTTP (GET o POST JSON).
    """
    accion: Optional[str] = None
    parametros: Dict[str, Any] = {}

    try:
        # Priorizar cuerpo JSON para POST/PUT/PATCH
        if req.method in ('POST', 'PUT', 'PATCH') and req.get_body():
            try:
                body = req.get_json()
                accion = body.get("accion")
                parametros = body.get("parametros", {})
                # Log para depuración
                # logger.debug(f"Parámetros extraídos del cuerpo JSON: accion='{accion}', params={parametros}")
            except ValueError:
                logger.warning("No se pudo decodificar el cuerpo JSON de la solicitud.")
                # Podrías intentar leer como form data si es necesario
                pass # Continuar para verificar query params

        # Si no se encontró en el cuerpo o es GET, verificar query parameters
        if not accion:
            accion = req.params.get("accion")
            # Extraer todos los demás query params como parámetros
            # Excluir 'accion' si ya estaba presente
            parametros_query = {k: v for k, v in req.params.items() if k != 'accion'}
            # Fusionar con parámetros del body (si había), dando prioridad a los del body
            parametros = {**parametros_query, **parametros}
            # Log para depuración
            # if accion: logger.debug(f"Parámetros extraídos de query string: accion='{accion}', params={parametros}")


    except Exception as e:
        logger.error(f"Error inesperado al extraer acción y parámetros: {e}")
        # Devolver None, {} para indicar fallo en la extracción
        return None, {}

    # Devolver acción y parámetros (pueden ser None y {} si no se encontraron)
    return accion, parametros


def preparar_respuesta(resultado: Any) -> func.HttpResponse:
    """
    Prepara la respuesta HTTP basada en el tipo de resultado.
    """
    try:
        if resultado is None:
            # Éxito sin contenido (ej. para DELETE o acciones sin retorno)
            return func.HttpResponse(status_code=204)
        elif isinstance(resultado, func.HttpResponse):
            # Si la acción ya devuelve una HttpResponse, pasarla tal cual
            return resultado
        elif isinstance(resultado, (dict, list)):
            # Resultado JSON
            return func.HttpResponse(
                json.dumps(resultado, default=str, ensure_ascii=False), # default=str para manejar tipos no serializables como datetime
                mimetype="application/json",
                status_code=200
            )
        elif isinstance(resultado, bytes):
            # Resultado binario (ej. descarga de archivo)
            # Determinar mimetype adecuado si es posible, o usar genérico
            return func.HttpResponse(resultado, mimetype="application/octet-stream", status_code=200)
        elif isinstance(resultado, str):
             # Resultado de texto plano
             return func.HttpResponse(resultado, mimetype="text/plain", status_code=200)
        else:
            # Otros tipos, convertir a string
            return func.HttpResponse(str(resultado), mimetype="text/plain", status_code=200)
    except Exception as e:
        logger.exception(f"Error al preparar la respuesta HTTP: {e}")
        return func.HttpResponse("Error interno al formatear la respuesta.", status_code=500)


