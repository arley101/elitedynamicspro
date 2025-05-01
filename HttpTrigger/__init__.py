"""
HttpTrigger/__init__.py

Gestión de solicitudes HTTP para Azure Functions. Este módulo procesa acciones
específicas basadas en solicitudes entrantes mediante Graph API y servicios relacionados.
"""

import json
import logging
from typing import Dict, Any, Optional
import azure.functions as func
from shared.constants import BASE_URL, GRAPH_API_TIMEOUT  # Usar constantes compartidas
from helpers.validadores import validar_parametros  # Validaciones centralizadas
from helpers.ejecutor import ejecutar_accion  # Ejecución de funciones

# --- Configuración del Logger ---
logger = logging.getLogger("azure.functions")
logger.setLevel(logging.INFO)

# --- Cargar Acciones Disponibles ---
try:
    from mapeo_acciones import acciones_disponibles  # Diccionario centralizado de acciones
    ALL_ACTIONS_LOADED = True
    logger.info("Todos los módulos de acciones importados correctamente.")
except ImportError as e:
    logger.error(f"Error al importar acciones: {e}", exc_info=True)
    ALL_ACTIONS_LOADED = False

# --- Función Principal ---
def main(req: func.HttpRequest) -> func.HttpResponse:
    """
    Función principal que procesa la solicitud HTTP entrante.
    """
    logger.info("Procesando solicitud HTTP...")
    invocation_id = req.headers.get('X-Azure-Functions-InvocationId', 'N/A')

    # Validar si las acciones están cargadas
    if not ALL_ACTIONS_LOADED:
        return func.HttpResponse("Error interno: No se pudieron cargar las acciones.", status_code=500)

    try:
        # Extraer acción y parámetros
        accion, parametros = extraer_accion_y_parametros(req)
        if not accion:
            return func.HttpResponse("Falta parámetro 'accion'.", status_code=400)

        logger.info(f"Invocación {invocation_id}: Acción solicitada: {accion}")

        # Validar si la acción existe
        if accion not in acciones_disponibles:
            acciones_validas = list(acciones_disponibles.keys())
            return func.HttpResponse(f"Acción '{accion}' no reconocida. Válidas: {acciones_validas}", status_code=400)

        # Ejecutar la acción
        funcion_a_ejecutar = acciones_disponibles[accion]
        resultado = ejecutar_accion(funcion_a_ejecutar, parametros, req.headers)

        # Devolver el resultado
        return preparar_respuesta(resultado)

    except Exception as e:
        logger.exception(f"Error general: {e}")
        return func.HttpResponse("Error interno del servidor.", status_code=500)


# --- Funciones Auxiliares ---
def extraer_accion_y_parametros(req: func.HttpRequest) -> (Optional[str], Dict[str, Any]):
    """
    Extrae la acción y los parámetros de la solicitud HTTP.
    """
    try:
        if req.method in ('POST', 'PUT', 'PATCH') and 'application/json' in req.headers.get('Content-Type', '').lower():
            body = req.get_json()
            return body.get("accion"), body.get("parametros", {})
        else:
            return req.params.get("accion"), dict(req.params)
    except Exception as e:
        logger.error(f"Error al extraer parámetros: {e}")
        return None, {}

def preparar_respuesta(resultado: Any) -> func.HttpResponse:
    """
    Prepara la respuesta HTTP basada en el resultado.
    """
    if isinstance(resultado, (dict, list)):
        return func.HttpResponse(json.dumps(resultado, default=str), mimetype="application/json")
    elif isinstance(resultado, bytes):
        return func.HttpResponse(resultado, mimetype="application/octet-stream")
    else:
        return func.HttpResponse(str(resultado), mimetype="text/plain")
