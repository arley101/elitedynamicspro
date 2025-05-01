import json
import logging
import requests
import azure.functions as func
# Corregido: Añadir Any a los imports de typing
from typing import Dict, Any, Callable, List, Optional, Union
from datetime import datetime, timezone
import os
import io

# --- Configuración de Logging ---
# Corregido: Asegurar que está definido antes de cualquier uso
logger = logging.getLogger("azure.functions")
logger.setLevel(logging.INFO) # O logging.DEBUG

# --- Imports de Constantes Compartidas ---
try:
    # Usar import directo desde el paquete 'shared'
    from shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError as e:
    logger.critical(f"Error crítico: No se pudo importar desde 'shared.constants'. Verifica la estructura. {e}")
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_API_TIMEOUT = 45

# --- Imports de Acciones ---
try:
    # Corregido: Añadir Any donde sea necesario en las funciones importadas
    # (Las propias funciones definen sus tipos, pero Any puede ser necesario aquí para el Callable)
    # Correo
    from actions.correo import (
        listar_correos, leer_correo, enviar_correo, guardar_borrador,
        enviar_borrador, responder_correo, reenviar_correo, eliminar_correo
    )
    # Calendario
    from actions.calendario import (
        listar_eventos, crear_evento, actualizar_evento, eliminar_evento,
        crear_reunion_teams
    )
    # OneDrive
    from actions.onedrive import (
        listar_archivos as od_listar_archivos,
        subir_archivo as od_subir_archivo,
        descargar_archivo as od_descargar_archivo,
        eliminar_archivo as od_eliminar_archivo,
        crear_carpeta as od_crear_carpeta,
        mover_archivo as od_mover_archivo,
        copiar_archivo as od_copiar_archivo,
        obtener_metadatos_archivo as od_obtener_metadatos_archivo,
        actualizar_metadatos_archivo as od_actualizar_metadatos_archivo
    )
    # SharePoint
    from actions.sharepoint import (
         crear_lista as sp_crear_lista,
         listar_listas as sp_listar_listas,
         agregar_elemento as sp_agregar_elemento_lista,
         listar_elementos as sp_listar_elementos_lista,
         actualizar_elemento as sp_actualizar_elemento_lista,
         eliminar_elemento as sp_eliminar_elemento_lista,
         listar_documentos_biblioteca as sp_listar_documentos_biblioteca,
         subir_documento as sp_subir_documento,
         eliminar_archivo as sp_eliminar_archivo_biblioteca,
         crear_carpeta_biblioteca as sp_crear_carpeta_biblioteca,
         mover_archivo as sp_mover_archivo,
         copiar_archivo as sp_copiar_archivo,
         obtener_metadatos_archivo as sp_obtener_metadatos_archivo,
         actualizar_metadatos_archivo as sp_actualizar_metadatos_archivo,
         obtener_contenido_archivo as sp_obtener_contenido_archivo,
         actualizar_contenido_archivo as sp_actualizar_contenido_archivo,
         crear_enlace_compartido_archivo as sp_crear_enlace_compartido_archivo
    )
    # Teams
    from actions.teams import (
        listar_chats as team_listar_chats,
        obtener_chat as team_obtener_chat,
        crear_chat as team_crear_chat,
        enviar_mensaje_chat as team_enviar_mensaje_chat,
        obtener_mensajes_chat as team_obtener_mensajes_chat,
        actualizar_mensaje_chat as team_actualizar_mensaje_chat,
        eliminar_mensaje_chat as team_eliminar_mensaje_chat,
        listar_equipos as team_listar_equipos,
        obtener_equipo as team_obtener_equipo,
        crear_equipo as team_crear_equipo,
        archivar_equipo as team_archivar_equipo,
        unarchivar_equipo as team_unarchivar_equipo,
        eliminar_equipo as team_eliminar_equipo,
        listar_canales as team_listar_canales,
        obtener_canal as team_obtener_canal,
        crear_canal as team_crear_canal,
        actualizar_canal as team_actualizar_canal,
        eliminar_canal as team_eliminar_canal,
        enviar_mensaje_canal as team_enviar_mensaje_canal
    )
    # Office (Word/Excel via Graph)
    from actions.office import (
        crear_documento_word, insertar_texto_word, obtener_documento_word,
        crear_excel, escribir_celda_excel, leer_celda_excel,
        crear_tabla_excel, agregar_datos_tabla_excel
    )
    # Planner & ToDo
    from actions.planner_todo import (
        listar_planes, obtener_plan, crear_plan, actualizar_plan, eliminar_plan,
        listar_tareas_planner, crear_tarea_planner, actualizar_tarea_planner, eliminar_tarea_planner,
        listar_listas_todo, crear_lista_todo, actualizar_lista_todo, eliminar_lista_todo,
        listar_tareas_todo, crear_tarea_todo, actualizar_tarea_todo, eliminar_tarea_todo,
        completar_tarea_todo
    )
    # Power Automate
    from actions.power_automate import (
        listar_flows, obtener_flow, crear_flow, actualizar_flow,
        eliminar_flow, ejecutar_flow, obtener_estado_ejecucion_flow
    )
    # Power BI
    from actions.power_bi import (
        listar_workspaces, obtener_workspace, listar_dashboards, obtener_dashboard,
        listar_reports, obtener_reporte, listar_datasets, obtener_dataset,
        refrescar_dataset, obtener_estado_refresco_dataset, obtener_embed_url
    )
    ALL_ACTIONS_LOADED = True
    logger.info("Todos los módulos de acciones importados correctamente.")
except ImportError as e:
    logger.error(f"Error al importar módulos de acciones desde 'actions/' o 'shared/': {e}. Verifica la estructura y nombres de archivos.", exc_info=True)
    ALL_ACTIONS_LOADED = False
except Exception as e:
    logger.error(f"Error inesperado durante importación de acciones: {e}", exc_info=True)
    ALL_ACTIONS_LOADED = False


# --- Mapeo de Acciones ---
if ALL_ACTIONS_LOADED:
    # Corregido: Usar Any para el tipo de retorno del Callable si no es uniforme
    acciones_disponibles: Dict[str, Callable[..., Any]] = {
        # Correo
        "mail_listar": listar_correos, "mail_leer": leer_correo, "mail_enviar": enviar_correo,
        "mail_guardar_borrador": guardar_borrador, "mail_enviar_borrador": enviar_borrador,
        "mail_responder": responder_correo, "mail_reenviar": reenviar_correo, "mail_eliminar": eliminar_correo,
        # Calendario
        "cal_listar_eventos": listar_eventos, "cal_crear_evento": crear_evento,
        "cal_actualizar_evento": actualizar_evento, "cal_eliminar_evento": eliminar_evento,
        "cal_crear_reunion_teams": crear_reunion_teams,
        # OneDrive (/me)
        "od_listar_archivos": od_listar_archivos, "od_subir_archivo": od_subir_archivo,
        "od_descargar_archivo": od_descargar_archivo, "od_eliminar_archivo": od_eliminar_archivo,
        "od_crear_carpeta": od_crear_carpeta, "od_mover_archivo": od_mover_archivo,
        "od_copiar_archivo": od_copiar_archivo, "od_obtener_metadatos_archivo": od_obtener_metadatos_archivo,
        "od_actualizar_metadatos_archivo": od_actualizar_metadatos_archivo,
        # SharePoint
        "sp_crear_lista": sp_crear_lista, "sp_listar_listas": sp_listar_listas,
        "sp_agregar_elemento_lista": sp_agregar_elemento_lista, "sp_listar_elementos_lista": sp_listar_elementos_lista,
        "sp_actualizar_elemento_lista": sp_actualizar_elemento_lista, "sp_eliminar_elemento_lista": sp_eliminar_elemento_lista,
        "sp_listar_documentos_biblioteca": sp_listar_documentos_biblioteca, "sp_subir_documento": sp_subir_documento,
        "sp_eliminar_archivo_biblioteca": sp_eliminar_archivo_biblioteca, "sp_crear_carpeta_biblioteca": sp_crear_carpeta_biblioteca,
        "sp_mover_archivo": sp_mover_archivo, "sp_copiar_archivo": sp_copiar_archivo,
        "sp_obtener_metadatos_archivo": sp_obtener_metadatos_archivo, "sp_actualizar_metadatos_archivo": sp_actualizar_metadatos_archivo,
        "sp_obtener_contenido_archivo": sp_obtener_contenido_archivo, "sp_actualizar_contenido_archivo": sp_actualizar_contenido_archivo,
        "sp_crear_enlace_compartido_archivo": sp_crear_enlace_compartido_archivo,
        # Teams
        "team_listar_chats": team_listar_chats, "team_obtener_chat": team_obtener_chat,
        "team_crear_chat": team_crear_chat, "team_enviar_mensaje_chat": team_enviar_mensaje_chat,
        "team_obtener_mensajes_chat": team_obtener_mensajes_chat, "team_actualizar_mensaje_chat": team_actualizar_mensaje_chat,
        "team_eliminar_mensaje_chat": team_eliminar_mensaje_chat,
        "team_listar_equipos": team_listar_equipos, "team_obtener_equipo": team_obtener_equipo,
        "team_crear_equipo": team_crear_equipo, "team_archivar_equipo": team_archivar_equipo,
        "team_unarchivar_equipo": team_unarchivar_equipo, "team_eliminar_equipo": team_eliminar_equipo,
        "team_listar_canales": team_listar_canales, "team_obtener_canal": team_obtener_canal,
        "team_crear_canal": team_crear_canal, "team_actualizar_canal": team_actualizar_canal,
        "team_eliminar_canal": team_eliminar_canal, "team_enviar_mensaje_canal": team_enviar_mensaje_canal,
        # Office (Word/Excel)
        "office_crear_word": crear_documento_word, "office_insertar_texto_word": insertar_texto_word,
        "office_obtener_word": obtener_documento_word, "office_crear_excel": crear_excel,
        "office_escribir_celda": escribir_celda_excel, "office_leer_celda": leer_celda_excel,
        "office_crear_tabla_excel": crear_tabla_excel, "office_agregar_datos_tabla": agregar_datos_tabla_excel,
        # Planner & ToDo
        "planner_listar_planes": listar_planes, "planner_obtener_plan": obtener_plan,
        "planner_crear_plan": crear_plan, "planner_actualizar_plan": actualizar_plan,
        "planner_eliminar_plan": eliminar_plan, "planner_listar_tareas": listar_tareas_planner,
        "planner_crear_tarea": crear_tarea_planner, "planner_actualizar_tarea": actualizar_tarea_planner,
        "planner_eliminar_tarea": eliminar_tarea_planner,
        "todo_listar_listas": listar_listas_todo, "todo_crear_lista": crear_lista_todo,
        "todo_actualizar_lista": actualizar_lista_todo, "todo_eliminar_lista": eliminar_lista_todo,
        "todo_listar_tareas": listar_tareas_todo, "todo_crear_tarea": crear_tarea_todo,
        "todo_actualizar_tarea": actualizar_tarea_todo, "todo_eliminar_tarea": eliminar_tarea_todo,
        "todo_completar_tarea": completar_tarea_todo,
        # Power Automate
        "flow_listar": listar_flows, "flow_obtener": obtener_flow, "flow_crear": crear_flow,
        "flow_actualizar": actualizar_flow, "flow_eliminar": eliminar_flow, "flow_ejecutar": ejecutar_flow,
        "flow_obtener_estado_ejecucion": obtener_estado_ejecucion_flow,
        # Power BI
        "pbi_listar_workspaces": listar_workspaces, "pbi_obtener_workspace": obtener_workspace,
        "pbi_listar_dashboards": listar_dashboards, "pbi_obtener_dashboard": obtener_dashboard,
        "pbi_listar_reports": listar_reports, "pbi_obtener_reporte": obtener_reporte,
        "pbi_listar_datasets": listar_datasets, "pbi_obtener_dataset": obtener_dataset,
        "pbi_refrescar_dataset": refrescar_dataset, "pbi_obtener_estado_refresco": obtener_estado_refresco_dataset,
        "pbi_obtener_embed_url": obtener_embed_url,
    }
else:
    acciones_disponibles = {}
    logging.critical("Mapa de acciones vacío debido a errores de importación previos.")


# --- Función Principal ---

def main(req: func.HttpRequest) -> func.HttpResponse:
    """Punto de entrada principal. Extrae token delegado, llama a la acción apropiada."""
    logging.info(f'Python HTTP trigger function procesando solicitud. Method: {req.method}, URL: {req.url}')
    invocation_id = req.headers.get('X-Azure-Functions-InvocationId', 'N/A')
    logging.info(f"Invocation ID: {invocation_id}")

    if not ALL_ACTIONS_LOADED:
         return func.HttpResponse("Error interno: No se pudieron cargar las acciones.", status_code=500)

    accion: Optional[str] = None
    parametros: Dict[str, Any] = {}
    funcion_a_ejecutar: Optional[Callable[..., Any]] = None # Corregido tipo
    resultado: Any = None
    request_headers: Dict[str, str] = {}

    try:
        # --- 1. Extraer Acción y Parámetros ---
        content_type = req.headers.get('Content-Type', '').lower()
        if req.method in ('POST', 'PUT', 'PATCH'):
            if 'application/json' in content_type:
                try:
                    req_body = req.get_json(); assert isinstance(req_body, dict)
                    accion = req_body.get('accion'); params_input = req_body.get('parametros')
                    if isinstance(params_input, dict): parametros = params_input
                    else: parametros = {}
                except (ValueError, AssertionError): return func.HttpResponse("Cuerpo JSON inválido.", status_code=400)
            elif 'multipart/form-data' in content_type:
                 accion = req.form.get('accion'); parametros = {}
                 for key, value in req.form.items():
                     if key not in ['accion', 'file']: parametros[key] = value
                 file = req.files.get('file')
                 if file:
                      parametros['contenido_bytes'] = file.read()
                      if hasattr(file, 'filename') and file.filename:
                           parametros['nombre_archivo_original'] = file.filename
                 if not accion and 'accion' in req.form: accion = req.form['accion']
            else: accion = req.params.get('accion'); parametros = dict(req.params)
        else: accion = req.params.get('accion'); parametros = dict(req.params)
        if 'accion' in parametros: del parametros['accion']

        if not accion or not isinstance(accion, str):
            return func.HttpResponse("Falta parámetro 'accion' (string).", status_code=400)
        log_params = {k: f"<bytes len={len(v)}>" if isinstance(v, bytes) else v for k, v in parametros.items()}
        logging.info(f"Invocation {invocation_id}: Acción: '{accion}'. Params: {log_params}")

        # --- 2. Extraer y Validar Token Delegado Entrante ---
        auth_header = req.headers.get('Authorization')
        if not auth_header or not auth_header.lower().startswith('bearer '):
            logger.error(f"Invocation {invocation_id}: Cabecera 'Authorization: Bearer <token>' faltante/inválida para '{accion}'.")
            return func.HttpResponse("No autorizado: Token delegado faltante.", status_code=401)

        # Preparar cabeceras básicas
        request_headers = {
            'Authorization': auth_header,
            'Content-Type': 'application/json'
        }
        logger.info(f"Invocation {invocation_id}: Token Bearer delegado detectado.")

        # --- 3. Buscar y Ejecutar Función de Acción ---
        if accion in acciones_disponibles:
            funcion_a_ejecutar = acciones_disponibles[accion]
            logger.info(f"Invocation {invocation_id}: Mapeado a función: {funcion_a_ejecutar.__name__}")

            # --- 4. Validar/Convertir Parámetros ---
            params_procesados: Dict[str, Any] = {}
            try:
                params_procesados = parametros.copy()
                type_hints = getattr(funcion_a_ejecutar, '__annotations__', {})
                for param_name, param_type in type_hints.items():
                     # Saltar 'headers' y 'return' de las anotaciones
                     if param_name == 'headers' or param_name == 'return': continue
                     if param_name in params_procesados and params_procesados[param_name] is not None:
                         original_value = params_procesados[param_name]
                         try:
                             # Corregido: Añadir chequeo de tipo existente antes de convertir
                             target_type = param_type
                             # Simplificar si es Optional[T] -> obtener T
                             if hasattr(param_type, '__origin__') and param_type.__origin__ is Union:
                                 args = getattr(param_type, '__args__', ())
                                 non_none_args = [t for t in args if t is not type(None)]
                                 if len(non_none_args) == 1: target_type = non_none_args[0]

                             if target_type is int and not isinstance(original_value, int): params_procesados[param_name] = int(original_value)
                             elif target_type is bool and not isinstance(original_value, bool): params_procesados[param_name] = str(original_value).lower() in ['true', '1', 'yes']
                             elif target_type is float and not isinstance(original_value, float): params_procesados[param_name] = float(original_value)
                             elif target_type is datetime and isinstance(original_value, str): params_procesados[param_name] = datetime.fromisoformat(original_value.replace('Z', '+00:00'))
                             elif target_type is list and isinstance(original_value, str): params_procesados[param_name] = json.loads(original_value) # Intentar parsear JSON string a lista
                             elif target_type is dict and isinstance(original_value, str): params_procesados[param_name] = json.loads(original_value) # Intentar parsear JSON string a dict

                         except (ValueError, TypeError, json.JSONDecodeError) as conv_e: raise ValueError(f"Error convirtiendo param '{param_name}' a tipo {target_type}: {conv_e}")

                # Manejo especial para subidas
                upload_actions = ["od_subir_archivo", "sp_subir_documento", "sp_actualizar_contenido_archivo"]
                if accion in upload_actions:
                    if 'contenido_bytes' not in params_procesados: raise ValueError(f"Acción '{accion}' requiere 'contenido_bytes'.")
                    # Renombrar params si es necesario (ejemplo para sp_actualizar_contenido_archivo)
                    if accion == "sp_actualizar_contenido_archivo":
                        params_procesados['nuevo_contenido'] = params_procesados.pop('contenido_bytes') # Usar pop para evitar duplicados
                        if 'nombre_archivo' not in params_procesados: params_procesados['nombre_archivo'] = params_procesados.get('nombre_archivo_original')
                    # Renombrar para od_subir_archivo y sp_subir_documento si esperan otro nombre de bytes
                    # if accion == "od_subir_archivo": ...
                    # if accion == "sp_subir_documento": ...
                    # Asegurarse que el nombre del archivo exista
                    if 'nombre_archivo' not in params_procesados and 'nombre_archivo_destino' not in params_procesados:
                        params_procesados['nombre_archivo'] = params_procesados.get('nombre_archivo_original')
                    # Consolidar nombre (ej. la función siempre espera 'nombre_archivo')
                    if 'nombre_archivo_destino' in params_procesados and 'nombre_archivo' not in params_procesados:
                         params_procesados['nombre_archivo'] = params_procesados.pop('nombre_archivo_destino')


            except (ValueError, TypeError, KeyError) as conv_err:
                logger.error(f"Invocation {invocation_id}: Error parámetros '{accion}': {conv_err}. Recibido: {log_params}", exc_info=True)
                return func.HttpResponse(f"Parámetros inválidos '{accion}': {conv_err}", status_code=400)

            # --- 5. Llamar a la Función de Acción ---
            logger.info(f"Invocation {invocation_id}: Ejecutando {funcion_a_ejecutar.__name__}...")
            try:
                resultado = funcion_a_ejecutar(headers=request_headers, **params_procesados)
                logger.info(f"Invocation {invocation_id}: Ejecución de '{accion}' completada.")

            except TypeError as type_err:
                 import inspect; sig = inspect.signature(funcion_a_ejecutar)
                 if 'headers' not in sig.parameters: logger.error(f"Invocation {invocation_id}: ¡Falta 'headers' en {funcion_a_ejecutar.__name__}!", exc_info=True); return func.HttpResponse(f"Error interno: Acción '{accion}' no refactorizada.", status_code=500)
                 else: logger.error(f"Invocation {invocation_id}: Error argumento {funcion_a_ejecutar.__name__}: {type_err}. Params: {params_procesados}", exc_info=True); return func.HttpResponse(f"Error argumentos '{accion}': {type_err}", status_code=400)
            except Exception as exec_err:
                logger.exception(f"Invocation {invocation_id}: Error ejecución delegada '{accion}': {exec_err}")
                error_msg = f"Error al ejecutar '{accion}'." ; status_code = 500
                if isinstance(exec_err, requests.exceptions.RequestException) and exec_err.response is not None:
                    status_code = exec_err.response.status_code
                    try: error_details = exec_err.response.json(); error_msg = f"Error API ({status_code}): {error_details.get('error', {}).get('message', exec_err.response.text)}"
                    except json.JSONDecodeError: error_msg = f"Error API ({status_code}): {exec_err.response.text}"
                else: error_msg = f"Error interno al ejecutar '{accion}': {exec_err}"[:500]
                status_code = status_code if 400 <= status_code <= 599 else 500
                return func.HttpResponse(error_msg, status_code=status_code)

            # --- 6. Devolver Resultado ---
            if isinstance(resultado, bytes):
                 # Corregido: Uso de variable intermedia para basename
                 filename_base: str = parametros.get('nombre_archivo') or parametros.get('ruta_archivo') or 'download'
                 filename = os.path.basename(filename_base)
                 mimetype = "application/octet-stream"; logger.info(f"Invocation {invocation_id}: Devolviendo {len(resultado)} bytes. Filename: {filename}"); return func.HttpResponse(resultado, mimetype=mimetype, headers={'Content-Disposition': f'attachment; filename="{filename}"'}, status_code=200)
            elif isinstance(resultado, (dict, list)):
                 logger.info(f"Invocation {invocation_id}: Devolviendo resultado JSON.");
                 try: return func.HttpResponse(json.dumps(resultado, default=str), mimetype="application/json", status_code=200)
                 except TypeError as serialize_err: logger.error(f"Error serializar JSON '{accion}': {serialize_err}.", exc_info=True); return func.HttpResponse("Error interno: Respuesta no serializable.", status_code=500)
            elif isinstance(resultado, requests.Response):
                 logger.warning(f"Acción '{accion}' devolvió objeto Response. Status: {resultado.status_code}."); return func.HttpResponse(resultado.text, status_code=resultado.status_code, mimetype=resultado.headers.get('Content-Type', 'text/plain'))
            else:
                 logger.info(f"Invocation {invocation_id}: Devolviendo resultado como texto plano."); return func.HttpResponse(str(resultado), mimetype="text/plain", status_code=200)

        else: # Acción no encontrada
            logger.warning(f"Invocation {invocation_id}: Acción '{accion}' no reconocida.");
            acciones_validas = list(acciones_disponibles.keys());
            return func.HttpResponse(f"Acción '{accion}' no reconocida. Válidas: {acciones_validas}", status_code=400)

    except Exception as e: # Error general en main
        func_name = getattr(funcion_a_ejecutar, '__name__', 'N/A')
        logger.exception(f"Invocation {invocation_id}: Error GENERAL main() acción '{accion or '?'}' (Func: {func_name}): {e}")
        return func.HttpResponse("Error interno del servidor.", status_code=500)

# --- FIN: Función Principal ---
