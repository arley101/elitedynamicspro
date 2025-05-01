import json
import logging
import requests
import azure.functions as func
from typing import Dict, Any, Callable, List, Optional, Union
from datetime import datetime, timezone
import os
import io

# --- Configuración de Logging ---
logger = logging.getLogger("azure.functions")
logger.setLevel(logging.INFO) # O logging.DEBUG

# --- Imports de Acciones (Desde la carpeta 'actions') ---
# Importa TODAS las funciones que quieres exponer desde tus módulos
# Asegúrate de que los nombres no colisionen o usa alias (import ... as ...)
try:
    # Correo
    from actions.correo import (
        listar_correos, leer_correo, enviar_correo, guardar_borrador,
        enviar_borrador, responder_correo, reenviar_correo, eliminar_correo
    )
    # Calendario
    from actions.calendario import (
        listar_eventos, crear_evento, actualizar_evento, eliminar_evento,
        crear_reunion_teams # Asumiendo que esta es la función relevante de calendario.py
    )
    # OneDrive
    from actions.onedrive import (
        listar_archivos as od_listar_archivos, # Usar alias para claridad/evitar colisión
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
         agregar_elemento as sp_agregar_elemento_lista, # Cambiado nombre para claridad
         listar_elementos as sp_listar_elementos_lista,
         actualizar_elemento as sp_actualizar_elemento_lista,
         eliminar_elemento as sp_eliminar_elemento_lista,
         listar_documentos_biblioteca as sp_listar_documentos_biblioteca,
         subir_documento as sp_subir_documento,
         eliminar_archivo as sp_eliminar_archivo_biblioteca, # Cambiado nombre para claridad
         crear_carpeta_biblioteca as sp_crear_carpeta_biblioteca,
         mover_archivo as sp_mover_archivo,
         copiar_archivo as sp_copiar_archivo,
         obtener_metadatos_archivo as sp_obtener_metadatos_archivo,
         actualizar_metadatos_archivo as sp_actualizar_metadatos_archivo,
         obtener_contenido_archivo as sp_obtener_contenido_archivo,
         actualizar_contenido_archivo as sp_actualizar_contenido_archivo,
         crear_enlace_compartido_archivo as sp_crear_enlace_compartido_archivo
         # obtener_site_root no se expone como acción directa
    )
    # Teams
    from actions.teams import (
        listar_chats as team_listar_chats, # Usar alias
        obtener_chat as team_obtener_chat,
        crear_chat as team_crear_chat,
        enviar_mensaje_chat as team_enviar_mensaje_chat,
        obtener_mensajes_chat as team_obtener_mensajes_chat,
        actualizar_mensaje_chat as team_actualizar_mensaje_chat,
        eliminar_mensaje_chat as team_eliminar_mensaje_chat,
        # listar_reuniones y crear_reunion_teams se manejan desde calendario.py
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
    logger.error(f"Error al importar módulos de acciones: {e}. Verifica la estructura y nombres de archivos en 'actions/'.")
    ALL_ACTIONS_LOADED = False
except Exception as e:
    logger.error(f"Error inesperado durante importación de acciones: {e}", exc_info=True)
    ALL_ACTIONS_LOADED = False


# --- Configuración de Logging ---
# (Se configura globalmente al inicio del archivo de función)
# logger = logging.getLogger("azure.functions")
# logger.setLevel(logging.INFO)

# --- Variables de Entorno (Leídas por necesidad donde hagan falta) ---
# Intentar leerlas una vez si se usan en múltiples sitios, o dentro de las funciones que las necesiten
# Ejemplo: Leer las de SP aquí para pasarlas a sus funciones si fuera necesario,
#          o mejor aún, importarlas directamente en actions/sharepoint.py desde .. como hicimos con BASE_URL.
#          Por ahora, las acciones de PA y PBI leen sus propias variables de entorno directamente.

# --- Constantes Globales ---
BASE_URL = "https://graph.microsoft.com/v1.0" # Base para la mayoría de llamadas a Graph
GRAPH_API_TIMEOUT = 45 # Timeout para llamadas a Graph

# --- Mapeo de Acciones ---
# Construir el diccionario solo si todas las acciones se cargaron
if ALL_ACTIONS_LOADED:
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
        "team_eliminar_mensaje_chat": team_eliminar_mensaje_chat, # "team_listar_reuniones" eliminada
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

# --- Función Principal (Refactorizada para usar token entrante) ---

def main(req: func.HttpRequest) -> func.HttpResponse:
    """Punto de entrada principal. Extrae token delegado, llama a la acción apropiada."""
    logging.info(f'Python HTTP trigger function procesando solicitud. Method: {req.method}, URL: {req.url}')
    invocation_id = req.headers.get('X-Azure-Functions-InvocationId', 'N/A')
    logging.info(f"Invocation ID: {invocation_id}")

    if not ALL_ACTIONS_LOADED:
         return func.HttpResponse("Error interno: No se pudieron cargar las acciones.", status_code=500)

    accion: Optional[str] = None
    parametros: Dict[str, Any] = {}
    funcion_a_ejecutar: Optional[Callable] = None
    resultado: Any = None
    request_headers: Dict[str, str] = {} # Headers para pasar a las funciones de acción

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
                      parametros['nombre_archivo_original'] = file.filename # Pasar nombre original
                 # Asegurarse que 'accion' se leyó aunque no haya file
                 if not accion and 'accion' in req.form: accion = req.form['accion']
            else: accion = req.params.get('accion'); parametros = dict(req.params)
        else: accion = req.params.get('accion'); parametros = dict(req.params)
        if 'accion' in parametros: del parametros['accion']

        if not accion or not isinstance(accion, str):
            return func.HttpResponse("Falta parámetro 'accion' (string).", status_code=400)
        log_params = {k: f"<bytes len={len(v)}>" if isinstance(v, bytes) else v for k, v in parametros.items()}
        logging.info(f"Invocation {invocation_id}: Acción: '{accion}'. Params: {log_params}")

        # --- 2. Extraer y Validar Token Delegado Entrante (OAuth de OpenAI) ---
        auth_header = req.headers.get('Authorization')
        if not auth_header or not auth_header.lower().startswith('bearer '):
            # --- Fallback Opcional a Token de Aplicación ---
            # Si QUISIERAS soportar acciones que NO necesitan usuario (ej: leer todos los sitios)
            # podrías intentar obtener un token de aplicación aquí usando CLIENT_ID/SECRET.
            # Por ahora, asumimos que TODAS las acciones requieren el token delegado de OpenAI.
            logger.error(f"Invocation {invocation_id}: Cabecera 'Authorization: Bearer <token>' faltante/inválida para '{accion}'.")
            return func.HttpResponse("No autorizado: Token delegado faltante.", status_code=401)

        # Preparar cabeceras básicas para Graph API (las funciones de acción pueden ajustarlas)
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
                # Aplicar conversiones genéricas o específicas aquí si es necesario
                # (Basado en el código anterior, puedes copiar/pegar la lógica de conversión de tipos)
                type_hints = getattr(funcion_a_ejecutar, '__annotations__', {})
                for param_name, param_type in type_hints.items():
                     if param_name in params_procesados and params_procesados[param_name] is not None:
                         original_value = params_procesados[param_name]
                         try:
                             if param_type is int: params_procesados[param_name] = int(original_value)
                             elif param_type is bool: params_procesados[param_name] = str(original_value).lower() in ['true', '1', 'yes']
                             elif param_type is float: params_procesados[param_name] = float(original_value)
                             elif param_type is datetime and isinstance(original_value, str): params_procesados[param_name] = datetime.fromisoformat(original_value.replace('Z', '+00:00'))
                         except (ValueError, TypeError) as conv_e: raise ValueError(f"Error convirtiendo param '{param_name}' a tipo {param_type}: {conv_e}")

                # Manejo especial para subidas de archivo
                if accion in ["od_subir_archivo", "sp_subir_documento", "sp_actualizar_contenido_archivo"]:
                    if 'contenido_bytes' not in params_procesados: raise ValueError(f"Acción '{accion}' requiere 'contenido_bytes'.")
                    # Renombrar si el nombre del parámetro en la función es diferente
                    if accion == "sp_subir_documento" and 'nombre_archivo' not in params_procesados:
                         params_procesados['nombre_archivo'] = params_procesados.get('nombre_archivo_original')
                    if accion == "sp_actualizar_contenido_archivo":
                         params_procesados['nuevo_contenido'] = params_procesados['contenido_bytes'] # Renombrar
                         if 'nombre_archivo' not in params_procesados:
                             params_procesados['nombre_archivo'] = params_procesados.get('nombre_archivo_original')


            except (ValueError, TypeError, KeyError) as conv_err:
                logger.error(f"Invocation {invocation_id}: Error parámetros '{accion}': {conv_err}. Recibido: {log_params}", exc_info=True)
                return func.HttpResponse(f"Parámetros inválidos '{accion}': {conv_err}", status_code=400)

            # --- 5. Llamar a la Función de Acción (Pasando Headers) ---
            logger.info(f"Invocation {invocation_id}: Ejecutando {funcion_a_ejecutar.__name__} con token delegado...")
            try:
                # Pasar las cabeceras y los parámetros procesados
                # Las funciones refactorizadas en actions/ aceptan 'headers'
                # Las funciones de PA y PBI lo ignorarán y usarán su propia auth interna
                resultado = funcion_a_ejecutar(headers=request_headers, **params_procesados)
                logger.info(f"Invocation {invocation_id}: Ejecución delegada de '{accion}' completada.")

            except TypeError as type_err:
                 if 'headers' in str(type_err) and 'required positional argument' in str(type_err):
                     logger.error(f"Invocation {invocation_id}: La función {funcion_a_ejecutar.__name__} no acepta 'headers'. ¡REFACTORIZAR actions/{funcion_a_ejecutar.__module__}.py!", exc_info=True)
                     return func.HttpResponse(f"Error interno: Acción '{accion}' no refactorizada correctamente.", status_code=500)
                 else:
                     logger.error(f"Invocation {invocation_id}: Error argumento {funcion_a_ejecutar.__name__}: {type_err}. Params: {params_procesados}", exc_info=True)
                     return func.HttpResponse(f"Error argumentos '{accion}': {type_err}", status_code=400)
            except Exception as exec_err:
                logger.exception(f"Invocation {invocation_id}: Error ejecución delegada '{accion}': {exec_err}")
                error_msg = f"Error al ejecutar '{accion}'."
                status_code = 500
                if isinstance(exec_err, requests.exceptions.RequestException) and exec_err.response is not None:
                    status_code = exec_err.response.status_code
                    try: error_details = exec_err.response.json(); error_msg = f"Error API ({status_code}): {error_details.get('error', {}).get('message', exec_err.response.text)}"
                    except json.JSONDecodeError: error_msg = f"Error API ({status_code}): {exec_err.response.text}"
                else: error_msg = f"Error interno al ejecutar '{accion}': {exec_err}"
                return func.HttpResponse(error_msg, status_code=min(status_code, 599)) # Limitar a < 600


            # --- 6. Devolver Resultado ---
            if isinstance(resultado, bytes):
                 filename = os.path.basename(parametros.get('nombre_archivo') or parametros.get('ruta_archivo') or 'download')
                 return func.HttpResponse(resultado, mimetype="application/octet-stream", headers={'Content-Disposition': f'attachment; filename="{filename}"'}, status_code=200)
            elif isinstance(resultado, (dict, list)):
                 try: return func.HttpResponse(json.dumps(resultado, default=str), mimetype="application/json", status_code=200)
                 except TypeError as serialize_err: logger.error(f"Error serializar JSON '{accion}': {serialize_err}.", exc_info=True); return func.HttpResponse("Error interno: Respuesta no serializable.", status_code=500)
            elif isinstance(resultado, requests.Response):
                 logger.warning(f"Acción '{accion}' devolvió objeto Response crudo. Status: {resultado.status_code}")
                 return func.HttpResponse(resultado.text, status_code=resultado.status_code, mimetype=resultado.headers.get('Content-Type', 'text/plain'))
            else:
                 return func.HttpResponse(str(resultado), mimetype="text/plain", status_code=200)

        else: # Acción no encontrada
            logger.warning(f"Invocation {invocation_id}: Acción '{accion}' no reconocida.");
            acciones_validas = list(acciones_disponibles.keys());
            return func.HttpResponse(f"Acción '{accion}' no reconocida. Válidas: {acciones_validas}", status_code=400)

    except Exception as e: # Error general en main
        func_name = getattr(funcion_a_ejecutar, '__name__', 'N/A')
        logger.exception(f"Invocation {invocation_id}: Error GENERAL main() acción '{accion or '?'}' (Func: {func_name}): {e}")
        return func.HttpResponse("Error interno del servidor.", status_code=500)

# --- FIN: Función Principal ---
