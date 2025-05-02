"""
mapeo_acciones.py (Corregido v3 - Final)

Centraliza el mapeo de nombres de acción a funciones ejecutables.
Usa alias renombrados para SharePoint y tipo Callable genérico.
"""

import logging
from typing import Dict, Any, Callable

logger = logging.getLogger(__name__)

# Definir un tipo para las funciones de acción para mejorar legibilidad
# Usar 'Any' como tipo de retorno para permitir flexibilidad manejada por preparar_respuesta
AccionCallable = Callable[[Dict[str, Any], Dict[str, str]], Any]

# --- Importaciones de funciones por categoría ---
# Envolver en try-except para manejo robusto de errores de importación

acciones_disponibles: Dict[str, AccionCallable] = {}

# Correo
try:
    from actions.correo import (
        listar_correos, leer_correo, enviar_correo, guardar_borrador, enviar_borrador,
        responder_correo, reenviar_correo, eliminar_correo
    )
    acciones_disponibles.update({
        "mail_listar": listar_correos, "mail_leer": leer_correo, "mail_enviar": enviar_correo,
        "mail_guardar_borrador": guardar_borrador, "mail_enviar_borrador": enviar_borrador,
        "mail_responder": responder_correo, "mail_reenviar": reenviar_correo, "mail_eliminar": eliminar_correo,
    })
except ImportError as e: logger.warning(f"No se pudo importar actions.correo: {e}")

# Calendario
try:
    from actions.calendario import (
        listar_eventos, crear_evento, actualizar_evento, eliminar_evento, crear_reunion_teams
    )
    acciones_disponibles.update({
        "cal_listar_eventos": listar_eventos, "cal_crear_evento": crear_evento,
        "cal_actualizar_evento": actualizar_evento, "cal_eliminar_evento": eliminar_evento,
        "cal_crear_reunion_teams": crear_reunion_teams,
    })
except ImportError as e: logger.warning(f"No se pudo importar actions.calendario: {e}")

# OneDrive (/me/drive)
try:
    from actions.onedrive import (
        listar_archivos as od_listar_archivos, subir_archivo as od_subir_archivo,
        descargar_archivo as od_descargar_archivo, eliminar_archivo as od_eliminar_archivo,
        crear_carpeta as od_crear_carpeta, mover_archivo as od_mover_archivo,
        copiar_archivo as od_copiar_archivo, obtener_metadatos_archivo as od_obtener_metadatos_archivo,
        actualizar_metadatos_archivo as od_actualizar_metadatos_archivo
    )
    acciones_disponibles.update({
        "od_listar_archivos": od_listar_archivos, "od_subir_archivo": od_subir_archivo,
        "od_descargar_archivo": od_descargar_archivo, # Devuelve bytes
        "od_eliminar_archivo": od_eliminar_archivo, "od_crear_carpeta": od_crear_carpeta,
        "od_mover_archivo": od_mover_archivo, "od_copiar_archivo": od_copiar_archivo,
        "od_obtener_metadatos_archivo": od_obtener_metadatos_archivo,
        "od_actualizar_metadatos_archivo": od_actualizar_metadatos_archivo,
    })
except ImportError as e: logger.warning(f"No se pudo importar actions.onedrive: {e}")

# SharePoint
try:
    # CORRECCIÓN: Usar los nombres finales de las funciones en sharepoint.py
    from actions.sharepoint import (
        crear_lista as sp_crear_lista, listar_listas as sp_listar_listas,
        agregar_elemento_lista as sp_agregar_elemento_lista,
        listar_elementos_lista as sp_listar_elementos_lista,
        actualizar_elemento_lista as sp_actualizar_elemento_lista,
        eliminar_elemento_lista as sp_eliminar_elemento_lista,
        listar_documentos_biblioteca as sp_listar_documentos_biblioteca,
        subir_documento as sp_subir_documento,
        eliminar_archivo_biblioteca as sp_eliminar_archivo_biblioteca, # Nombre final
        crear_carpeta_biblioteca as sp_crear_carpeta_biblioteca, # Nombre final
        mover_archivo_biblioteca as sp_mover_archivo_biblioteca, # Nombre final
        copiar_archivo_biblioteca as sp_copiar_archivo_biblioteca, # Nombre final
        obtener_metadatos_archivo_biblioteca as sp_obtener_metadatos_archivo_biblioteca, # Nombre final
        actualizar_metadatos_archivo_biblioteca as sp_actualizar_metadatos_archivo_biblioteca, # Nombre final
        obtener_contenido_archivo_biblioteca as sp_obtener_contenido_archivo_biblioteca, # Nombre final
        actualizar_contenido_archivo_biblioteca as sp_actualizar_contenido_archivo_biblioteca, # Nombre final
        crear_enlace_compartido_archivo_biblioteca as sp_crear_enlace_compartido_archivo_biblioteca, # Nombre final
        # Funciones de Memoria
        guardar_dato_memoria as sp_guardar_dato_memoria,
        recuperar_datos_sesion as sp_recuperar_datos_sesion,
        eliminar_dato_memoria as sp_eliminar_dato_memoria,
        eliminar_memoria_sesion as sp_eliminar_memoria_sesion,
        exportar_datos_lista as sp_exportar_datos_lista,
    )
    acciones_disponibles.update({
        # Listas SP
        "sp_crear_lista": sp_crear_lista, "sp_listar_listas": sp_listar_listas,
        "sp_agregar_elemento_lista": sp_agregar_elemento_lista, "sp_listar_elementos_lista": sp_listar_elementos_lista,
        "sp_actualizar_elemento_lista": sp_actualizar_elemento_lista, "sp_eliminar_elemento_lista": sp_eliminar_elemento_lista,
        # Documentos SP
        "sp_listar_documentos_biblioteca": sp_listar_documentos_biblioteca, "sp_subir_documento": sp_subir_documento,
        "sp_eliminar_archivo_biblioteca": sp_eliminar_archivo_biblioteca,
        "sp_crear_carpeta_biblioteca": sp_crear_carpeta_biblioteca,
        "sp_mover_archivo_biblioteca": sp_mover_archivo_biblioteca, # CORREGIDO
        "sp_copiar_archivo_biblioteca": sp_copiar_archivo_biblioteca, # CORREGIDO
        "sp_obtener_metadatos_archivo_biblioteca": sp_obtener_metadatos_archivo_biblioteca, # CORREGIDO
        "sp_actualizar_metadatos_archivo_biblioteca": sp_actualizar_metadatos_archivo_biblioteca, # CORREGIDO
        "sp_obtener_contenido_archivo_biblioteca": sp_obtener_contenido_archivo_biblioteca, # CORREGIDO (Devuelve bytes)
        "sp_actualizar_contenido_archivo_biblioteca": sp_actualizar_contenido_archivo_biblioteca, # CORREGIDO
        "sp_crear_enlace_compartido_archivo_biblioteca": sp_crear_enlace_compartido_archivo_biblioteca, # CORREGIDO
        # Memoria SP
        "sp_guardar_dato_memoria": sp_guardar_dato_memoria, "sp_recuperar_datos_sesion": sp_recuperar_datos_sesion,
        "sp_eliminar_dato_memoria": sp_eliminar_dato_memoria, "sp_eliminar_memoria_sesion": sp_eliminar_memoria_sesion,
        "sp_exportar_datos_lista": sp_exportar_datos_lista,
    })
except ImportError as e: logger.warning(f"No se pudo importar actions.sharepoint: {e}")
except AttributeError as e: logger.warning(f"Error de atributo importando desde actions.sharepoint: {e}. Verifica nombres.")


# Teams
try:
    from actions.teams import (
        listar_chats, obtener_chat, crear_chat, enviar_mensaje_chat, obtener_mensajes_chat,
        actualizar_mensaje_chat, eliminar_mensaje_chat, listar_equipos, obtener_equipo,
        crear_equipo, archivar_equipo, unarchivar_equipo, eliminar_equipo, listar_canales,
        obtener_canal, crear_canal, actualizar_canal, eliminar_canal, enviar_mensaje_canal
    )
    acciones_disponibles.update({
        "team_listar_chats": listar_chats, "team_obtener_chat": obtener_chat, "team_crear_chat": crear_chat,
        "team_enviar_mensaje_chat": enviar_mensaje_chat, "team_obtener_mensajes_chat": obtener_mensajes_chat,
        "team_actualizar_mensaje_chat": actualizar_mensaje_chat, "team_eliminar_mensaje_chat": eliminar_mensaje_chat,
        "team_listar_equipos": listar_equipos, "team_obtener_equipo": obtener_equipo, "team_crear_equipo": crear_equipo,
        "team_archivar_equipo": archivar_equipo, "team_unarchivar_equipo": unarchivar_equipo, "team_eliminar_equipo": eliminar_equipo,
        "team_listar_canales": listar_canales, "team_obtener_canal": obtener_canal, "team_crear_canal": crear_canal,
        "team_actualizar_canal": actualizar_canal, "team_eliminar_canal": eliminar_canal,
        "team_enviar_mensaje_canal": enviar_mensaje_canal,
    })
except ImportError as e: logger.warning(f"No se pudo importar actions.teams: {e}")

# Office (Word/Excel)
try:
    from actions.office import (
        crear_documento_word, insertar_texto_word, obtener_documento_word,
        crear_excel, escribir_celda_excel, leer_celda_excel,
        crear_tabla_excel, agregar_datos_tabla_excel
    )
    acciones_disponibles.update({
        "office_crear_word": crear_documento_word, "office_insertar_texto_word": insertar_texto_word,
        "office_obtener_documento_word": obtener_documento_word, # Devuelve bytes
        "office_crear_excel": crear_excel, "office_escribir_celda_excel": escribir_celda_excel,
        "office_leer_celda_excel": leer_celda_excel, "office_crear_tabla_excel": crear_tabla_excel,
        "office_agregar_datos_tabla_excel": agregar_datos_tabla_excel,
    })
except ImportError as e: logger.warning(f"No se pudo importar actions.office: {e}")

# Planner & ToDo
try:
    from actions.planner_todo import (
        listar_planes, obtener_plan, crear_plan, actualizar_plan, eliminar_plan,
        listar_tareas_planner, crear_tarea_planner, actualizar_tarea_planner, eliminar_tarea_planner,
        listar_listas_todo, crear_lista_todo, actualizar_lista_todo, eliminar_lista_todo,
        listar_tareas_todo, crear_tarea_todo, actualizar_tarea_todo, eliminar_tarea_todo,
        completar_tarea_todo
    )
    acciones_disponibles.update({
        "planner_listar_planes": listar_planes, "planner_obtener_plan": obtener_plan, "planner_crear_plan": crear_plan,
        "planner_actualizar_plan": actualizar_plan, "planner_eliminar_plan": eliminar_plan,
        "planner_listar_tareas": listar_tareas_planner, "planner_crear_tarea": crear_tarea_planner,
        "planner_actualizar_tarea": actualizar_tarea_planner, "planner_eliminar_tarea": eliminar_tarea_planner,
        "todo_listar_listas": listar_listas_todo, "todo_crear_lista": crear_lista_todo,
        "todo_actualizar_lista": actualizar_lista_todo, "todo_eliminar_lista": eliminar_lista_todo,
        "todo_listar_tareas": listar_tareas_todo, "todo_crear_tarea": crear_tarea_todo,
        "todo_actualizar_tarea": actualizar_tarea_todo, "todo_eliminar_tarea": eliminar_tarea_todo,
        "todo_completar_tarea": completar_tarea_todo,
    })
except ImportError as e: logger.warning(f"No se pudo importar actions.planner_todo: {e}")

# Power Automate
try:
    from actions.power_automate import (
        listar_flows, obtener_flow, crear_flow, actualizar_flow,
        eliminar_flow, ejecutar_flow, obtener_estado_ejecucion_flow
    )
    acciones_disponibles.update({
        "flow_listar": listar_flows, "flow_obtener": obtener_flow, "flow_crear": crear_flow,
        "flow_actualizar": actualizar_flow, "flow_eliminar": eliminar_flow, "flow_ejecutar": ejecutar_flow,
        "flow_obtener_estado_ejecucion": obtener_estado_ejecucion_flow,
    })
except ImportError as e: logger.warning(f"No se pudo importar actions.power_automate: {e}")

# Power BI
try:
    from actions.power_bi import (
        listar_workspaces, obtener_workspace, listar_dashboards,
        obtener_dashboard, listar_reports, obtener_reporte,
        listar_datasets, obtener_dataset, refrescar_dataset,
        obtener_estado_refresco_dataset, obtener_embed_url
    )
    acciones_disponibles.update({
        "pbi_listar_workspaces": listar_workspaces, "pbi_obtener_workspace": obtener_workspace,
        "pbi_listar_dashboards": listar_dashboards, "pbi_obtener_dashboard": obtener_dashboard,
        "pbi_listar_reports": listar_reports, "pbi_obtener_reporte": obtener_reporte,
        "pbi_listar_datasets": listar_datasets, "pbi_obtener_dataset": obtener_dataset,
        "pbi_refrescar_dataset": refrescar_dataset, "pbi_obtener_estado_refresco": obtener_estado_refresco_dataset,
        "pbi_obtener_embed_url": obtener_embed_url,
    })
except ImportError as e: logger.warning(f"No se pudo importar actions.power_bi: {e}")


# --- Verificación Final ---
logger.info(f"Acciones disponibles cargadas ({len(acciones_disponibles)}): {list(acciones_disponibles.keys())}")
if not acciones_disponibles:
    logger.error("¡Advertencia Crítica! No se cargó ninguna acción. Verifica imports y módulos en 'actions/'.")

