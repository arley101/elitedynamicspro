"""
mapeo_acciones.py

Este archivo centraliza todas las acciones disponibles en el sistema
y las mapea a las funciones que las ejecutan. Las acciones están
organizadas por categorías para facilitar su mantenimiento.

Correcciones:
- Se ajustaron los nombres de las funciones importadas y usadas en el diccionario
  basándose en los errores 'AttributeError: ... maybe ...?' del log de mypy.
"""

import logging

logger = logging.getLogger(__name__)

# --- Importaciones de funciones por categoría ---
# Se envuelven en try-except para manejar posibles errores si un módulo
# de acción aún no está completamente implementado o tiene errores de importación.

try:
    # Correo
    from actions.correo import (
        listar_correos, leer_correo, enviar_correo, guardar_borrador, enviar_borrador,
        responder_correo, reenviar_correo, eliminar_correo
    )
    correo_loaded = True
except ImportError as e:
    logger.warning(f"No se pudo importar el módulo actions.correo: {e}")
    correo_loaded = False

try:
    # Calendario
    from actions.calendario import (
        listar_eventos, crear_evento, actualizar_evento, eliminar_evento, crear_reunion_teams
    )
    calendario_loaded = True
except ImportError as e:
    logger.warning(f"No se pudo importar el módulo actions.calendario: {e}")
    calendario_loaded = False

try:
    # OneDrive
    from actions.onedrive import (
        listar_archivos as od_listar_archivos, subir_archivo as od_subir_archivo,
        descargar_archivo as od_descargar_archivo, eliminar_archivo as od_eliminar_archivo,
        crear_carpeta as od_crear_carpeta, mover_archivo as od_mover_archivo,
        copiar_archivo as od_copiar_archivo, obtener_metadatos_archivo as od_obtener_metadatos_archivo,
        actualizar_metadatos_archivo as od_actualizar_metadatos_archivo
    )
    onedrive_loaded = True
except ImportError as e:
    logger.warning(f"No se pudo importar el módulo actions.onedrive: {e}")
    onedrive_loaded = False

try:
    # SharePoint
    # CORRECCIÓN: Se usan los nombres sugeridos por mypy (ej: agregar_elemento_lista)
    from actions.sharepoint import (
        crear_lista as sp_crear_lista, listar_listas as sp_listar_listas,
        agregar_elemento_lista as sp_agregar_elemento_lista, # CORREGIDO: agregar_elemento -> agregar_elemento_lista
        listar_elementos_lista as sp_listar_elementos_lista, # CORREGIDO: listar_elementos -> listar_elementos_lista
        actualizar_elemento_lista as sp_actualizar_elemento_lista,
        eliminar_elemento_lista as sp_eliminar_elemento_lista,
        listar_documentos_biblioteca as sp_listar_documentos_biblioteca,
        subir_documento as sp_subir_documento,
        # CORRECCIÓN: Asegúrate que la función 'eliminar_archivo_biblioteca' existe en sharepoint.py
        # Si no existe, debes implementarla o removerla del mapeo.
        # eliminar_archivo_biblioteca as sp_eliminar_archivo_biblioteca, # Nombre sugerido por mypy, verificar existencia
        crear_carpeta_biblioteca as sp_crear_carpeta_biblioteca,
        mover_archivo as sp_mover_archivo_sp, # Renombrado para evitar colisión con OneDrive
        copiar_archivo as sp_copiar_archivo_sp, # Renombrado para evitar colisión con OneDrive
        obtener_metadatos_archivo as sp_obtener_metadatos_archivo,
        actualizar_metadatos_archivo as sp_actualizar_metadatos_archivo,
        obtener_contenido_archivo as sp_obtener_contenido_archivo,
        actualizar_contenido_archivo as sp_actualizar_contenido_archivo,
        crear_enlace_compartido_archivo as sp_crear_enlace_compartido_archivo
    )
    sharepoint_loaded = True
except ImportError as e:
    logger.warning(f"No se pudo importar el módulo actions.sharepoint: {e}")
    sharepoint_loaded = False
except AttributeError as e:
    logger.warning(f"Error de atributo importando desde actions.sharepoint: {e}. Verifica que todas las funciones existan.")
    sharepoint_loaded = False


try:
    # Teams
    from actions.teams import (
        listar_chats as team_listar_chats, obtener_chat as team_obtener_chat,
        crear_chat as team_crear_chat, enviar_mensaje_chat as team_enviar_mensaje_chat,
        obtener_mensajes_chat as team_obtener_mensajes_chat, actualizar_mensaje_chat as team_actualizar_mensaje_chat,
        eliminar_mensaje_chat as team_eliminar_mensaje_chat, listar_equipos as team_listar_equipos,
        obtener_equipo as team_obtener_equipo, crear_equipo as team_crear_equipo,
        archivar_equipo as team_archivar_equipo, unarchivar_equipo as team_unarchivar_equipo,
        eliminar_equipo as team_eliminar_equipo, listar_canales as team_listar_canales,
        obtener_canal as team_obtener_canal, crear_canal as team_crear_canal,
        actualizar_canal as team_actualizar_canal, eliminar_canal as team_eliminar_canal,
        enviar_mensaje_canal as team_enviar_mensaje_canal
    )
    teams_loaded = True
except ImportError as e:
    logger.warning(f"No se pudo importar el módulo actions.teams: {e}")
    teams_loaded = False

try:
    # Office (Word/Excel)
    from actions.office import (
        crear_documento_word, insertar_texto_word, obtener_documento_word,
        crear_excel, escribir_celda_excel, leer_celda_excel,
        crear_tabla_excel, agregar_datos_tabla_excel
    )
    office_loaded = True
except ImportError as e:
    logger.warning(f"No se pudo importar el módulo actions.office: {e}")
    office_loaded = False

try:
    # Planner & ToDo
    # CORRECCIÓN: Se usan los nombres sugeridos por mypy
    from actions.planner_todo import (
        listar_planes, obtener_plan, crear_plan, actualizar_plan, eliminar_plan,
        listar_tareas_planner, crear_tarea_planner, actualizar_tarea_planner, eliminar_tarea_planner,
        listar_listas_todo,             # CORREGIDO: listar_listas -> listar_listas_todo
        crear_lista_todo,               # CORREGIDO: crear_lista -> crear_lista_todo
        actualizar_lista_todo,          # CORREGIDO: actualizar_lista -> actualizar_lista_todo
        eliminar_lista_todo,            # CORREGIDO: eliminar_lista -> eliminar_lista_todo
        listar_tareas_todo,             # CORREGIDO: listar_tareas -> listar_tareas_todo
        crear_tarea_todo,               # CORREGIDO: crear_tarea -> crear_tarea_todo
        actualizar_tarea_todo,          # CORREGIDO: actualizar_tarea -> actualizar_tarea_todo
        eliminar_tarea_todo,            # CORREGIDO: eliminar_tarea -> eliminar_tarea_todo
        completar_tarea_todo            # CORREGIDO: completar_tarea -> completar_tarea_todo
    )
    planner_todo_loaded = True
except ImportError as e:
    logger.warning(f"No se pudo importar el módulo actions.planner_todo: {e}")
    planner_todo_loaded = False
except AttributeError as e:
    logger.warning(f"Error de atributo importando desde actions.planner_todo: {e}. Verifica que todas las funciones existan.")
    planner_todo_loaded = False


try:
    # Power Automate
    from actions.power_automate import (
        listar_flows, obtener_flow, crear_flow, actualizar_flow,
        eliminar_flow, ejecutar_flow, obtener_estado_ejecucion_flow
    )
    power_automate_loaded = True
except ImportError as e:
    logger.warning(f"No se pudo importar el módulo actions.power_automate: {e}")
    power_automate_loaded = False

try:
    # Power BI
    # CORRECCIÓN: Se usan los nombres sugeridos por mypy
    from actions.power_bi import (
        listar_workspaces, obtener_workspace, listar_dashboards,
        obtener_dashboard, listar_reports, obtener_reporte,
        listar_datasets, obtener_dataset, refrescar_dataset,
        obtener_estado_refresco_dataset, # CORREGIDO: obtener_estado_refresco -> obtener_estado_refresco_dataset
        obtener_embed_url
    )
    power_bi_loaded = True
except ImportError as e:
    logger.warning(f"No se pudo importar el módulo actions.power_bi: {e}")
    power_bi_loaded = False
except AttributeError as e:
    logger.warning(f"Error de atributo importando desde actions.power_bi: {e}. Verifica que todas las funciones existan.")
    power_bi_loaded = False


# --- Diccionario de acciones disponibles ---
# Se construye dinámicamente basado en las importaciones exitosas
acciones_disponibles = {}

if correo_loaded:
    acciones_disponibles.update({
        "mail_listar": listar_correos,
        "mail_leer": leer_correo,
        "mail_enviar": enviar_correo,
        "mail_guardar_borrador": guardar_borrador,
        "mail_enviar_borrador": enviar_borrador,
        "mail_responder": responder_correo,
        "mail_reenviar": reenviar_correo,
        "mail_eliminar": eliminar_correo,
    })

if calendario_loaded:
    acciones_disponibles.update({
        "cal_listar_eventos": listar_eventos,
        "cal_crear_evento": crear_evento,
        "cal_actualizar_evento": actualizar_evento,
        "cal_eliminar_evento": eliminar_evento,
        "cal_crear_reunion_teams": crear_reunion_teams,
    })

if onedrive_loaded:
    acciones_disponibles.update({
        "od_listar_archivos": od_listar_archivos,
        "od_subir_archivo": od_subir_archivo,
        "od_descargar_archivo": od_descargar_archivo,
        "od_eliminar_archivo": od_eliminar_archivo,
        "od_crear_carpeta": od_crear_carpeta,
        "od_mover_archivo": od_mover_archivo,
        "od_copiar_archivo": od_copiar_archivo,
        "od_obtener_metadatos_archivo": od_obtener_metadatos_archivo,
        "od_actualizar_metadatos_archivo": od_actualizar_metadatos_archivo,
    })

if sharepoint_loaded:
    acciones_disponibles.update({
        "sp_crear_lista": sp_crear_lista,
        "sp_listar_listas": sp_listar_listas,
        "sp_agregar_elemento_lista": sp_agregar_elemento_lista, # CORREGIDO
        "sp_listar_elementos_lista": sp_listar_elementos_lista, # CORREGIDO
        "sp_actualizar_elemento_lista": sp_actualizar_elemento_lista,
        "sp_eliminar_elemento_lista": sp_eliminar_elemento_lista,
        "sp_listar_documentos_biblioteca": sp_listar_documentos_biblioteca,
        "sp_subir_documento": sp_subir_documento,
        # "sp_eliminar_archivo_biblioteca": sp_eliminar_archivo_biblioteca, # CORREGIDO (Verificar existencia)
        "sp_crear_carpeta_biblioteca": sp_crear_carpeta_biblioteca,
        "sp_mover_archivo": sp_mover_archivo_sp, # CORREGIDO (Renombrado)
        "sp_copiar_archivo": sp_copiar_archivo_sp, # CORREGIDO (Renombrado)
        "sp_obtener_metadatos_archivo": sp_obtener_metadatos_archivo,
        "sp_actualizar_metadatos_archivo": sp_actualizar_metadatos_archivo,
        "sp_obtener_contenido_archivo": sp_obtener_contenido_archivo,
        "sp_actualizar_contenido_archivo": sp_actualizar_contenido_archivo,
        "sp_crear_enlace_compartido_archivo": sp_crear_enlace_compartido_archivo,
    })

if teams_loaded:
    acciones_disponibles.update({
        "team_listar_chats": team_listar_chats,
        "team_obtener_chat": team_obtener_chat,
        "team_crear_chat": team_crear_chat,
        "team_enviar_mensaje_chat": team_enviar_mensaje_chat,
        "team_obtener_mensajes_chat": team_obtener_mensajes_chat,
        "team_actualizar_mensaje_chat": team_actualizar_mensaje_chat,
        "team_eliminar_mensaje_chat": team_eliminar_mensaje_chat,
        "team_listar_equipos": team_listar_equipos,
        "team_obtener_equipo": team_obtener_equipo,
        "team_crear_equipo": team_crear_equipo,
        "team_archivar_equipo": team_archivar_equipo,
        "team_unarchivar_equipo": team_unarchivar_equipo,
        "team_eliminar_equipo": team_eliminar_equipo,
        "team_listar_canales": team_listar_canales,
        "team_obtener_canal": team_obtener_canal,
        "team_crear_canal": team_crear_canal,
        "team_actualizar_canal": team_actualizar_canal,
        "team_eliminar_canal": team_eliminar_canal,
        "team_enviar_mensaje_canal": team_enviar_mensaje_canal,
    })

if office_loaded:
    acciones_disponibles.update({
        "office_crear_word": crear_documento_word,
        "office_insertar_texto_word": insertar_texto_word,
        "office_leer_word": obtener_documento_word,
        "office_crear_excel": crear_excel,
        "office_escribir_celda_excel": escribir_celda_excel,
        "office_leer_celda_excel": leer_celda_excel,
        "office_crear_tabla_excel": crear_tabla_excel,
        "office_agregar_datos_tabla_excel": agregar_datos_tabla_excel,
    })

if planner_todo_loaded:
    acciones_disponibles.update({
        "planner_listar_planes": listar_planes,
        "planner_obtener_plan": obtener_plan,
        "planner_crear_plan": crear_plan,
        "planner_actualizar_plan": actualizar_plan,
        "planner_eliminar_plan": eliminar_plan,
        "planner_listar_tareas": listar_tareas_planner,
        "planner_crear_tarea": crear_tarea_planner,
        "planner_actualizar_tarea": actualizar_tarea_planner,
        "planner_eliminar_tarea": eliminar_tarea_planner,
        "todo_listar_listas": listar_listas_todo, # CORREGIDO
        "todo_crear_lista": crear_lista_todo, # CORREGIDO
        "todo_actualizar_lista": actualizar_lista_todo, # CORREGIDO
        "todo_eliminar_lista": eliminar_lista_todo, # CORREGIDO
        "todo_listar_tareas": listar_tareas_todo, # CORREGIDO
        "todo_crear_tarea": crear_tarea_todo, # CORREGIDO
        "todo_actualizar_tarea": actualizar_tarea_todo, # CORREGIDO
        "todo_eliminar_tarea": eliminar_tarea_todo, # CORREGIDO
        "todo_completar_tarea": completar_tarea_todo, # CORREGIDO
    })

if power_automate_loaded:
    acciones_disponibles.update({
        "flow_listar": listar_flows,
        "flow_obtener": obtener_flow,
        "flow_crear": crear_flow,
        "flow_actualizar": actualizar_flow,
        "flow_eliminar": eliminar_flow,
        "flow_ejecutar": ejecutar_flow,
        "flow_obtener_estado_ejecucion": obtener_estado_ejecucion_flow,
    })

if power_bi_loaded:
    acciones_disponibles.update({
        "pbi_listar_workspaces": listar_workspaces,
        "pbi_obtener_workspace": obtener_workspace,
        "pbi_listar_dashboards": listar_dashboards,
        "pbi_obtener_dashboard": obtener_dashboard,
        "pbi_listar_reports": listar_reports,
        "pbi_obtener_reporte": obtener_reporte,
        "pbi_listar_datasets": listar_datasets,
        "pbi_obtener_dataset": obtener_dataset,
        "pbi_refrescar_dataset": refrescar_dataset,
        "pbi_obtener_estado_refresco": obtener_estado_refresco_dataset, # CORREGIDO
        "pbi_obtener_embed_url": obtener_embed_url,
    })

logger.info(f"Acciones disponibles cargadas: {list(acciones_disponibles.keys())}")

# Validar que el diccionario no esté vacío si se esperaba cargar acciones
if not acciones_disponibles:
    logger.error("¡Advertencia! No se cargó ninguna acción. Verifica los imports y los módulos en la carpeta 'actions'.")

