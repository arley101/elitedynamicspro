"""
mapeo_acciones.py

Este archivo centraliza todas las acciones disponibles en el sistema
y las mapea a las funciones que las ejecutan. Las acciones están
organizadas por categorías para facilitar su mantenimiento.
"""

# --- Importaciones de funciones por categoría ---

# Correo
from actions.correo import (
    listar_correos, leer_correo, enviar_correo, guardar_borrador, enviar_borrador,
    responder_correo, reenviar_correo, eliminar_correo
)

# Calendario
from actions.calendario import (
    listar_eventos, crear_evento, actualizar_evento, eliminar_evento, crear_reunion_teams
)

# OneDrive
from actions.onedrive import (
    listar_archivos as od_listar_archivos, subir_archivo as od_subir_archivo,
    descargar_archivo as od_descargar_archivo, eliminar_archivo as od_eliminar_archivo,
    crear_carpeta as od_crear_carpeta, mover_archivo as od_mover_archivo,
    copiar_archivo as od_copiar_archivo, obtener_metadatos_archivo as od_obtener_metadatos_archivo,
    actualizar_metadatos_archivo as od_actualizar_metadatos_archivo
)

# SharePoint
from actions.sharepoint import (
    crear_lista as sp_crear_lista, listar_listas as sp_listar_listas,
    agregar_elemento as sp_agregar_elemento_lista, listar_elementos as sp_listar_elementos_lista,
    actualizar_elemento_lista as sp_actualizar_elemento_lista, eliminar_elemento_lista as sp_eliminar_elemento_lista,
    listar_documentos_biblioteca as sp_listar_documentos_biblioteca, subir_documento as sp_subir_documento,
    eliminar_archivo_biblioteca as sp_eliminar_archivo_biblioteca, crear_carpeta_biblioteca as sp_crear_carpeta_biblioteca,
    mover_archivo as sp_mover_archivo, copiar_archivo as sp_copiar_archivo,
    obtener_metadatos_archivo as sp_obtener_metadatos_archivo, actualizar_metadatos_archivo as sp_actualizar_metadatos_archivo,
    obtener_contenido_archivo as sp_obtener_contenido_archivo, actualizar_contenido_archivo as sp_actualizar_contenido_archivo,
    crear_enlace_compartido_archivo as sp_crear_enlace_compartido_archivo
)

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

# Office (Word/Excel)
from actions.office import (
    crear_documento_word, insertar_texto_word, obtener_documento_word,
    crear_excel, escribir_celda_excel, leer_celda_excel,
    crear_tabla_excel, agregar_datos_tabla_excel
)

# Planner & ToDo
from actions.planner_todo import (
    listar_planes, obtener_plan, crear_plan, actualizar_plan, eliminar_plan,
    listar_tareas_planner, crear_tarea_planner, actualizar_tarea_planner, eliminar_tarea_planner,
    listar_listas as listar_listas_todo, crear_lista as crear_lista_todo,
    actualizar_lista as actualizar_lista_todo, eliminar_lista as eliminar_lista_todo,
    listar_tareas as listar_tareas_todo, crear_tarea as crear_tarea_todo,
    actualizar_tarea as actualizar_tarea_todo, eliminar_tarea as eliminar_tarea_todo,
    completar_tarea as completar_tarea_todo
)

# Power Automate
from actions.power_automate import (
    listar_flows, obtener_flow, crear_flow, actualizar_flow,
    eliminar_flow, ejecutar_flow, obtener_estado_ejecucion_flow
)

# Power BI
from actions.power_bi import (
    listar_workspaces, obtener_workspace, listar_dashboards,
    obtener_dashboard, listar_reports, obtener_reporte,
    listar_datasets, obtener_dataset, refrescar_dataset,
    obtener_estado_refresco as obtener_estado_refresco_dataset, obtener_embed_url
)

# --- Diccionario de acciones disponibles ---
acciones_disponibles = {
    # Correo
    "mail_listar": listar_correos,
    "mail_leer": leer_correo,
    "mail_enviar": enviar_correo,
    "mail_guardar_borrador": guardar_borrador,
    "mail_enviar_borrador": enviar_borrador,
    "mail_responder": responder_correo,
    "mail_reenviar": reenviar_correo,
    "mail_eliminar": eliminar_correo,
    # Calendario
    "cal_listar_eventos": listar_eventos,
    "cal_crear_evento": crear_evento,
    "cal_actualizar_evento": actualizar_evento,
    "cal_eliminar_evento": eliminar_evento,
    "cal_crear_reunion_teams": crear_reunion_teams,
    # OneDrive
    "od_listar_archivos": od_listar_archivos,
    "od_subir_archivo": od_subir_archivo,
    "od_descargar_archivo": od_descargar_archivo,
    "od_eliminar_archivo": od_eliminar_archivo,
    "od_crear_carpeta": od_crear_carpeta,
    "od_mover_archivo": od_mover_archivo,
    "od_copiar_archivo": od_copiar_archivo,
    "od_obtener_metadatos_archivo": od_obtener_metadatos_archivo,
    "od_actualizar_metadatos_archivo": od_actualizar_metadatos_archivo,
    # SharePoint
    "sp_crear_lista": sp_crear_lista,
    "sp_listar_listas": sp_listar_listas,
    "sp_agregar_elemento_lista": sp_agregar_elemento_lista,
    "sp_listar_elementos_lista": sp_listar_elementos_lista,
    "sp_actualizar_elemento_lista": sp_actualizar_elemento_lista,
    "sp_eliminar_elemento_lista": sp_eliminar_elemento_lista,
    "sp_listar_documentos_biblioteca": sp_listar_documentos_biblioteca,
    "sp_subir_documento": sp_subir_documento,
    "sp_eliminar_archivo_biblioteca": sp_eliminar_archivo_biblioteca,
    "sp_crear_carpeta_biblioteca": sp_crear_carpeta_biblioteca,
    "sp_mover_archivo": sp_mover_archivo,
    "sp_copiar_archivo": sp_copiar_archivo,
    "sp_obtener_metadatos_archivo": sp_obtener_metadatos_archivo,
    "sp_actualizar_metadatos_archivo": sp_actualizar_metadatos_archivo,
    "sp_obtener_contenido_archivo": sp_obtener_contenido_archivo,
    "sp_actualizar_contenido_archivo": sp_actualizar_contenido_archivo,
    "sp_crear_enlace_compartido_archivo": sp_crear_enlace_compartido_archivo,
    # Teams
    "team_listar_chats": team_listar_chats,
    "team_obtener_chat": team_obtener_chat,
    "team_crear_chat": team_crear_chat,
    "team_enviar_mensaje_chat": team_enviar_mensaje_chat,
    # Office
    "office_crear_word": crear_documento_word,
    "office_insertar_texto_word": insertar_texto_word,
    "office_leer_word": obtener_documento_word,
}
