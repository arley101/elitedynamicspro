# actions/office.py (Refactorizado)

import logging
import requests # Solo para tipos de excepción
import json
import os
from typing import Dict, List, Optional, Union, Any

# Usar logger estándar de Azure Functions
logger = logging.getLogger("azure.functions")

# Importar helper y constantes desde la estructura compartida
try:
    # Asume que shared está un nivel arriba de actions
    from ..shared.helpers.http_client import hacer_llamada_api
    from ..shared.constants import BASE_URL, GRAPH_API_TIMEOUT
except ImportError as e:
    logging.critical(f"Error CRÍTICO importando helpers/constantes en Office: {e}. Verifica la estructura y PYTHONPATH.", exc_info=True)
    BASE_URL = "https://graph.microsoft.com/v1.0"; GRAPH_API_TIMEOUT = 45
    def hacer_llamada_api(*args, **kwargs):
        raise NotImplementedError("Dependencia 'hacer_llamada_api' no importada correctamente.")

# ---- FUNCIONES DE WORD ONLINE (via OneDrive /me/drive) ----
# Todas usan la firma (parametros: Dict[str, Any], headers: Dict[str, str])

def crear_documento_word(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea un nuevo documento de Word vacío en OneDrive.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo'.
                                     Opcional: 'ruta' (carpeta destino, default '/').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos del archivo Word creado.
    """
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo")
    ruta: str = parametros.get("ruta", "/") # Carpeta raíz por defecto

    if not nombre_archivo: raise ValueError("Parámetro 'nombre_archivo' es requerido.")
    # Asegurar extensión .docx
    if not nombre_archivo.lower().endswith(".docx"):
        nombre_archivo += ".docx"
        logger.debug(f"Añadida extensión .docx al nombre: {nombre_archivo}")

    # Construir path relativo al root de OneDrive del usuario
    target_folder_path = ruta.strip('/')
    # Asegurar que el path para Graph API no empiece con '/' si no es la raíz
    target_file_path = f"{nombre_archivo}" if not target_folder_path else f"{target_folder_path}/{nombre_archivo}"
    # El endpoint para crear/reemplazar por path es /root:/path/to/file.docx
    url = f"{BASE_URL}/me/drive/root:/{target_file_path}"

    # Headers y body para crear archivo vacío
    create_headers = headers.copy()
    # Tipo MIME correcto para .docx
    create_headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    # El body debe contener el nombre y un objeto 'file' vacío para creación
    # Si el archivo ya existe, PUT lo reemplazará (comportamiento por defecto)
    # Se podría añadir @microsoft.graph.conflictBehavior al endpoint si se quiere 'rename' o 'fail'
    body = {"name": nombre_archivo, "file": {}}

    logger.info(f"Creando/Reemplazando Word '{nombre_archivo}' en ruta '/{target_folder_path}' de OneDrive")
    # Usamos PUT para crear/reemplazar por path. El helper maneja json_data.
    # Graph API devuelve los metadatos del archivo creado/reemplazado.
    return hacer_llamada_api("PUT", url, create_headers, json_data=body, expect_json=True)


def insertar_texto_word(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    REEMPLAZA el contenido de un documento de Word con texto plano.
    ADVERTENCIA: Pierde todo el formato original. No es una inserción real.

    Args:
        parametros (Dict[str, Any]): Debe contener 'item_id' (ID del archivo Word), 'texto'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos del archivo actualizado.
    """
    item_id: Optional[str] = parametros.get("item_id")
    texto: Optional[str] = parametros.get("texto")

    if not item_id: raise ValueError("Parámetro 'item_id' es requerido.")
    if texto is None: raise ValueError("Parámetro 'texto' es requerido.") # Texto vacío es permitido

    # Endpoint para actualizar contenido
    url = f"{BASE_URL}/me/drive/items/{item_id}/content"
    update_headers = headers.copy()
    # Indicar que estamos enviando texto plano
    update_headers['Content-Type'] = 'text/plain'

    logger.warning(f"REEMPLAZANDO contenido del Word con ID '{item_id}' con texto plano.")
    # Usamos PUT con el texto codificado en UTF-8 como 'data'
    # Aumentar timeout por si el texto es largo
    insert_timeout = max(GRAPH_API_TIMEOUT, 30) # Ej: 30 segundos mínimo
    return hacer_llamada_api(
        "PUT",
        url,
        update_headers,
        data=texto.encode('utf-8'), # Pasar bytes
        timeout=insert_timeout,
        expect_json=True # PUT en /content devuelve metadatos
    )


def obtener_documento_word(parametros: Dict[str, Any], headers: Dict[str, str]) -> bytes:
    """
    Obtiene el contenido binario (.docx) de un documento de Word.

    Args:
        parametros (Dict[str, Any]): Debe contener 'item_id'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        bytes: El contenido binario del archivo .docx.
    """
    item_id: Optional[str] = parametros.get("item_id")
    if not item_id: raise ValueError("Parámetro 'item_id' es requerido.")

    # Endpoint para obtener contenido
    url = f"{BASE_URL}/me/drive/items/{item_id}/content"
    logger.info(f"Obteniendo contenido binario del Word ID '{item_id}'")

    # Necesitamos la respuesta cruda, usamos el helper con expect_json=False
    download_timeout = max(GRAPH_API_TIMEOUT, 60) # Timeout más largo para descarga
    response = hacer_llamada_api("GET", url, headers, timeout=download_timeout, expect_json=False)

    if isinstance(response, requests.Response):
        logger.info(f"Contenido Word ID '{item_id}' obtenido ({len(response.content)} bytes).")
        return response.content
    else:
        logger.error(f"Respuesta inesperada del helper al obtener contenido Word: {type(response)}")
        raise Exception("Error interno al obtener contenido del archivo Word.")


# ---- FUNCIONES DE EXCEL ONLINE (via OneDrive /me/drive) ----

def crear_excel(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea un nuevo libro de Excel vacío en OneDrive.

    Args:
        parametros (Dict[str, Any]): Debe contener 'nombre_archivo'.
                                     Opcional: 'ruta' (carpeta destino, default '/').
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Metadatos del archivo Excel creado.
    """
    nombre_archivo: Optional[str] = parametros.get("nombre_archivo")
    ruta: str = parametros.get("ruta", "/")

    if not nombre_archivo: raise ValueError("Parámetro 'nombre_archivo' es requerido.")
    # Asegurar extensión .xlsx
    if not nombre_archivo.lower().endswith(".xlsx"):
        nombre_archivo += ".xlsx"
        logger.debug(f"Añadida extensión .xlsx al nombre: {nombre_archivo}")

    # Construir path relativo al root de OneDrive
    target_folder_path = ruta.strip('/')
    target_file_path = f"{nombre_archivo}" if not target_folder_path else f"{target_folder_path}/{nombre_archivo}"
    url = f"{BASE_URL}/me/drive/root:/{target_file_path}"

    # Headers y body para crear archivo vacío
    create_headers = headers.copy()
    # Tipo MIME correcto para .xlsx
    create_headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    body = {"name": nombre_archivo, "file": {}}

    logger.info(f"Creando/Reemplazando Excel '{nombre_archivo}' en ruta '/{target_folder_path}' de OneDrive")
    # Usamos PUT para crear/reemplazar por path
    return hacer_llamada_api("PUT", url, create_headers, json_data=body, expect_json=True)


def escribir_celda_excel(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Escribe un valor en una celda específica de una hoja de Excel.

    Args:
        parametros (Dict[str, Any]): Debe contener 'item_id', 'hoja', 'celda', 'valor'.
                                     'valor' puede ser str, int, float, bool.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Respuesta de Graph API (usualmente info del rango actualizado).
    """
    item_id: Optional[str] = parametros.get("item_id")
    hoja: Optional[str] = parametros.get("hoja") # Nombre o ID de la hoja
    celda: Optional[str] = parametros.get("celda") # Notación A1 (ej. "A1", "C5")
    valor: Any = parametros.get("valor") # Valor a escribir

    if not item_id: raise ValueError("Parámetro 'item_id' es requerido.")
    if not hoja: raise ValueError("Parámetro 'hoja' (nombre o ID) es requerido.")
    if not celda: raise ValueError("Parámetro 'celda' (ej. 'A1') es requerido.")
    if valor is None: raise ValueError("Parámetro 'valor' es requerido.")
    # Validar tipo de valor? Graph API maneja str, num, bool.
    if not isinstance(valor, (str, int, float, bool)):
        logger.warning(f"Escribiendo tipo no estándar '{type(valor)}' en celda. Se convertirá a string.")
        valor = str(valor)

    # Endpoint para el rango específico
    # Usar comillas simples alrededor de la dirección en range()
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/range(address='{celda}')"
    # El cuerpo debe ser un objeto con 'values', que es una lista de listas (matriz)
    body = {"values": [[valor]]} # Para una sola celda, es una matriz 1x1

    logger.info(f"Escribiendo valor '{valor}' en celda '{celda}', hoja '{hoja}', item Excel '{item_id}'")
    # Usamos PATCH para actualizar el rango
    return hacer_llamada_api("PATCH", url, headers, json_data=body)


def leer_celda_excel(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Lee el valor, texto y dirección de una celda específica.

    Args:
        parametros (Dict[str, Any]): Debe contener 'item_id', 'hoja', 'celda'.
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Información del rango (incluye 'text', 'values', 'address').
    """
    item_id: Optional[str] = parametros.get("item_id")
    hoja: Optional[str] = parametros.get("hoja")
    celda: Optional[str] = parametros.get("celda")

    if not item_id: raise ValueError("Parámetro 'item_id' es requerido.")
    if not hoja: raise ValueError("Parámetro 'hoja' es requerido.")
    if not celda: raise ValueError("Parámetro 'celda' es requerido.")

    # Endpoint del rango, seleccionando campos útiles
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/range(address='{celda}')?$select=text,values,address,formulas"
    logger.info(f"Leyendo celda '{celda}', hoja '{hoja}', item Excel '{item_id}'")
    # Usamos GET
    return hacer_llamada_api("GET", url, headers)


def crear_tabla_excel(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Crea una tabla de Excel sobre un rango existente.

    Args:
        parametros (Dict[str, Any]): Debe contener 'item_id', 'hoja', 'rango'.
                                     'rango' debe ser notación A1 (ej. "A1:C5").
                                     Opcional: 'tiene_headers' (bool, default False).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Información de la tabla creada.
    """
    item_id: Optional[str] = parametros.get("item_id")
    hoja: Optional[str] = parametros.get("hoja")
    rango: Optional[str] = parametros.get("rango") # Ej. "A1:D10"
    tiene_headers: bool = parametros.get("tiene_headers", False)

    if not item_id: raise ValueError("Parámetro 'item_id' es requerido.")
    if not hoja: raise ValueError("Parámetro 'hoja' es requerido.")
    if not rango or ':' not in rango: # Validar formato básico del rango
        raise ValueError("Parámetro 'rango' (ej. 'A1:C5') es requerido.")

    # Endpoint para añadir tablas
    url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/tables/add"
    # El cuerpo requiere la dirección completa (Hoja!Rango)
    body = {"address": f"{hoja}!{rango}", "hasHeaders": tiene_headers}

    logger.info(f"Creando tabla Excel en rango '{rango}', hoja '{hoja}', item '{item_id}'")
    # Usamos POST
    return hacer_llamada_api("POST", url, headers, json_data=body)


def agregar_datos_tabla_excel(parametros: Dict[str, Any], headers: Dict[str, str]) -> Dict[str, Any]:
    """
    Agrega filas de datos al final de una tabla de Excel existente.

    Args:
        parametros (Dict[str, Any]): Debe contener 'item_id', 'tabla_id_o_nombre', 'valores'.
                                     'valores' debe ser una lista de listas (filas).
                                     Opcional: 'hoja' (necesario si se usa nombre de tabla en lugar de ID).
        headers (Dict[str, str]): Cabeceras con token.

    Returns:
        Dict[str, Any]: Información sobre las filas añadidas.
    """
    item_id: Optional[str] = parametros.get("item_id")
    tabla_id_o_nombre: Optional[str] = parametros.get("tabla_id_o_nombre")
    valores: Optional[List[List[Any]]] = parametros.get("valores")
    hoja: Optional[str] = parametros.get("hoja") # Necesario si tabla_id_o_nombre es un nombre

    if not item_id: raise ValueError("Parámetro 'item_id' es requerido.")
    if not tabla_id_o_nombre: raise ValueError("Parámetro 'tabla_id_o_nombre' es requerido.")
    if not valores or not isinstance(valores, list):
        raise ValueError("Parámetro 'valores' (lista de listas) es requerido.")
    # Validar que sea lista de listas? Podría ser complejo. Asumimos formato correcto.
    if not all(isinstance(row, list) for row in valores):
        raise ValueError("'valores' debe ser una lista de listas.")

    # Construir endpoint. Necesita la hoja si se usa nombre de tabla.
    if hoja:
        # /workbook/worksheets/{id|name}/tables/{id|name}/rows
        url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/worksheets/{hoja}/tables/{tabla_id_o_nombre}/rows"
    else:
        # /workbook/tables/{id|name}/rows (Funciona si tabla_id_o_nombre es ID único)
        logger.warning("Usando endpoint de tabla sin especificar hoja. Asegúrate que 'tabla_id_o_nombre' es un ID único o que Graph puede resolverlo.")
        url = f"{BASE_URL}/me/drive/items/{item_id}/workbook/tables/{tabla_id_o_nombre}/rows"

    # El cuerpo requiere 'values' con la matriz de datos
    body = {"values": valores}

    logger.info(f"Agregando {len(valores)} filas a tabla Excel '{tabla_id_o_nombre}', item '{item_id}'")
    # Usamos POST para añadir filas
    return hacer_llamada_api("POST", url, headers, json_data=body)

# --- FIN DEL MÓDULO actions/office.py ---

