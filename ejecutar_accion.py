import os
import json
from acciones import correo  # Asegúrate de que la carpeta 'acciones' esté bien ubicada

# Configuración de variables de entorno (puedes ponerlas aquí directamente para probar)
CLIENT_ID = os.getenv('CLIENT_ID', 'tu-client-id')
TENANT_ID = os.getenv('TENANT_ID', 'tu-tenant-id')
CLIENT_SECRET = os.getenv('CLIENT_SECRET', 'tu-client-secret')
GRAPH_SCOPE = os.getenv('GRAPH_SCOPE', 'https://graph.microsoft.com/.default')

# Verificación rápida
if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET, GRAPH_SCOPE]):
    print("❌ Faltan variables de entorno necesarias.")
    exit(1)

# Acción a ejecutar (simulación del JSON que enviabas)
peticion_json = {
    "accion": "listar_correos",
    "parametros": {"top": 1}
}

# Ejecutar acción
if peticion_json["accion"] == "listar_correos":
    resultado = correo.listar_correos(top=peticion_json["parametros"]["top"])
    print("✅ Resultado de la acción:")
    print(json.dumps(resultado, indent=4))
else:
    print("⚠️ Acción no reconocida.")