import psutil
from datetime import datetime
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
import socket  # Para obtener el nombre del equipo

# Configuración de SharePoint
site_url = 'https://idtsas.sharepoint.com/sites/Bk-Empresas'
list_name_chequeo_servidor = 'Chequeo Servidor Construsol'
list_name_inventario = 'Inventario PC Construsol'  # Lista relacionada para la columna Lookup
username = 'info@idtsas.com'
password = '1D2022++'
ruta_archivos_guardados = "C:\\Archivos\\Backup"
ruta_archivos_guardados_system_state = "C:\\Archivos\\SystemState"

# Obtener información del sistema
def obtener_datos_sistema():
    # Procesador y memoria
    cpu_usage = psutil.cpu_percent(interval=1)
    memory_info = psutil.virtual_memory()
    memory_usage = memory_info.percent

    # Información de los discos
    discos = {}
    for partition in psutil.disk_partitions():
        try:
            disk_usage = psutil.disk_usage(partition.mountpoint)
            disco_nombre = partition.device.strip("\\")  # Ejemplo: C:, D:, sin barra
            discos[disco_nombre] = round(disk_usage.free / (1024**3), 2)  # Espacio libre en GB
        except PermissionError:
            discos[partition.device] = 0  # Si no hay acceso, se registra como 0

    # Combinar datos en un diccionario
    datos = {
        "Ruta Archivos Guardados": ruta_archivos_guardados,
        "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Ruta Archivos Guardados System State": ruta_archivos_guardados_system_state,
        "Procesador (% usado)": cpu_usage,
        "Memoria (% usado)": memory_usage,
        "Tamaño Discos": round(sum(psutil.disk_usage(p.mountpoint).total for p in psutil.disk_partitions()) / (1024**3), 2),
        "Espacio Libre Disco C": discos.get("C:", 0),
        "Espacio Libre Disco D": discos.get("D:", 0),
        "Espacio Libre Disco I": discos.get("I:", 0),
        "Version de la Actualizacion": "1.0.0",
        "Virus Detectados": 0,
        "Sistema Operativo": f"{psutil.os.name}",
    }

    return datos

# Obtener el ID del equipo en la lista "Inventario PC Construsol"
def obtener_servidor_id(nombre_equipo):
    context_auth = AuthenticationContext(site_url)
    if context_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(site_url, context_auth)
        inventario_list = ctx.web.lists.get_by_title(list_name_inventario)
        items = inventario_list.get_items().execute_query()

        for item in items:
            if item.properties["Title"] == nombre_equipo:  # Comparar con el nombre del equipo
                return item.properties["ID"]

        print(f"No se encontró el equipo con nombre: {nombre_equipo} en la lista {list_name_inventario}.")
        return None
    else:
        print("Error de autenticación al obtener el ID del servidor.")
        return None

# Subir datos a SharePoint
def subir_a_sharepoint(datos, servidor_lookup_id):
    context_auth = AuthenticationContext(site_url)
    if context_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(site_url, context_auth)
        target_list = ctx.web.lists.get_by_title(list_name_chequeo_servidor)

        # Crear un nuevo elemento en la lista
        item_properties = {
            "Title": datos["Ruta Archivos Guardados"],
            "Fecha": datos["Fecha"],
            "Ruta_x0020_Archivos_x0020_Guarda": datos["Ruta Archivos Guardados System State"],
            "Procesador_x0020__x0028__x0025__": datos["Procesador (% usado)"] / 100,
            "Memoria_x0020__x0028__x0025__x00": datos["Memoria (% usado)"] / 100,
            "Tama_x00f1_o_x0020_Discos": datos["Tamaño Discos"],
            "Espacio_x0020_Libre_x0020_Disco_": datos["Espacio Libre Disco C"],
            "Espacio_x0020_Libre_x0020_Disco_0": datos["Espacio Libre Disco D"],
            "Espacio_x0020_Libre_x0020_Disco_1": datos["Espacio Libre Disco I"],
            "Version_x0020_de_x0020_la_x0020_": datos["Version de la Actualizacion"],
            "Virus_x0020_Detectados": datos["Virus Detectados"],
            "Sistemas_x0020_Operativo": datos["Sistema Operativo"],
            "ServidorId": servidor_lookup_id,  # Enviar el ID directamente como un número entero
        }
        target_list.add_item(item_properties).execute_query()
        print("Datos subidos exitosamente a SharePoint.")
    else:
        print("Error de autenticación.")

# Ejecución
if __name__ == "__main__":
    # Obtener datos del sistema
    datos_sistema = obtener_datos_sistema()

    # Obtener el nombre del equipo actual
    # nombre_equipo_actual = socket.gethostname()
    nombre_equipo_actual = "CONTADOR"  # Nombre de ejemplo
    print(f"Nombre del equipo actual: {nombre_equipo_actual}")

    # Buscar el ID del equipo en la lista "Inventario PC Construsol"
    servidor_lookup_id = obtener_servidor_id(nombre_equipo_actual)

    if servidor_lookup_id:
        # Subir datos a SharePoint
        subir_a_sharepoint(datos_sistema, servidor_lookup_id)
    else:
        print("No se pudo encontrar el equipo en la lista. No se subieron datos.")
