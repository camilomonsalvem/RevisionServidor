import psutil
from datetime import datetime
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

# Configuración de SharePoint (usar variables de entorno para mayor seguridad)
site_url = 'https://idtsas.sharepoint.com/sites/Bk-Empresas'
list_name_chequeo_servidor = 'Chequeo Servidor Construsol'
username = 'info@idtsas.com'
password = '1D2022++'
ruta_archivos_guardados = "C:\\Archivos\\Backup"  # Ajustar ruta según tu caso
ruta_archivos_guardados_system_state = "C:\\Archivos\\SystemState"  # Ajustar ruta

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
        "Ruta Archivos Guardados": ruta_archivos_guardados,  # Ajustar ruta según tu caso
        "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # Fecha y hora actual
        "Ruta Archivos Guardados System State": ruta_archivos_guardados_system_state,  # Ajustar ruta
        "Procesador (% usado)": cpu_usage,
        "Memoria (% usado)": memory_usage,
        "Tamaño Discos": round(sum(psutil.disk_usage(p.mountpoint).total for p in psutil.disk_partitions()) / (1024**3), 2),
        "Espacio Libre Disco C": discos.get("C:", 0),
        "Espacio Libre Disco D": discos.get("D:", 0),
        "Espacio Libre Disco I": discos.get("I:", 0),
        "Version de la Actualizacion": "1.0.0",  # Ajustar según corresponda
        "Virus Detectados": 0,  # Suponiendo 0 para este caso
        "Sistema Operativo": f"{psutil.os.name}",  # Nombre del sistema operativo
    }

    return datos

# Subir datos a SharePoint
def subir_a_sharepoint(datos):
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
            "Tama_x00f1_o_x0020_Discos": 1862,
            "Espacio_x0020_Libre_x0020_Disco_": datos["Espacio Libre Disco C"],
            "Espacio_x0020_Libre_x0020_Disco_0": datos["Espacio Libre Disco D"],
            "Espacio_x0020_Libre_x0020_Disco_1": datos["Espacio Libre Disco I"],
            "Version_x0020_de_x0020_la_x0020_": datos["Version de la Actualizacion"],
            "Virus_x0020_Detectados": datos["Virus Detectados"],
            "Sistemas_x0020_Operativo": datos["Sistema Operativo"],
        }
        target_list.add_item(item_properties).execute_query()
        print("Datos subidos exitosamente a SharePoint.")
    else:
        print("Error de autenticación.")

# Ejecución
if __name__ == "__main__":
    # Obtener datos del sistema
    datos_sistema = obtener_datos_sistema()

    # Subir datos a SharePoint
    subir_a_sharepoint(datos_sistema)
