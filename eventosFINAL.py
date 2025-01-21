import wmi
import csv
import time
import psutil
import smtplib
from collections import Counter
from datetime import datetime, timedelta, timezone
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configuración de SharePoint VISOR DE EVENTOS
site_url = 'https://idtsas.sharepoint.com/sites/Bk-Empresas'
list_name = 'Visor de Eventos Construsol'
username = 'info@idtsas.com'
password = '1D2022++'

# Configuración de SharePoint CHEQUEO SERVIDOR
list_name_inventario = 'Inventario PC Construsol'  # Lista relacionada para la columna Lookup
list_name_chequeo_servidor = 'Chequeo Servidor Construsol'
# nombre_equipo_actual = socket.gethostname()
nombre_equipo_actual = "CONSTRUSOL002"  # Nombre de ejemplo
ruta_archivos_guardados = "C:\\Archivos\\Backup"  # Ajustar ruta según tu caso
ruta_archivos_guardados_system_state = "C:\\Archivos\\SystemState"  # Ajustar ruta

# Configuración de correo
smtp_server = "smtp.office365.com"  # Reemplaza con tu servidor SMTP
smtp_port = 587  # Puerto SMTP
email_sender = username  # Dirección de correo del remitente
email_password = password  # Contraseña del correo
email_recipient = "cmonsalve@idtsas.com"  # Dirección de correo del destinatario

# Funciones para el VISOR DE EVENTOS

def format_wmi_time(wmi_time):
    try:
        utc_time = datetime.strptime(wmi_time.split('.')[0], '%Y%m%d%H%M%S').replace(tzinfo=timezone.utc)
        bogota_time = utc_time.astimezone(timezone(timedelta(hours=-5)))
        return bogota_time.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return "N/A"

def get_events_from_yesterday():
    bogota_tz = timezone(timedelta(hours=-5))
    today_local = datetime.now(bogota_tz)
    yesterday_start_local = (today_local - timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    yesterday_end_local = (today_local - timedelta(days=1)).replace(hour=23, minute=59, second=59, microsecond=0)

    yesterday_start_utc = yesterday_start_local.astimezone(timezone.utc).strftime('%Y%m%d%H%M%S.000000+000')
    yesterday_end_utc = yesterday_end_local.astimezone(timezone.utc).strftime('%Y%m%d%H%M%S.999999+000')

    print(f"Buscando eventos entre {yesterday_start_utc} y {yesterday_end_utc} (UTC)")

    wmi_o = wmi.WMI('.')
    query = (
        f"SELECT * FROM Win32_NTLogEvent WHERE Logfile='System' "
        f"AND TimeGenerated >= '{yesterday_start_utc}' AND TimeGenerated <= '{yesterday_end_utc}'"
    )

    try:
        results = wmi_o.query(query)
        if not results:
            print("No se encontraron eventos para el día anterior.")
            return []
        return results

    except Exception as e:
        print(f"Error al ejecutar la consulta: {e}")
        return []

def save_events_to_csv(events, csv_file_name):
    headers = ["Nombre de Registro", "Origen", "ID", "Nivel", "Categoria de Tarea", "Registrado", "Equipo", "Usuario", "Mensaje", "EventRecordID"]
    try:
        with open(csv_file_name, mode='w', newline='', encoding='utf-8') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=headers)
            writer.writeheader()

            for event in events:
                full_message = event.Message or ""
                if hasattr(event, "StringInserts") and event.StringInserts:
                    full_message += "\n" + "\n".join(event.StringInserts)

                formatted_message = " ".join(full_message.splitlines())
                task_category = "Ninguno" if event.Category == 0 else getattr(event, "CategoryString", None) or event.Category

                row = {
                    "Nombre de Registro": event.Logfile,
                    "Origen": event.SourceName,
                    "ID": event.EventCode,
                    "Nivel": event.Type,
                    "Categoria de Tarea": task_category,
                    "Registrado": format_wmi_time(event.TimeGenerated),
                    "Equipo": event.ComputerName,
                    "Usuario": event.User,
                    "Mensaje": formatted_message,
                    "EventRecordID": event.RecordNumber
                }
                writer.writerow(row)

            print(f"Eventos guardados en el archivo CSV: {csv_file_name}")
    except Exception as e:
        print(f"Error al guardar en el archivo CSV: {e}")

def consolidate_events(events):
    consolidated = {}
    for event in events:
        key = (event.SourceName, event.EventCode, getattr(event, "CategoryString", "Ninguno"))
        full_message = event.Message or ""
        if hasattr(event, "StringInserts") and event.StringInserts:
            full_message += "\n" + "\n".join(event.StringInserts)

        formatted_message = full_message.strip()
        if key not in consolidated:
            consolidated[key] = {
                "SourceName": event.SourceName,
                "EventCode": event.EventCode,
                "CategoryString": getattr(event, "CategoryString", "Ninguno"),
                "Message": formatted_message,
                "TimeGenerated": format_wmi_time(event.TimeGenerated),
                "Type": event.Type,
                "User": event.User,
                "ComputerName": event.ComputerName,
                "NoEventos": 1
            }
        else:
            consolidated[key]["NoEventos"] += 1
    return consolidated.values()

def upload_events_to_sharepoint(events):
    try:
        ctx_auth = AuthenticationContext(site_url)
        if ctx_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(site_url, ctx_auth)
            print("Autenticación exitosa")
        else:
            print("Error al autenticar:", ctx_auth.get_last_error())
            return

        list_obj = ctx.web.lists.get_by_title(list_name)
        event_count = 0

        for event in events:
            try:
                item_properties = {
                    'Title': event['SourceName'],
                    'Id_x0020_del_x0020_Evento': str(event['EventCode']),
                    'Origen': event['SourceName'],
                    'Categoria_x0020_de_x0020_Tarea': event['CategoryString'],
                    'Detalle': event['Message'],
                    'Fecha': event['TimeGenerated'],
                    'Nivel': event['Type'],
                    'User': event['User'],
                    'No_x0020_de_x0020_Eventos': event['NoEventos']
                }
                list_obj.add_item(item_properties)
                ctx.execute_query()
                event_count += 1
                time.sleep(0.5)
            except Exception as e:
                print(f"Error al cargar el evento {event['EventCode']}: {e}")

        print(f"Total de eventos cargados a SharePoint: {event_count}")
        return event_count

    except Exception as e:
        print(f"Error al subir los eventos a SharePoint: {e}")
        return 0

def send_email_notification(event_count):
    try:
        msg = MIMEMultipart()
        msg['From'] = email_sender
        msg['To'] = email_recipient
        msg['Subject'] = f"Subida de eventos completada ({event_count} eventos)"

        body = f"Se han subido correctamente {event_count} eventos al portal de SharePoint para el día anterior.\nPor favor, revise los eventos subidos en el portal."
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(email_sender, email_password)
        server.sendmail(email_sender, email_recipient, msg.as_string())
        server.quit()

        print(f"Correo de notificación enviado a {email_recipient}")

    except Exception as e:
        print(f"Error al enviar el correo de notificación: {e}")

# FUNCIONES PARA EL CHEQUEO SERVIDOR

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
def subir_chequeo_servidor_sharepoint(datos, servidor_lookup_id):
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
        print("Chequeo Servidor subido exitosamente a SharePoint.")
    else:
        print("Error de autenticación.")

if __name__ == "__main__":
    csv_file_name = "Eventos_ayer.csv"
    events = get_events_from_yesterday()
    
    # CHEQUEO SERVIDOR 
    datos_sistema = obtener_datos_sistema()
    print(f"Nombre del equipo actual: {nombre_equipo_actual}")
    servidor_lookup_id = obtener_servidor_id(nombre_equipo_actual)
    # Subir datos a SharePoint
    if servidor_lookup_id:
        # Subir datos a SharePoint
        subir_chequeo_servidor_sharepoint(datos_sistema, servidor_lookup_id)
    else:
        print("No se pudo encontrar el equipo en la lista. No se subieron datos.")
    
    if events:
        save_events_to_csv(events, csv_file_name)
        consolidated_events = consolidate_events(events)
        event_count = upload_events_to_sharepoint(consolidated_events)
        if event_count > 0:
            send_email_notification(event_count)
    else:
        print("No se encontraron eventos para el día anterior.")
