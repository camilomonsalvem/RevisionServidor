import wmi
import csv
import time
import psutil
import smtplib
import os
import socket
from dotenv import load_dotenv
from collections import Counter
from datetime import datetime, timedelta, timezone
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

load_dotenv(".env")

username = os.getenv("idt_username")
password = os.getenv("idt_password")
site_url = os.getenv("site_url")

# Configuración de SharePoint VISOR DE EVENTOS
list_name_visor_eventos = os.getenv("list_name_visor_eventos")

# Configuración de SharePoint CHEQUEO SERVIDOR
list_name_inventario = os.getenv("list_name_inventario")
list_name_chequeo_servidor = os.getenv("list_name_chequeo_servidor")
nombre_equipo_actual = socket.gethostname()
ruta_archivos_guardados = os.getenv("ruta_archivos_guardados")
ruta_archivos_guardados_system_state = os.getenv("ruta_archivos_guardados_system_state")

# Configuración de correo
smtp_server = os.getenv("smtp_server")
smtp_port = os.getenv("smtp_port")
email_sender = os.getenv("email_sender")
email_password = os.getenv("email_password")
email_recipient = os.getenv("email_recipient")

# Funciones para el VISOR DE EVENTOS

def format_wmi_time(wmi_time):
    try:
        if not wmi_time:
            return "N/A"
        utc_time = datetime.strptime(wmi_time.split('.')[0], '%Y%m%d%H%M%S').replace(tzinfo=timezone.utc)
        bogota_time = utc_time.astimezone(timezone(timedelta(hours=-5)))
        return bogota_time.strftime('%Y-%m-%d %H:%M:%S')
    except Exception as e:
        print(f"Error al formatear el tiempo: {e}")
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
        key = (
            event.SourceName or "Desconocido",
            event.EventCode or 0,
            getattr(event, "CategoryString", "Ninguno")
        )
        full_message = (event.Message or "").strip()
        if hasattr(event, "StringInserts") and event.StringInserts:
            full_message += "\n" + "\n".join(event.StringInserts)

        if key not in consolidated:
            consolidated[key] = {
                "SourceName": event.SourceName or "Desconocido",
                "EventCode": event.EventCode or 0,
                "CategoryString": getattr(event, "CategoryString", "Ninguno"),
                "Message": full_message,
                "TimeGenerated": format_wmi_time(event.TimeGenerated),
                "Type": event.Type or "Desconocido",
                "User": event.User or "Desconocido",
                "ComputerName": event.ComputerName or "Desconocido",
                "NoEventos": 1
            }
        else:
            consolidated[key]["NoEventos"] += 1

    return list(consolidated.values())

def upload_events_to_sharepoint(events, chequeo_servidor_id):
    try:
        ctx_auth = AuthenticationContext(site_url)
        if ctx_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(site_url, ctx_auth)
            print("Autenticación exitosa")
        else:
            print("Error al autenticar:", ctx_auth.get_last_error())
            return

        list_obj = ctx.web.lists.get_by_title(list_name_visor_eventos)
        event_count = 0

        for event in events:
            try:
                item_properties = {
                    'Title': event['SourceName'] or "Desconocido",
                    'Id_x0020_del_x0020_Evento': str(event['EventCode'] or 0),
                    'Origen': event['SourceName'] or "Desconocido",
                    'Categoria_x0020_de_x0020_Tarea': event['CategoryString'] or "Ninguno",
                    'Detalle': event['Message'] or "N/A",
                    'Fecha': event['TimeGenerated'] or "N/A",
                    'Nivel': event['Type'] or "N/A",
                    'User': event['User'] or "N/A",
                    'No_x0020_de_x0020_Eventos': event['NoEventos'] or 0,
                    'ID_x0020_Chequeo_x0020_ServidorId': chequeo_servidor_id  # Nota: "_Id" al final
                }
                list_obj.add_item(item_properties)
                ctx.execute_query()
                event_count += 1
                time.sleep(0.5)
            except Exception as e:
                print(f"Error al cargar el evento {event.get('EventRecordID', 'Desconocido')}: {e}")

        print(f"Total de eventos cargados a SharePoint: {event_count}")
        return event_count

    except Exception as e:
        print(f"Error al subir los eventos a SharePoint: {e}")
        return 0

def send_email_notification(total_events, error_events, critical_events):
    try:
        with open("email_template.html", "r", encoding="utf-8") as file:
            html_template = file.read()

        html_body = html_template.replace("{{total_events}}", str(total_events))
        html_body = html_body.replace("{{error_events}}", str(error_events))
        html_body = html_body.replace("{{critical_events}}", str(critical_events))

        msg = MIMEMultipart("alternative")
        msg['From'] = email_sender
        msg['To'] = email_recipient
        msg['Subject'] = f"Reporte de Eventos: {total_events} procesados, {error_events} errores, {critical_events} críticos"
        msg.attach(MIMEText(html_body, "html"))

        server = smtplib.SMTP(smtp_server, int(smtp_port))
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
    total_disks = 0  # Tamaño total de los discos accesibles

    for partition in psutil.disk_partitions():
        try:
            # Filtrar solo discos duros locales (skip CD/DVD drives y unidades extraíbles)
            if "cdrom" in partition.opts or partition.fstype == "":
                print(f"Unidad omitida (no es disco local): {partition.device}")
                continue

            disk_usage = psutil.disk_usage(partition.mountpoint)
            disco_nombre = partition.device.strip("\\") if partition.device else "Desconocido"  # Validar None
            discos[disco_nombre] = round(disk_usage.free / (1024**3), 2)  # Espacio libre en GB
            total_disks += disk_usage.total  # Sumar el tamaño total de discos accesibles
        except PermissionError:
            print(f"Unidad no accesible: {partition.device}")
            discos[partition.device if partition.device else "Desconocido"] = 0
        except Exception as e:
            print(f"Error al acceder a la unidad {partition.device}: {e}")
            discos[partition.device if partition.device else "Desconocido"] = 0

    # Obtener información detallada del sistema operativo
    try:
        wmi_conn = wmi.WMI()
        os_info = wmi_conn.Win32_OperatingSystem()[0]
        sistema_operativo = f"{os_info.Caption} {os_info.OSArchitecture}"  # Ejemplo: Windows 11 Home Single Language 64-bit
    except Exception as e:
        print(f"Error al obtener información del sistema operativo: {e}")
        sistema_operativo = "No disponible"

    # Combinar datos en un diccionario
    datos = {
        "Ruta Archivos Guardados": ruta_archivos_guardados,
        "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Ruta Archivos Guardados System State": ruta_archivos_guardados_system_state,
        "Procesador (% usado)": cpu_usage,
        "Memoria (% usado)": memory_usage,
        "Tamaño Discos": round(total_disks / (1024**3), 2),  # Convertir a GB
        "Espacio Libre Disco C": discos.get("C:", 0),
        "Espacio Libre Disco D": discos.get("D:", 0),
        "Espacio Libre Disco I": discos.get("I:", 0),
        "Version de la Actualizacion": "1.0.0",
        "Virus Detectados": 0,
        "Sistema Operativo": sistema_operativo,  # Sistema operativo detallado
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
        item = target_list.add_item(item_properties).execute_query()
        print("Chequeo Servidor subido exitosamente a SharePoint.")
        return item.properties["ID"]  # Retorna el ID del elemento creado
    else:
        print("Error de autenticación.")
        return None

if __name__ == "__main__":
    csv_file_name = "Eventos_ayer.csv"
    events = get_events_from_yesterday()
    
    # CHEQUEO SERVIDOR
    try:
        datos_sistema = obtener_datos_sistema()
        print(f"Nombre del equipo actual: {nombre_equipo_actual}")
        
        servidor_lookup_id = obtener_servidor_id(nombre_equipo_actual)
        if servidor_lookup_id:
            chequeo_servidor_id = subir_chequeo_servidor_sharepoint(datos_sistema, servidor_lookup_id)
            
            if chequeo_servidor_id:
                if events:
                    # Consolidar eventos
                    consolidated_events = consolidate_events(events)

                    # Contar ítems de tipo Error y Critical en el consolidado
                    error_events = sum(1 for event in consolidated_events if event["Type"] and event["Type"].lower() == "error")
                    critical_events = sum(1 for event in consolidated_events if event["Type"] and event["Type"].lower() == "critical")

                    # Subir los eventos al visor de SharePoint asociados al chequeo del servidor
                    event_count = upload_events_to_sharepoint(consolidated_events, chequeo_servidor_id)
                    
                    # Crear el archivo CSV como respaldo
                    save_events_to_csv(events, csv_file_name)
                    print(f"Archivo CSV guardado en: {csv_file_name}")

                    # Enviar el correo de notificación con el resumen
                    if event_count > 0:
                        send_email_notification(len(consolidated_events), error_events, critical_events)
                else:
                    print("No se encontraron eventos para el día anterior.")
            else:
                print("No se pudo crear el Chequeo Servidor. No se subieron eventos.")
        else:
            print("No se pudo encontrar el equipo en la lista. No se realizó ninguna operación.")
    except Exception as e:
        print(f"Error general durante la ejecución: {e}")