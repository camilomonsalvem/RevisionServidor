import csv
import logging
import os
import smtplib
import socket
import time
import psutil
import wmi
import io
import subprocess
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

# Configure logging
logging.basicConfig(
    filename="app.log",
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(filename)s - Line %(lineno)d: %(message)s"
)

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

# Configuración de SharePoint Soporte Tecnico para archivo CSV
site_url_soporte = os.getenv("site_url_soporte")
sharepoint_folder = os.getenv("sharepoint_folder")

print("Configuración cargada correctamente...")
print(f"Usuario: {username}")
print(f"URL del sitio: {site_url}")
print(f"Lista Visor de Eventos: {list_name_visor_eventos}")
print(f"Lista Inventario PC Construsol: {list_name_inventario}")
print(f"Lista Chequeo Servidor: {list_name_chequeo_servidor}")
print(f"Nombre del equipo actual: {nombre_equipo_actual}")
print(f"Ruta Archivos Guardados: {ruta_archivos_guardados}")
print(f"Ruta Archivos Guardados System State: {ruta_archivos_guardados_system_state}")

# Funciones para el VISOR DE EVENTOS

def format_wmi_time(wmi_time):
    try:
        if not wmi_time:
            return "N/A"
        utc_time = datetime.strptime(wmi_time.split('.')[0], '%Y%m%d%H%M%S').replace(tzinfo=timezone.utc)
        bogota_time = utc_time.astimezone(timezone(timedelta(hours=-5)))
        return bogota_time.strftime('%Y-%m-%d %H:%M:%S')
    except Exception as e:
        msg_error = f"Error al formatear el tiempo: {e}"
        print(msg_error)
        logging.error(msg_error)
        return "N/A"

def clear_event_log(log_type):
    try:
        print(f"Intentando eliminar los eventos del registro '{log_type}'...")
        # Usar wevtutil para limpiar el registro
        subprocess.run(["wevtutil", "cl", log_type], check=True)
        print(f"Todos los eventos del registro '{log_type}' han sido eliminados.")
        logging.info(f"Todos los eventos del registro '{log_type}' han sido eliminados.")
    except subprocess.CalledProcessError as e:
        msg_error = f"Error al eliminar los eventos del registro '{log_type}': {e}"
        print(msg_error)
        logging.error(msg_error)
    except Exception as e:
        msg_error = f"Error inesperado al intentar eliminar el registro '{log_type}': {e}"
        print(msg_error)
        logging.error(msg_error)

# Modificar get_events para traer todos los eventos
def get_all_events(log_type):
    print(f"Obteniendo todos los eventos en el registro '{log_type}'....")
    logging.info(f"Obteniendo todos los eventos en el registro '{log_type}'.")
    wmi_o = wmi.WMI('.')
    query = f"SELECT * FROM Win32_NTLogEvent WHERE Logfile='{log_type}'"

    try:
        results = wmi_o.query(query)
        if not results:
            print(f"No se encontraron eventos en el registro '{log_type}'.")
            return []
        return results
    except Exception as e:
        msg_error = f"Error al obtener eventos en '{log_type}': {e}"
        print(msg_error)
        logging.error(msg_error)
        return []

def save_events_to_csv_and_upload(events, sharepoint_folder, file_name):
    headers = ["Nombre de Registro", "Origen", "ID", "Nivel", "Categoria de Tarea", "Registrado", "Equipo", "Usuario", "Mensaje", "EventRecordID"]
    try:
        # Usar StringIO para escribir el contenido del CSV en memoria
        csv_buffer = io.StringIO()
        writer = csv.DictWriter(csv_buffer, fieldnames=headers)
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

        # Subir el contenido del CSV a SharePoint
        csv_buffer.seek(0)  # Volver al inicio del buffer para leerlo
        upload_csv_buffer_to_sharepoint(csv_buffer, file_name, sharepoint_folder)
        print("Archivo CSV generado y subido exitosamente a SharePoint.")
        logging.info("Archivo CSV generado y subido exitosamente a SharePoint.")
    except Exception as e:
        msg_error = f"Error al generar o subir el archivo CSV: {e}"
        print(msg_error)
        logging.error(msg_error)

def upload_csv_buffer_to_sharepoint(csv_buffer, file_name, sharepoint_folder):
    try:
        # Autenticación
        context_auth = AuthenticationContext(site_url_soporte)
        if context_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(site_url_soporte, context_auth)
            logging.info("Autenticación exitosa para subir el archivo CSV.")
        else:
            error_msg = f"Error de autenticación: {context_auth.get_last_error()}"
            logging.error(error_msg)
            print(error_msg)
            return False

        # Ruta relativa correcta
        target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_folder)

        # Subir el contenido del CSV al archivo en SharePoint
        target_folder.upload_file(file_name, csv_buffer.getvalue().encode("utf-8")).execute_query()
        print(f"Archivo {file_name} subido exitosamente a SharePoint en {sharepoint_folder}.")
        logging.info(f"Archivo {file_name} subido exitosamente a SharePoint en {sharepoint_folder}.")
        return True
    except Exception as e:
        error_msg = f"Error al subir el archivo CSV a SharePoint: {e}"
        logging.error(error_msg)
        print(error_msg)
        return False

def consolidate_events(events):
    consolidated = {}

    for event in events:
        try:
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
        except Exception as e:
            msg_error = f"Error al consolidar eventos: {e}"
            print(msg_error)
            logging.error(msg_error)

    return list(consolidated.values())

def upload_events_to_sharepoint(events, chequeo_servidor_id, log_type):
    try:
        ctx_auth = AuthenticationContext(site_url)
        if ctx_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(site_url, ctx_auth)
            logging.info("Autenticación exitosa")
            print("Autenticación exitosa")
        else:
            logging.error(f"Error al autenticar: {ctx_auth.get_last_error()}")
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
                    'TipoRegistro': log_type,  # Nuevo campo para el tipo de registro
                    'ID_x0020_Chequeo_x0020_ServidorId': chequeo_servidor_id
                }
                list_obj.add_item(item_properties)
                ctx.execute_query()
                event_count += 1
                time.sleep(0.5)
            except Exception as e:
                msg_error = f"Error al cargar el evento {event.get('EventRecordID', 'Desconocido')} del registro {log_type}: {e}"
                print(msg_error)
                logging.error(msg_error)

        print(f"Total de eventos cargados a SharePoint para {log_type}: {event_count}")
        return event_count

    except Exception as e:
        msg_error = f"Error al subir los eventos a SharePoint para {log_type}: {e}"
        print(msg_error)
        logging.error(msg_error)
        return 0

def send_email_notification(total_events, error_events, critical_events, id_inventario, id_chequeo_servidor):
    try:
        # Crear el enlace al aplicativo de PowerApps
        powerapps_link = (
            f"https://apps.powerapps.com/play/e/default-13fbbcde-1002-4ff4-b26f-ae75208bb81b/a/290f92a6-7699-4859-9b46-4ae1ee60b047"
            f"?tenantId=13fbbcde-1002-4ff4-b26f-ae75208bb81b&hint=59684f0a-d15c-451a-9ffd-b73d5f5fae3a&sourcetime=1737999244624"
            f"&screen=visor&idInventario={id_inventario}&idChequeoServidor={id_chequeo_servidor}"
        )

        # Leer la plantilla HTML
        with open("email_template.html", "r", encoding="utf-8") as file:
            html_template = file.read()

        # Reemplazar variables en el HTML
        html_body = html_template.replace("{{total_events}}", str(total_events))
        html_body = html_body.replace("{{error_events}}", str(error_events))
        html_body = html_body.replace("{{critical_events}}", str(critical_events))
        html_body = html_body.replace("{{powerapps_link}}", powerapps_link)

        # Configurar el correo
        msg = MIMEMultipart("alternative")
        msg['From'] = email_sender
        msg['To'] = email_recipient
        msg['Subject'] = f"Reporte de Eventos: {total_events} procesados, {error_events} errores, {critical_events} críticos"
        msg.attach(MIMEText(html_body, "html"))

        # Enviar el correo
        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.starttls()
        server.login(email_sender, email_password)
        server.sendmail(email_sender, email_recipient, msg.as_string())
        server.quit()

        print(f"Correo de notificación enviado a {email_recipient}")

    except Exception as e:
        msg_error = f"Error al enviar el correo de notificación: {e}"
        print(msg_error)
        logging.error(msg_error)

# FUNCIONES PARA EL CHEQUEO SERVIDOR

# Obtener información del sistema
def get_system_data():
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
            msg_error = f"Permiso denegado al acceder a la unidad {partition.device}"
            print(msg_error)
            logging.error(msg_error)
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
        msg_error = f"Error al obtener información del sistema operativo: {e}"
        print(msg_error)
        logging.error(msg_error)
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
def get_server_id(nombre_equipo):
    try:
        context_auth = AuthenticationContext(site_url)
        if context_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(site_url, context_auth)
            inventario_list = ctx.web.lists.get_by_title(list_name_inventario)
            items = inventario_list.get_items().execute_query()

            for item in items:
                if item.properties["Title"] == nombre_equipo:  # Comparar con el nombre del equipo
                    return item.properties["ID"]

            logging.error(f"No se encontró el equipo con nombre: {nombre_equipo} en la lista {list_name_inventario}.")
            print(f"No se encontró el equipo con nombre: {nombre_equipo} en la lista {list_name_inventario}.")
            return None
        else:
            msg_error = "Error de autenticación al obtener el ID del servidor"
            print(msg_error)
            logging.error(msg_error)
            return None
    except Exception as e:
        msg_error = f"Error al obtener el ID del servidor: {e}"
        print(msg_error)
        logging.error(msg_error)
        return None

# Subir datos a SharePoint
def send_chequeo_servidor_sharepoint(datos, servidor_lookup_id):
    try:
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
            logging.error("Error de autenticación al subir el Chequeo Servidor.")
            print("Error de autenticación.")
            return None
    except Exception as e:
        msg_error = f"Error al subir el Chequeo Servidor a SharePoint: {e}"
        print(msg_error)
        logging.error(msg_error)
        return None

if __name__ == "__main__":
    fecha_actual = datetime.now().strftime("%d-%m-%Y")

    try:
        # CHEQUEO SERVIDOR
        datos_sistema = get_system_data()
        print(f"Nombre del equipo actual: {nombre_equipo_actual}")

        # Obtener el ID del servidor
        servidor_lookup_id = get_server_id(nombre_equipo_actual)
        print(f"ID del servidor: {servidor_lookup_id}")

        if servidor_lookup_id:
            # Subir datos del chequeo del servidor a SharePoint
            chequeo_servidor_id = send_chequeo_servidor_sharepoint(datos_sistema, servidor_lookup_id)
            print(f"ID del Chequeo Servidor: {chequeo_servidor_id}")

            if chequeo_servidor_id:
                # Procesar ambos tipos de registro
                for log_type in ["System", "Application"]:
                    events = get_all_events(log_type)
                    if events:
                        # Consolidar eventos
                        consolidated_events = consolidate_events(events)

                        # Contar ítems de tipo Error y Critical
                        error_events = sum(1 for event in consolidated_events if event["Type"] and event["Type"].lower() == "error")
                        critical_events = sum(1 for event in consolidated_events if event["Type"] and event["Type"].lower() == "critical")

                        # Subir eventos a SharePoint
                        event_count = upload_events_to_sharepoint(consolidated_events, chequeo_servidor_id, log_type)

                        # Generar y subir archivo CSV a SharePoint
                        csv_file_name = f"VisorEventos_{log_type}_{fecha_actual}.csv"
                        save_events_to_csv_and_upload(events, sharepoint_folder, csv_file_name)

                        # Enviar el correo de notificación con el resumen para este tipo de registro
                        send_email_notification(
                            event_count,         # Total de eventos procesados para este log_type
                            error_events,        # Total de eventos con errores para este log_type
                            critical_events,     # Total de eventos críticos para este log_type
                            servidor_lookup_id,  # idInventario
                            chequeo_servidor_id  # idChequeoServidor
                        )
                        print(f"Procesamiento de eventos completado para '{log_type}'.")
                    else:
                        print(f"No se encontraron eventos en el registro '{log_type}'.")

                # Eliminar eventos de ambos registros solo si todo se realizó correctamente
                for log_type in ["System", "Application"]:
                    clear_event_log(log_type)

            else:
                logging.error("No se pudo crear el Chequeo Servidor. No se subieron eventos.")
                print("No se pudo crear el Chequeo Servidor. No se subieron eventos.")
        else:
            logging.error("No se encontró el ID del servidor. No se subieron eventos.")
            print("No se pudo encontrar el equipo en la lista. No se realizó ninguna operación.")
    except Exception as e:
        logging.error(f"Error general durante la ejecución: {e}")
        print(f"Error general durante la ejecución: {e}")