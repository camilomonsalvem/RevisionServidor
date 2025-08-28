import csv
import io
import logging
import os
import smtplib
import socket
import subprocess
import time
import sys
import psutil
import wmi
import re
import requests
import json
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
from urllib.parse import urlparse

# Configure logging
logging.basicConfig(
    filename="app.log",
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(filename)s - Line %(lineno)d: %(message)s"
)

# Detectar si el script está empaquetado como .exe
if getattr(sys, 'frozen', False):
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))

# Cargar variables de entorno
def load_environment():
    """Cargar variables de entorno desde .env.secure y .env"""
    env_path_secure = os.path.join(script_dir, ".env.secure")
    env_path_local = os.path.join(script_dir, ".env")
    
    # Para PyInstaller, también buscar en la carpeta temporal
    if getattr(sys, 'frozen', False):
        temp_dir = sys._MEIPASS
        env_path_secure_bundled = os.path.join(temp_dir, ".env.secure")
        if os.path.exists(env_path_secure_bundled):
            load_dotenv(env_path_secure_bundled)
    
    # Cargar archivos de configuración
    if os.path.exists(env_path_secure):
        load_dotenv(env_path_secure)
    if os.path.exists(env_path_local):
        load_dotenv(env_path_local)
    
    return {
        # Azure AD
        'tenant_id': os.getenv("tenant_id"),
        'client_id': os.getenv("client_id"),  
        'client_secret': os.getenv("client_secret"),
        
        # SharePoint
        'site_url': os.getenv("site_url"),
        'site_url_soporte': os.getenv("site_url_soporte"),
        'sharepoint_folder': os.getenv("sharepoint_folder"),
        
        # Listas
        'list_name_visor_eventos': os.getenv("list_name_visor_eventos"),
        'list_name_inventario': os.getenv("list_name_inventario"),
        'list_name_chequeo_servidor': os.getenv("list_name_chequeo_servidor"),
        
        # Sistema
        'empresa': os.getenv("empresa", "EmpresaDefault"),
        'ruta_archivos_guardados': os.getenv("ruta_archivos_guardados"),
        'ruta_archivos_guardados_system_state': os.getenv("ruta_archivos_guardados_system_state"),
        
        # Email
        'smtp_server': os.getenv("smtp_server"),
        'smtp_port': os.getenv("smtp_port"),
        'email_sender': os.getenv("email_sender"),
        'email_password': os.getenv("email_password"),
        'email_recipients': os.getenv("email_recipients", ""),
        'powerapps_app_link': os.getenv("powerapps_app_link")
    }

# Cargar configuración
env = load_environment()

# Variables globales
empresa = env['empresa']
nombre_equipo_actual = socket.gethostname()
ruta_archivos_guardados = env['ruta_archivos_guardados']
ruta_archivos_guardados_system_state = env['ruta_archivos_guardados_system_state']

# SharePoint
site_url = env['site_url']
site_url_soporte = env['site_url_soporte']
sharepoint_folder = env['sharepoint_folder']

# Listas
list_name_visor_eventos = env['list_name_visor_eventos']
list_name_inventario = env['list_name_inventario']
list_name_chequeo_servidor = env['list_name_chequeo_servidor']

# Email
smtp_server = env['smtp_server']
smtp_port = env['smtp_port']
email_sender = env['email_sender']
email_password = env['email_password']
powerapps_app_link = env['powerapps_app_link']

# Procesar lista de destinatarios de email
email_recipients = env['email_recipients'].replace(" ", "").split(',') if env['email_recipients'] else []

# Imprimir configuración cargada
print("\n===== CONFIGURACIÓN CARGADA CORRECTAMENTE =====\n")
print(f"Empresa: {empresa}")
print(f"Nombre del equipo actual: {nombre_equipo_actual}")
print(f"Ruta Archivos Guardados: {ruta_archivos_guardados}")
print(f"Ruta Archivos Guardados System State: {ruta_archivos_guardados_system_state}")

# Verificar configuración crítica
if not all([env['tenant_id'], env['client_id'], env['client_secret'], site_url]):
    raise ValueError("Error: Faltan credenciales de Azure AD en .env.secure")
if not email_recipients:
    raise ValueError("Error: No se encontró la variable 'email_recipients' en el archivo .env")

# ===========================
# FUNCIONES AZURE AD
# ===========================

def get_azure_token():
    """Obtener token de acceso usando Azure AD"""
    url = f"https://login.microsoftonline.com/{env['tenant_id']}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": env['client_id'],
        "client_secret": env['client_secret'],
        "scope": "https://graph.microsoft.com/.default",
    }
    
    try:
        response = requests.post(url, data=data)
        response.raise_for_status()
        return response.json()["access_token"]
    except Exception as e:
        logging.error(f"Error al obtener token de Azure AD: {e}")
        raise

def parse_site_url(site_url):
    """Parsear URL de SharePoint"""
    parsed = urlparse(site_url)
    hostname = parsed.hostname
    path = parsed.path.strip("/")
    return hostname, path

def get_site_id(token, site_url):
    """Obtener ID del sitio de SharePoint"""
    hostname, path = parse_site_url(site_url)
    url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{path}?$select=id"
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()["id"]
    except Exception as e:
        logging.error(f"Error al obtener Site ID: {e}")
        raise

def get_list_id(token, site_id, list_name):
    """Obtener ID de una lista específica"""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists?$filter=displayName eq '{list_name}'"
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        data = response.json()
        lists = data.get('value', [])
        
        if lists:
            return lists[0]['id']
        else:
            logging.error(f"Lista '{list_name}' no encontrada")
            return None
    except Exception as e:
        logging.error(f"Error al obtener ID de lista '{list_name}': {e}")
        return None

# ===========================
# FUNCIONES PARA EL VISOR DE EVENTOS (CONSERVADAS)
# ===========================

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

def get_all_events(logfile):
    print(f"Obteniendo eventos del registro: {logfile}")
    logging.info(f"Obteniendo eventos del registro: {logfile}")

    wmi_o = wmi.WMI('.')
    query = f"SELECT * FROM Win32_NTLogEvent WHERE Logfile='{logfile}'"
    try:
        results = wmi_o.query(query)
        print(f"Eventos obtenidos del registro {logfile}: {len(results)}")
        logging.info(f"Eventos obtenidos del registro {logfile}: {len(results)}")
        return results
    except Exception as e:
        msg_error = f"Error al ejecutar la consulta para {logfile}: {e}"
        print(msg_error)
        logging.error(msg_error)
        return []

def clear_event_log(log_name):
    """Elimina todos los eventos de un registro específico en el Visor de Eventos."""
    try:
        print(f"Intentando eliminar eventos del registro: {log_name}")
        subprocess.run(f"wevtutil cl \"{log_name}\"", shell=True, check=True, text=True)
        print(f"Eventos del registro '{log_name}' eliminados correctamente.")
        logging.info(f"Eventos del registro '{log_name}' eliminados correctamente.")
    except subprocess.CalledProcessError as e:
        print(f"Error al eliminar los eventos del registro '{log_name}': {e}")
        logging.error(f"Error al eliminar los eventos del registro '{log_name}': {e}")
    except Exception as e:
        print(f"Error inesperado al procesar el registro '{log_name}': {e}")
        logging.error(f"Error inesperado al procesar el registro '{log_name}': {e}")

def consolidate_events(events):
    consolidated = {}

    for event in events:
        try:
            key = (
                event.SourceName or "Desconocido",
                event.Type or "Desconocido",
                event.EventCode or "Desconocido"
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

def get_folder_id(token, site_id, folder_path):
    """Obtener ID de la carpeta navegando paso a paso"""
    # Limpiar y dividir la ruta
    clean_path = folder_path.replace('/sites/SoporteTecnico813/Documentos compartidos/', '')
    path_parts = [part for part in clean_path.split('/') if part]
    
    current_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root"
    headers = {"Authorization": f"Bearer {token}"}
    
    # Navegar paso a paso
    for folder_name in path_parts:
        try:
            children_url = f"{current_url}/children"
            response = requests.get(children_url, headers=headers)
            response.raise_for_status()
            
            data = response.json()
            items = data.get('value', [])
            
            # Buscar la carpeta
            found = False
            for item in items:
                if item.get('name') == folder_name and item.get('folder'):
                    current_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item['id']}"
                    found = True
                    break
            
            if not found:
                logging.error(f"Carpeta '{folder_name}' no encontrada")
                return None
                        
        except Exception as e:
            logging.error(f"Error navegando a {folder_name}: {e}")
            return None
    
    # Extraer el ID de la URL final
    folder_id = current_url.split('/')[-1]
    return folder_id

def upload_csv_buffer_to_sharepoint(csv_buffer, file_name, sharepoint_folder):
    """Subir CSV a SharePoint usando Microsoft Graph API con ID de carpeta"""
    try:
        # Obtener token y site ID para soporte
        token = get_azure_token()
        site_id = get_site_id(token, site_url_soporte)
        
        # Obtener ID de la carpeta objetivo
        folder_id = get_folder_id(token, site_id, sharepoint_folder)
        
        if not folder_id:
            error_msg = f"No se pudo encontrar la carpeta: {sharepoint_folder}"
            logging.error(error_msg)
            print(error_msg)
            return False
        
        # URL correcta usando ID de carpeta
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{folder_id}:/{file_name}:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "text/csv"
        }
        
        # Subir el archivo
        response = requests.put(url, headers=headers, data=csv_buffer.getvalue().encode("utf-8"))
        response.raise_for_status()
        
        print(f"Archivo {file_name} subido exitosamente a SharePoint.")
        logging.info(f"Archivo {file_name} subido exitosamente a SharePoint.")
        return True
        
    except Exception as e:
        error_msg = f"Error al subir el archivo CSV a SharePoint: {e}"
        logging.error(error_msg)
        print(error_msg)
        return False

def save_events_to_csv_and_upload(events, sharepoint_folder, file_name):
    headers = ["Nombre de Registro", "Origen", "ID", "Nivel", "Categoria de Tarea", "Registrado", "Equipo", "Usuario", "Mensaje", "EventRecordID"]
    try:
        # Usar StringIO para escribir el contenido del CSV en memoria
        csv_buffer = io.StringIO()
        writer = csv.DictWriter(csv_buffer, fieldnames=headers)
        writer.writeheader()

        for event in events:
            if isinstance(event, str):  # Si el evento es una cadena (obtenido con wevtutil)
                parsed_event = parse_event_setup(event)
                writer.writerow(parsed_event)
            else:  # Si el evento es un objeto (obtenido con WMI)
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
        csv_buffer.seek(0)
        upload_csv_buffer_to_sharepoint(csv_buffer, file_name, sharepoint_folder)
        print("Archivo CSV generado y subido exitosamente a SharePoint.")
        logging.info("Archivo CSV generado y subido exitosamente a SharePoint.")
    except Exception as e:
        msg_error = f"Error al generar o subir el archivo CSV: {e}"
        print(msg_error)
        logging.error(msg_error)

def upload_events_to_sharepoint(events, chequeo_servidor_id, logfile):
    """Subir eventos a SharePoint usando Microsoft Graph API"""
    try:
        token = get_azure_token()
        site_id = get_site_id(token, site_url)
        list_id = get_list_id(token, site_id, list_name_visor_eventos)
        
        if not list_id:
            logging.error("No se pudo obtener el ID de la lista de Visor de Eventos")
            return 0
        
        event_count = 0

        # Filtrar eventos por nivel en inglés y español
        filtered_events = [
            event for event in events if (event['Type'] or "").lower() in 
            ['error', 'critical', 'warning', 'crítico', 'advertencia']
        ]

        for event in filtered_events:
            try:
                # Preparar datos del elemento
                item_data = {
                    'fields': {
                        'Title': logfile,
                        'Id_x0020_del_x0020_Evento': str(event['EventCode'] or 0),
                        'Origen': event['SourceName'] or "Desconocido",
                        'Categoria_x0020_de_x0020_Tarea': event['CategoryString'] or "Ninguno",
                        'Detalle': event['Message'] or "N/A",
                        'Fecha': event['TimeGenerated'] or "N/A",
                        'Nivel': event['Type'] or "N/A",
                        'User': event['User'] or "N/A",
                        'No_x0020_de_x0020_Eventos': event['NoEventos'] or 0,
                        'ID_x0020_Chequeo_x0020_ServidorLookupId': chequeo_servidor_id
                    }
                }
                
                # Crear elemento en la lista
                url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"
                headers = {
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json"
                }
                
                response = requests.post(url, headers=headers, json=item_data)
                response.raise_for_status()
                
                event_count += 1
                time.sleep(0.5)
                
            except Exception as e:
                msg_error = f"Error al cargar el evento {event.get('EventRecordID', 'Desconocido')}: {e}"
                print(msg_error)
                logging.error(msg_error)

        print(f"Total de eventos cargados a SharePoint: {event_count}")
        return event_count

    except Exception as e:
        msg_error = f"Error al subir los eventos a SharePoint: {e}"
        print(msg_error)
        logging.error(msg_error)
        return 0

# ===========================
# FUNCIONES PARA CHEQUEO SERVIDOR (CONSERVADAS)
# ===========================

def get_system_data():
    # Procesador y memoria
    cpu_usage = psutil.cpu_percent(interval=1)
    memory_info = psutil.virtual_memory()
    memory_usage = memory_info.percent

    # Información de los discos
    discos = {}
    total_disks = 0

    for partition in psutil.disk_partitions():
        try:
            if "cdrom" in partition.opts or partition.fstype == "":
                print(f"Unidad omitida (no es disco local): {partition.device}")
                continue

            disk_usage = psutil.disk_usage(partition.mountpoint)
            disco_nombre = partition.device.strip("\\") if partition.device else "Desconocido"
            discos[disco_nombre] = round(disk_usage.free / (1024**3), 2)
            total_disks += disk_usage.total
        except PermissionError:
            msg_error = f"Permiso denegado al acceder a la unidad {partition.device}"
            print(msg_error)
            logging.error(msg_error)
            discos[partition.device if partition.device else "Desconocido"] = 0
        except Exception as e:
            print(f"Error al acceder a la unidad {partition.device}: {e}")
            discos[partition.device if partition.device else "Desconocido"] = 0

    # Obtener información del sistema operativo
    try:
        wmi_conn = wmi.WMI()
        os_info = wmi_conn.Win32_OperatingSystem()[0]
        sistema_operativo = f"{os_info.Caption} {os_info.Version} {os_info.OSArchitecture}"
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
        "Tamaño Discos": round(total_disks / (1024**3), 2),
        "Espacio Libre Disco C": discos.get("C:", 0),
        "Espacio Libre Disco D": discos.get("D:", 0),
        "Espacio Libre Disco I": discos.get("I:", 0),
        "Version de la Actualizacion": "1.0.0",
        "Virus Detectados": 0,
        "Sistema Operativo": sistema_operativo,
    }

    return datos

def get_server_id(nombre_equipo):
    """Obtener el ID del equipo en la lista Inventario usando Microsoft Graph API"""
    try:
        token = get_azure_token()
        site_id = get_site_id(token, site_url)
        list_id = get_list_id(token, site_id, list_name_inventario)
        
        if not list_id:
            logging.error(f"No se pudo obtener el ID de la lista {list_name_inventario}")
            return None
        
        # Buscar el equipo por nombre - arreglar el filtro OData
        # Usar comillas simples y escape adecuado
        filter_query = f"fields/Title eq '{nombre_equipo}'"
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"
        params = {
            '$expand': 'fields',
            '$filter': filter_query
        }
        headers = {"Authorization": f"Bearer {token}"}
        
        response = requests.get(url, headers=headers, params=params)
        
        # Si el filtro falla, intentar obtener todos los elementos y filtrar manualmente
        if response.status_code == 400:
            print("Filtro OData falló, intentando búsqueda manual...")
            logging.info("Filtro OData falló, intentando búsqueda manual...")
            
            # Obtener todos los elementos de la lista
            url_all = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?$expand=fields"
            response = requests.get(url_all, headers=headers)
            response.raise_for_status()
            
            data = response.json()
            items = data.get('value', [])
            
            # Buscar manualmente
            for item in items:
                if item.get('fields', {}).get('Title') == nombre_equipo:
                    return item['fields']['id']
            
            logging.error(f"No se encontró el equipo con nombre: {nombre_equipo} en la lista {list_name_inventario}.")
            print(f"No se encontró el equipo con nombre: {nombre_equipo} en la lista {list_name_inventario}.")
            return None
        else:
            response.raise_for_status()
            data = response.json()
            items = data.get('value', [])
            
            if items:
                return items[0]['fields']['id']
            else:
                logging.error(f"No se encontró el equipo con nombre: {nombre_equipo} en la lista {list_name_inventario}.")
                print(f"No se encontró el equipo con nombre: {nombre_equipo} en la lista {list_name_inventario}.")
                return None
            
    except Exception as e:
        msg_error = f"Error al obtener el ID del servidor: {e}"
        print(msg_error)
        logging.error(msg_error)
        return None

def send_chequeo_servidor_sharepoint(datos, servidor_lookup_id):
    """Subir datos del chequeo servidor usando Microsoft Graph API"""
    try:
        token = get_azure_token()
        site_id = get_site_id(token, site_url)
        list_id = get_list_id(token, site_id, list_name_chequeo_servidor)
        
        if not list_id:
            logging.error("No se pudo obtener el ID de la lista de Chequeo Servidor")
            return None
        
        # Preparar datos del elemento
        item_data = {
            'fields': {
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
                "ServidorLookupId": servidor_lookup_id,
            }
        }
        
        # Crear elemento en la lista
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        response = requests.post(url, headers=headers, json=item_data)
        response.raise_for_status()
        
        created_item = response.json()
        item_id = created_item['fields']['id']
        
        print("Chequeo Servidor subido exitosamente a SharePoint.")
        return item_id
        
    except Exception as e:
        msg_error = f"Error al subir el Chequeo Servidor a SharePoint: {e}"
        print(msg_error)
        logging.error(msg_error)
        return None

# ===========================
# FUNCIONES PARA EVENTOS SETUP (CONSERVADAS)
# ===========================

def get_setup_events_with_wevtutil_setup():
    print("Obteniendo eventos del registro Setup usando wevtutil...")
    try:
        command = "wevtutil qe Setup /f:Text"
        result = subprocess.check_output(command, shell=True, text=True, stderr=subprocess.DEVNULL)
        events = result.strip().split("\n\n")
        print(f"Eventos obtenidos: {len(events)}")
        return events
    except subprocess.CalledProcessError as e:
        print(f"Error al obtener eventos de Setup: {e}")
        return []

def parse_event_setup(event):
    def extract_field(pattern, text):
        match = re.search(pattern, text, re.MULTILINE)
        return match.group(1).strip().replace(',', ' ') if match else "N/A"

    # Formatear la fecha
    raw_date = extract_field(r"Date:\s*(.*)", event)
    try:
        formatted_date = datetime.strptime(raw_date, "%Y-%m-%dT%H:%M:%S.%f0Z").strftime("%Y-%m-%d %H:%M:%S")
    except ValueError:
        formatted_date = raw_date

    # Mapear SID a nombres de usuario conocidos
    user_sid = extract_field(r"User:\s*(.*)", event)
    user_mapping = {"S-1-5-18": "SYSTEM"}
    user_name = user_mapping.get(user_sid, user_sid)

    return {
        "Nombre de Registro": extract_field(r"Log Name:\s*(.*)", event),
        "Origen": extract_field(r"Source:\s*(.*)", event),
        "ID": extract_field(r"Event ID:\s*(\d+)", event),
        "Nivel": extract_field(r"Level:\s*(.*)", event),
        "Categoria de Tarea": extract_field(r"Task Category:\s*(.*)", event),
        "Registrado": formatted_date,
        "Equipo": extract_field(r"Computer:\s*(.*)", event),
        "Usuario": user_name,
        "Mensaje": extract_field(r"Description:\s*([\s\S]*?)\s*(?:Keywords:|$)", event).replace('\n', ' ')
    }

# ===========================
# FUNCIÓN DE EMAIL (CONSERVADA)
# ===========================

def send_email_notification(total_events, error_events, critical_events, warning_events, id_inventario, id_chequeo_servidor, log_type, tipoEventos):
    try:
        powerapps_link = (
            f"{powerapps_app_link}&screen=visor&idInventario={id_inventario}&idChequeoServidor={id_chequeo_servidor}&nombreRegistro={log_type}"
        )

        # Cargar plantilla de email
        email_template_path = os.path.join(script_dir, "email_template.html")
        if getattr(sys, 'frozen', False):
            temp_dir = sys._MEIPASS
            email_template_path_bundled = os.path.join(temp_dir, "email_template.html")
            if os.path.exists(email_template_path_bundled):
                email_template_path = email_template_path_bundled

        try:
            with open(email_template_path, "r", encoding="utf-8") as file:
                html_template = file.read()
        except FileNotFoundError:
            print(f"Error: email_template.html not found at {email_template_path}")
            return

        html_body = html_template.replace("{{total_events}}", str(total_events))
        html_body = html_body.replace("{{error_events}}", str(error_events))
        html_body = html_body.replace("{{critical_events}}", str(critical_events))
        html_body = html_body.replace("{{warning_events}}", str(warning_events))    
        html_body = html_body.replace("{{powerapps_link}}", powerapps_link)

        # Determinar el emoji basado en la jerarquía de eventos
        if tipoEventos == "Critico":
            emoji = "🛑"
            priority = "1"  # Alta prioridad
            importance = "High"
        elif tipoEventos == "Error":
            emoji = "❌"
            priority = "3"  # Normal
            importance = "Normal"
        elif tipoEventos == "Advertencia":
            emoji = "⚠️"
            priority = "3"  # Normal
            importance = "Normal"
        else:
            emoji = "✅"  # Información en caso de que no haya eventos
            priority = "5"  # Normal
            importance = "Low"
        
        fecha_actual = datetime.now().strftime("%d/%m/%Y")
        subject = f"{emoji} Se subió el Visor de Eventos de {empresa} - {log_type} - {fecha_actual}"

        msg = MIMEMultipart("alternative")
        msg['From'] = email_sender
        msg['To'] = ", ".join(email_recipients)
        msg['Subject'] = subject
        msg['X-Priority'] = priority
        msg['Importance'] = importance
        msg.attach(MIMEText(html_body, "html"))

        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.starttls()
        server.login(email_sender, email_password)
        server.sendmail(email_sender, email_recipients, msg.as_string())
        server.quit()

        print(f"Correo de notificación enviado a {', '.join(email_recipients)}")

    except Exception as e:
        msg_error = f"Error al enviar el correo de notificación: {e}"
        print(msg_error)
        logging.error(msg_error)

# ===========================
# FUNCIÓN PRINCIPAL
# ===========================

if __name__ == "__main__":
    fecha_actual = datetime.now().strftime("%d-%m-%Y")

    try:
        print("🔄 Iniciando proceso de monitoreo con Azure AD...")
        
        # Verificar conexión a Azure AD
        token = get_azure_token()
        print("✅ Conexión a Azure AD exitosa")

        # CHEQUEO SERVIDOR
        datos_sistema = get_system_data()

        # Obtener el ID del servidor
        servidor_lookup_id = get_server_id(nombre_equipo_actual)
        chequeo_servidor_id = None  # Inicializar la variable

        if servidor_lookup_id:
            # Subir datos del chequeo del servidor a SharePoint
            chequeo_servidor_id = send_chequeo_servidor_sharepoint(datos_sistema, servidor_lookup_id)
            print(f"ID del Chequeo Servidor: {chequeo_servidor_id}")

            if chequeo_servidor_id:
                logfiles = ['System', 'Application']  # Lista de registros a procesar
                for logfile in logfiles:
                    print(f"Procesando eventos del registro: {logfile}")
                    csv_file_name = f"VisorEventos_{logfile}_{fecha_actual}_{chequeo_servidor_id}.csv"

                    # Obtener eventos para el registro actual
                    events = get_all_events(logfile)
                
                    # Limpiar el visor de eventos si todo el proceso fue exitoso
                    clear_event_log(logfile)

                    if events:
                        # Consolidar eventos
                        consolidated_events = consolidate_events(events)

                        # Contar ítems de tipo Error, Critical y Warning en el consolidado
                        error_events = sum(1 for event in consolidated_events if event["Type"] and event["Type"].lower() == "error")
                        critical_events = sum(1 for event in consolidated_events if event["Type"] and event["Type"].lower() in ["critical", "crítico", "critico"])
                        warning_events = sum(1 for event in consolidated_events if event["Type"] and event["Type"].lower() in ["warning", "advertencia"])

                        # Subir los eventos al visor de SharePoint asociados al chequeo del servidor
                        event_count = upload_events_to_sharepoint(consolidated_events, chequeo_servidor_id, logfile)

                        # Generar y subir el archivo CSV a SharePoint
                        save_events_to_csv_and_upload(events, sharepoint_folder, csv_file_name)
                        
                        # Determinar mayor cantidad de eventos
                        if critical_events > 0:
                            tipoEventos = "Critico"
                        elif error_events > 0:
                            tipoEventos = "Error"
                        elif warning_events > 0:
                            tipoEventos = "Advertencia"
                        else:
                            tipoEventos = "Ninguno"
                        
                        # Enviar el correo de notificación con el resumen
                        send_email_notification(
                            len(consolidated_events),
                            error_events,
                            critical_events,
                            warning_events,
                            servidor_lookup_id,  # idInventario
                            chequeo_servidor_id,  # idChequeoServidor
                            logfile,
                            tipoEventos
                        )
                    else:
                        logging.error(f"No se encontraron eventos en el registro {logfile}.")
                        print(f"No se encontraron eventos en el registro {logfile}.")
            else:
                logging.error("No se pudo crear el Chequeo Servidor. No se subieron eventos.")
                print("No se pudo crear el Chequeo Servidor. No se subieron eventos.")
        else:
            logging.error("No se encontró el ID del servidor. No se subieron eventos.")
            print("No se pudo encontrar el equipo en la lista. No se realizó ninguna operación.")

        # Procesar únicamente los registros Security y Setup
        logfiles_security_setup = ['Setup','Security']  # Nuevos registros a procesar

        # Solo procesar estos registros si tenemos un chequeo_servidor_id válido
        if chequeo_servidor_id:
            for logfile in logfiles_security_setup:
                print(f"Procesando eventos del registro: {logfile}")
                csv_file_name = f"VisorEventos_{logfile}_{fecha_actual}_{chequeo_servidor_id}.csv"

                # Obtener eventos para el registro actual
                if logfile == 'Setup':
                    events = get_setup_events_with_wevtutil_setup()
                else:
                    events = get_all_events(logfile)
                
                clear_event_log(logfile)

                if events:
                    # Generar y subir el archivo CSV a SharePoint
                    save_events_to_csv_and_upload(events, sharepoint_folder, csv_file_name)
                    print(f"Archivo CSV para {logfile} generado y subido exitosamente.")
                else:
                    print(f"No se encontraron eventos en el registro {logfile}.")
        else:
            print("Saltando procesamiento de registros Security y Setup (no hay chequeo_servidor_id)")
            logging.info("Saltando procesamiento de registros Security y Setup (no hay chequeo_servidor_id)")

        print("🎉 Proceso completado exitosamente con Azure AD")

    except Exception as e:
        logging.error(f"Error general durante la ejecución: {e}")
        print(f"Error general durante la ejecución: {e}")
        import traceback
        traceback.print_exc()