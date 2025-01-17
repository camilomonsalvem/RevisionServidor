import wmi
import csv
import time
from collections import Counter
from datetime import datetime, timedelta, timezone
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configuración de SharePoint
site_url = 'https://idtsas.sharepoint.com/sites/Bk-Empresas'
list_name = 'Visor de Eventos Construsol'
username = 'info@idtsas.com'
password = '1D2022++'

# Configuración de correo
smtp_server = "smtp.office365.com"  # Reemplaza con tu servidor SMTP
smtp_port = 587  # Puerto SMTP
email_sender = username  # Dirección de correo del remitente
email_password = password  # Contraseña del correo
email_recipient = "cmonsalve@idtsas.com"  # Dirección de correo del destinatario

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

if __name__ == "__main__":
    csv_file_name = "Eventos_ayer.csv"
    events = get_events_from_yesterday()
    if events:
        save_events_to_csv(events, csv_file_name)
        consolidated_events = consolidate_events(events)
        event_count = upload_events_to_sharepoint(consolidated_events)
        if event_count > 0:
            send_email_notification(event_count)
    else:
        print("No se encontraron eventos para el día anterior.")
