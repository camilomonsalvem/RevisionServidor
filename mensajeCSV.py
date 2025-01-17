import wmi
import csv
import time
from datetime import datetime, timedelta, timezone
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

# Configuración de SharePoint
site_url = 'https://idtsas.sharepoint.com/sites/Bk-Empresas'
list_name = 'Visor de Eventos Construsol'
username = 'info@idtsas.com'
password = '1D2022++'

def format_wmi_time(wmi_time):
    """
    Convierte una fecha WMI en formato legible y ajusta a la zona horaria local.
    Formato de entrada: YYYYMMDDHHMMSS.mmmmmm±UTC_offset
    """
    try:
        utc_time = datetime.strptime(wmi_time.split('.')[0], '%Y%m%d%H%M%S').replace(tzinfo=timezone.utc)
        bogota_time = utc_time.astimezone(timezone(timedelta(hours=-5)))
        return bogota_time.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return "N/A"

def get_events_from_yesterday():
    """
    Obtiene todos los eventos del día anterior del log de Sistema y los guarda en un archivo CSV.
    """
    # Calcular el rango de tiempo del día anterior en la zona horaria de Bogotá (UTC-5)
    bogota_tz = timezone(timedelta(hours=-5))
    today_local = datetime.now(bogota_tz)
    yesterday_start_local = (today_local - timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    yesterday_end_local = (today_local - timedelta(days=1)).replace(hour=23, minute=59, second=59, microsecond=0)

    # Convertir el rango de Bogotá a UTC para la consulta WMI
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
    # Crear el archivo CSV
    headers = ["Nombre de Registro", "Origen", "ID", "Nivel", "Categoria de Tarea", "Registrado", "Equipo", "Usuario", "Mensaje", "EventRecordID"]
    try:
        with open(csv_file_name, mode='w', newline='', encoding='utf-8') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=headers)
            writer.writeheader()

            for event in events:
                # Construir el mensaje completo, uniendo "Message" y "StringInserts" si existen
                full_message = event.Message or ""
                if hasattr(event, "StringInserts") and event.StringInserts:
                    full_message += "\n" + "\n".join(event.StringInserts)

                # Remover saltos de línea adicionales para evitar problemas de formato en el CSV
                formatted_message = " ".join(full_message.splitlines())

                # Obtener el CategoryString si existe, manejar la lógica de "Categoría de Tarea"
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
                    "EventRecordID": event.RecordNumber  # Agregar EventRecordID
                }
                writer.writerow(row)

            print(f"Eventos guardados en el archivo CSV: {csv_file_name}")

    except Exception as e:
        print(f"Error al ejecutar la consulta o guardar en el archivo CSV: {e}")

def upload_events_to_sharepoint(events):
    """
    Sube los eventos a SharePoint.
    """
    try:
        # Autenticación con SharePoint
        ctx_auth = AuthenticationContext(site_url)
        if ctx_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(site_url, ctx_auth)
            print("Autenticación exitosa")
        else:
            print("Error al autenticar:", ctx_auth.get_last_error())
            return

        # Obtener la lista por nombre
        list_obj = ctx.web.lists.get_by_title(list_name)
        event_count = 0

        for event in events:
            # Construir el mensaje completo, uniendo "Message" y "StringInserts" si existen
            full_message = event.Message or ""
            if hasattr(event, "StringInserts") and event.StringInserts:
                full_message += "\n" + "\n".join(event.StringInserts)

            # Mantener el formato del mensaje para preservar saltos de línea y estructura
            formatted_message = full_message.strip()

            # Obtener el CategoryString si existe
            category_string = getattr(event, "CategoryString", "N/A")

            try:
                item_properties = {
                    'Title': event.SourceName,
                    'Id_x0020_del_x0020_Evento': str(event.EventCode),
                    'Origen': event.SourceName,
                    'Categoria_x0020_de_x0020_Tarea': category_string,
                    'Detalle': formatted_message,
                    'Fecha': format_wmi_time(event.TimeGenerated),
                    'Nivel': event.Type,
                    'User': event.User
                }
                list_obj.add_item(item_properties)
                ctx.execute_query()
                event_count += 1
                time.sleep(0.5)  # Evitar throttling
            except Exception as e:
                print(f"Error al cargar el evento {event.EventCode}: {e}")

        print(f"Total de eventos cargados a SharePoint: {event_count}")

    except Exception as e:
        print(f"Error al subir los eventos a SharePoint: {e}")

if __name__ == "__main__":
    csv_file_name = "Eventos_ayer.csv"  # Nombre del archivo CSV
    events = get_events_from_yesterday()
    if events:
        save_events_to_csv(events, csv_file_name)
    else:
        print("No se encontraron eventos para el día anterior.")
