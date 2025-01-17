import wmi
from datetime import datetime, timedelta, timezone

def format_wmi_time(wmi_time):
    """
    Convierte una fecha WMI en formato legible y ajusta a la zona horaria local.
    Formato de entrada: YYYYMMDDHHMMSS.mmmmmm±UTC_offset
    """
    try:
        # Convertir a objeto datetime en UTC
        utc_time = datetime.strptime(wmi_time.split('.')[0], '%Y%m%d%H%M%S').replace(tzinfo=timezone.utc)
        # Convertir a hora local (Bogotá, UTC-5)
        bogota_time = utc_time.astimezone(timezone(timedelta(hours=-5)))
        return bogota_time.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return wmi_time  # Retorna la fecha original si ocurre un error

def main():
    # Calcular rango del día anterior en hora local (Bogotá, UTC-5)
    bogota_tz = timezone(timedelta(hours=-5))
    today_local = datetime.now(bogota_tz)
    yesterday_start_local = (today_local - timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    yesterday_end_local = (today_local - timedelta(days=1)).replace(hour=23, minute=59, second=59, microsecond=0)

    # Convertir el rango de Bogotá a UTC para WMI
    yesterday_start_utc = yesterday_start_local.astimezone(timezone.utc)
    yesterday_end_utc = yesterday_end_local.astimezone(timezone.utc)

    # Convertir a formato WMI (YYYYMMDDHHMMSS.mmmmmm±UTC_offset)
    start_str = yesterday_start_utc.strftime('%Y%m%d%H%M%S.000000+000')
    end_str = yesterday_end_utc.strftime('%Y%m%d%H%M%S.999999+000')

    # Mostrar rango de tiempo para depuración
    print(f"Buscando eventos entre '{start_str}' y '{end_str}' (en UTC)")

    # Initialize WMI objects and query.
    wmi_o = wmi.WMI('.')
    wql = (f"SELECT * FROM Win32_NTLogEvent WHERE Logfile='System' "
           f"AND TimeGenerated >= '{start_str}' AND TimeGenerated <= '{end_str}'")
    
    # Query WMI object.
    try:
        wql_r = wmi_o.query(wql)
    except Exception as e:
        print(f"Error al ejecutar la consulta WQL: {e}")
        return None

    # Si hay resultados, ordenar los eventos por TimeGenerated de forma descendente
    if wql_r:
        # Ordenar los eventos por TimeGenerated en orden descendente
        sorted_events = sorted(wql_r, key=lambda x: x.TimeGenerated, reverse=True)
        events = []
        for event in sorted_events:
            # Lógica para determinar la categoría de tarea
            if event.Category == 0:
                task_category = "Ninguno"
            else:
                task_category = getattr(event, "CategoryString", None) or event.Category
            events.append({
                "Mensaje": event.Message,
                "Nombre del Registro": event.Logfile,
                "Origen": event.SourceName,
                "Id": event.EventCode,
                "Nivel": event.Type,
                "Usuario": event.User,
                "TimeGenerated": format_wmi_time(event.TimeGenerated),
                "Categoria de Tarea": task_category,
                "Equipo": event.ComputerName,
                "RecordNumber": event.RecordNumber,
            })
        return events
    else:
        print("No se encontraron eventos en el log de Sistema.")
        return None


if __name__ == '__main__':
    events = main()
    if events:
        with open("wmi_events.txt", "w", encoding="utf-8") as file:
            file.write("Eventos encontrados en el log de Sistema del día anterior (ordenados por fecha descendente):\n")
            for event in events:
                file.write("--------------------------------------------\n")
                for key, value in event.items():
                    file.write(f"{key}: {value}\n")
        print(f"Se encontraron {len(events)} eventos. Guardados en el archivo 'wmi_events.txt'.")
    else:
        print("No se encontraron eventos en el log de Sistema.")
