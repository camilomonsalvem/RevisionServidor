import win32evtlog
import win32evtlogutil
from datetime import datetime, timedelta

# Configuración del visor de eventos
log_type = 'System'  # Cambiar según necesidad (por ejemplo, 'Application', 'Security')
server = 'localhost'  # Servidor local
handle = win32evtlog.OpenEventLog(server, log_type)

flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ

# Calcular la fecha del día anterior
yesterday = datetime.now() - timedelta(days=1)
input_year = yesterday.year
input_month = yesterday.month
input_day = yesterday.day

# Nombre del archivo de salida
output_file_name = f"Eventos_{log_type}_DiaAnterior.txt"

print(f"Extrayendo datos del visor de eventos del día anterior... Guardando en {output_file_name}\n")

try:
    with open(output_file_name, mode='w', encoding='utf-8') as file:
        while True:
            events = win32evtlog.ReadEventLog(handle, flags, 0)
            if not events:
                break
            for event in events:
                # Filtrar eventos por el día anterior
                if (
                    event.TimeGenerated.year == input_year and
                    event.TimeGenerated.month == input_month and
                    event.TimeGenerated.day == input_day
                ):
                    # Obtener el Event ID Base (ignorar qualifiers)
                    base_event_id = event.EventID & 0xFFFF

                    # Obtener la descripción del evento
                    try:
                        description = win32evtlogutil.SafeFormatMessage(event, log_type)
                    except Exception as e:
                        description = f"Error al obtener la descripción: {e}"

                    # Si la descripción está vacía pero hay inserciones, formatear manualmente
                    if not description.strip() and event.StringInserts:
                        description = f"Descripción generada a partir de cadenas insertadas: {', '.join(event.StringInserts)}"

                    # Agregar los datos del evento al archivo
                    file.write("========== EVENTO ==========\n")
                    file.write(f"Nombre del registro: {log_type}\n")  # Nombre del registro
                    file.write(f"Event ID (Base): {base_event_id}\n")  # Event ID que coincide con el Visor de Eventos
                    file.write(f"Event ID (Completo): {event.EventID}\n")
                    file.write(f"Source: {event.SourceName}\n")
                    file.write(f"Nivel: {event.EventType}\n")
                    file.write(f"Time Generated: {event.TimeGenerated}\n")
                    file.write(f"Descripción: {description.strip()}\n")  # Descripción completa
                    file.write(f"Computer Name: {event.ComputerName}\n")
                    file.write(f"Record Number: {event.RecordNumber}\n")
                    file.write(f"String Inserts: {' | '.join(event.StringInserts) if event.StringInserts else 'No details'}\n")
                    file.write("===========================\n\n")
except Exception as e:
    print(f"Error al procesar eventos: {e}")

print(f"Procesamiento completado. Los datos del día anterior se han guardado en {output_file_name}.")
