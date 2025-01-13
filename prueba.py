import win32evtlog
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

# Diccionarios para traducir valores a texto
event_type_mapping = {
    1: "Error",
    2: "Advertencia",
    4: "Información",
    8: "Auditoría exitosa",
    16: "Error de auditoría"
}

task_category_mapping = {
    0: "Ninguna",
    1: "Operación",
    2: "Configuración",
    3: "Auditoría",
    4: "Tarea personalizada"
}

operation_code_mapping = {
    16: "Inicialización",
    32: "Actualización",
    48: "Eliminación",
    64: "Reinicio"
}

keywords_mapping = {
    0x8000000000000000: "Clave de auditoría exitosa",
    0x4000000000000000: "Clave de auditoría fallida",
    0x2000000000000000: "Clave de inicio del sistema",
    0x1000000000000000: "Clave de cierre del sistema",
    0x0800000000000000: "Clave de información general",
    0x8080000000000000: "Clave combinada de éxito y general"
}

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

                    # Convertir palabras clave en texto
                    keywords_text = keywords_mapping.get(event.EventType, "Desconocido")

                    file.write("========== EVENTO ==========\n")
                    file.write(f"Nombre del registro: {log_type}\n")  # Nombre del registro
                    file.write(f"Event ID (Base): {base_event_id}\n")  # Event ID que coincide con el Visor de Eventos
                    file.write(f"Event ID (Completo): {event.EventID}\n")
                    file.write(f"Source: {event.SourceName}\n")
                    file.write(f"Código de operación: {event.EventID}\n")
                    file.write(f"Categoría de tarea: {event.EventCategory}\n")
                    file.write(f"Nivel: {event.EventType}\n")
                    file.write(f"Palabras clave: {keywords_text}\n")
                    file.write(f"Time Generated: {event.TimeGenerated}\n")
                    file.write(f"Time Written: {event.TimeWritten}\n")
                    file.write(f"Computer Name: {event.ComputerName}\n")
                    file.write(f"Record Number: {event.RecordNumber}\n")
                    file.write(f"Data (Raw): {event.Data}\n")
                    file.write(f"String Inserts: {' | '.join(event.StringInserts) if event.StringInserts else 'No details'}\n")
                    file.write(f"User SID: {event.Sid if hasattr(event, 'Sid') else 'Not available'}\n")
                    file.write(f"Reserved: {event.Reserved}\n")
                    file.write(f"RecordNumber: {event.ClosingRecordNumber}\n")
                    file.write(f"ReservedFlags: {event.ReservedFlags}\n")
                    file.write("===========================\n\n")
except Exception as e:
    print(f"Error al procesar eventos: {e}")

print(f"Procesamiento completado. Los datos del día anterior se han guardado en {output_file_name}.")
