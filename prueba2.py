import win32evtlog
import win32evtlogutil
import win32security
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

# Mapeo de niveles de eventos
event_level_mapping = {
    1: "Crítico",
    2: "Error",
    3: "Advertencia",
    4: "Información",
    5: "Depuración"
}

# Mapeo de códigos de operación comunes (si es aplicable)
operation_code_mapping = {
    0: "Información",
    1: "Inicio",
    2: "Parada",
    3: "Éxito",
    4: "Error",
    5: "Depuración"
}

def get_username_from_sid(sid):
    try:
        account_name, domain_name, account_type = win32security.LookupAccountSid(None, sid)
        return f"{domain_name}\\{account_name}"
    except Exception:
        return "Usuario desconocido"

def get_operation_code(event):
    """Intenta traducir el código de operación."""
    try:
        # Si es posible, intenta obtener el texto directamente
        operation_code = win32evtlogutil.SafeFormatMessage(event, event.SourceName)
        return operation_code.strip()
    except Exception:
        # Si no, devuelve el código numérico (si existe)
        return operation_code_mapping.get(event.EventCategory, f"Desconocido ({event.EventCategory})")

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
                    except Exception:
                        description = "<No se encontró la descripción del evento>"

                    # Si la descripción está vacía pero hay inserciones, formatear manualmente
                    if "<No se encontró la descripción del evento>" in description and event.StringInserts:
                        insertion_text = ", ".join(event.StringInserts)
                        description += f" Contiene las siguientes cadenas de inserción: {insertion_text}"

                    # Extraer el nivel del evento desde el campo Level
                    level = getattr(event, "EventType", None)  # Extraer nivel del evento
                    event_level_text = event_level_mapping.get(level, f"Desconocido ({level})")

                    # Obtener el usuario
                    user = get_username_from_sid(event.Sid) if hasattr(event, 'Sid') and event.Sid else "No disponible"

                    # Traducir el código de operación
                    operation_code = get_operation_code(event)

                    # Agregar los datos del evento al archivo
                    file.write("========== EVENTO ==========\n")
                    file.write(f"Descripción: {description.strip()}\n")  # Descripción completa
                    file.write(f"Nombre del registro: {log_type}\n")  # Nombre del registro
                    file.write(f"Source: {event.SourceName}\n")  # Origen del evento
                    file.write(f"Event ID (Base): {base_event_id}\n")  # Event ID que coincide con el Visor de Eventos
                    file.write(f"Nivel: {event_level_text}\n")  # Nivel en texto y número
                    file.write(f"Usuario: {user}\n")  # Usuario
                    file.write(f"Código de Operación: {operation_code}\n")  # Código de operación traducido
                    file.write(f"Time Generated: {event.TimeGenerated}\n")
                    file.write(f"Categoría de tarea: {event.EventCategory}\n")
                    file.write(f"Equipo: {event.ComputerName}\n")
                    file.write("===========================\n\n")
except Exception as e:
    print(f"Error al procesar eventos: {e}")

print(f"Procesamiento completado. Los datos del día anterior se han guardado en {output_file_name}.")