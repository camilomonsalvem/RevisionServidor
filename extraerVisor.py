import win32evtlog
import csv

# Configuración
log_type = 'System'  # Cambiar a 'Application', 'Security', etc.
output_csv = 'visor_de_eventos_detallado.csv'
server = 'localhost'  # Servidor local

# Abrir el log del visor de eventos
handle = win32evtlog.OpenEventLog(server, log_type)
flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ

# Leer eventos
events = win32evtlog.ReadEventLog(handle, flags, 0)

# Crear archivo CSV con los campos requeridos
with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)

    # Escribir encabezados
    writer.writerow([
        "Nombre de registro", "Origen", "Id. del evento", "Nivel",
        "Usuario", "Código de operación", "Registrado", "Categoría de tarea",
        "Palabras clave", "Equipo"
    ])

    # Procesar eventos
    for event in events:
        try:
            # Extraer detalles del evento
            nombre_registro = log_type  # Nombre del registro (System, Application, etc.)
            origen = event.SourceName
            id_evento = event.EventID
            nivel = {
                1: "Error",
                2: "Advertencia",
                4: "Información",
                8: "Auditoría exitosa",
                16: "Error de auditoría"
            }.get(event.EventType, "Desconocido")
            usuario = "SYSTEM" if event.Sid is None else str(event.Sid)
            codigo_operacion = "Información"  # No disponible en win32evtlog, ajusta si necesario
            registrado = event.TimeGenerated.strftime('%Y-%m-%d %H:%M:%S')
            categoria_tarea = event.EventCategory
            palabras_clave = "Clásico"  # No disponible directamente, ajustar según necesidad
            equipo = event.ComputerName

            # Escribir fila en el CSV
            writer.writerow([
                nombre_registro, origen, id_evento, nivel,
                usuario, codigo_operacion, registrado, categoria_tarea,
                palabras_clave, equipo
            ])
        except Exception as e:
            print(f"Error al procesar un evento: {e}")
            continue

print(f"Información exportada con éxito a {output_csv}")
