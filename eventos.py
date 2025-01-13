import win32evtlog
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from datetime import datetime, timedelta
import csv
import time

# Configuración de SharePoint
site_url = 'https://idtsas.sharepoint.com/sites/Bk-Empresas'
list_name = 'Visor de Eventos Construsol'
username = 'info@idtsas.com'
password = '1D2022++'

# Calcular la fecha del día anterior
yesterday = datetime.now() - timedelta(days=1)
input_year = yesterday.year
input_month = yesterday.month
input_day = yesterday.day

# Autenticación
ctx_auth = AuthenticationContext(site_url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(site_url, ctx_auth)
    print("Autenticación exitosa")
else:
    print("Error al autenticar:", ctx_auth.get_last_error())
    exit()

# Obtener la lista por nombre
list_obj = ctx.web.lists.get_by_title(list_name)

# Capturar eventos del visor de eventos de Windows
log_type = 'System'  # Cambiar según necesidad
server = 'localhost'
handle = win32evtlog.OpenEventLog(server, log_type)

flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ

# Listado para guardar todos los eventos filtrados (para CSV)
all_events = []

# Diccionario para almacenar eventos únicos y sus conteos
event_summary = {}

try:
    print("Procesando eventos...")
    while True:
        events = win32evtlog.ReadEventLog(handle, flags, 0)
        if not events:  # Salir si no hay más eventos
            break
        for event in events:
            # Filtrar eventos por el día anterior
            if (
                event.TimeGenerated.year == input_year and
                event.TimeGenerated.month == input_month and
                event.TimeGenerated.day == input_day
            ):
                # Agregar evento al listado completo (para CSV)
                all_events.append({
                    'IdEvento': str(event.EventID),
                    'Origen': str(event.SourceName),
                    'Categoria': str(event.EventCategory),
                    'Detalle': " | ".join(event.StringInserts) if event.StringInserts else "Sin detalles",
                    'FechaGenerado': event.TimeGenerated.strftime('%Y-%m-%d %H:%M:%S'),
                    'Nivel': str(event.EventType),
                    'User': log_type,
                })

                # Crear una clave única para el evento
                event_key = (event.EventID, event.SourceName)

                # Incrementar el conteo en el diccionario
                if event_key in event_summary:
                    event_summary[event_key]['NoEventos'] += 1
                else:
                    event_summary[event_key] = {
                        'NombreRegistro': 'Sistema',
                        'IdEvento': str(event.EventID),
                        'Origen': str(event.SourceName),
                        'Categoria': str(event.EventCategory),
                        'Detalle': " | ".join(event.StringInserts) if event.StringInserts else "Sin detalles",
                        'FechaGenerado': event.TimeGenerated.strftime('%Y-%m-%d %H:%M:%S'),
                        'Nivel': str(event.EventType),
                        'User': log_type,
                        'NoEventos': 1
                    }
except Exception as e:
    print(f"Error al procesar eventos: {e}")

# Subir eventos únicos a SharePoint
event_count = 0
for event_key, event_data in event_summary.items():
    try:
        # Recortar texto largo para SharePoint
        event_data['Detalle'] = event_data['Detalle'][:255]

        # Crear un nuevo elemento en la lista de SharePoint
        item_properties = {
            'Title': event_data['NombreRegistro'],
            'Id_x0020_del_x0020_Evento': event_data['IdEvento'],
            'Origen': event_data['Origen'],
            'Categoria_x0020_de_x0020_Tarea': event_data['Categoria'],
            'Detalle': event_data['Detalle'],
            'Fecha': event_data['FechaGenerado'],
            'Nivel': {
                1: "Error",
                2: "Advertencia",
                4: "Información",
                8: "Auditoría exitosa",
                16: "Error de auditoría"
            }.get(int(event_data['Nivel']), "Desconocido"),
            'User': event_data['User'],
            'No_x0020_de_x0020_Eventos': event_data['NoEventos']
        }
        new_item = list_obj.add_item(item_properties)
        ctx.execute_query()

        # Incrementar el contador de eventos cargados
        event_count += 1
        time.sleep(0.5)  # Reducir throttling
    except Exception as e:
        print(f"Error al cargar el evento {event_data['IdEvento']}: {e}")

# Crear archivo CSV con todos los eventos filtrados
csv_file_name = f"Eventos_{yesterday.strftime('%Y-%m-%d')}.csv"
csv_headers = ['IdEvento', 'Origen', 'Categoria', 'Detalle', 'FechaGenerado', 'Nivel', 'User']

try:
    with open(csv_file_name, mode='w', newline='', encoding='utf-8') as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=csv_headers)
        writer.writeheader()
        writer.writerows(all_events)
    print(f"Archivo CSV creado: {csv_file_name}")
except Exception as e:
    print(f"Error al crear el archivo CSV: {e}")

# Imprimir el total de eventos cargados al final
print(f"Procesamiento completado. Total de eventos cargados: {event_count}")
