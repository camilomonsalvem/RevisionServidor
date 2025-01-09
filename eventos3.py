import win32evtlog
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
import time

# Configuración de SharePoint
site_url = 'https://idtsas.sharepoint.com/sites/Bk-Empresas'
list_name = 'Visor de Eventos Construsol'
username = 'info@idtsas.com'
password = '1D2022++'

# Preguntar al usuario por el año, mes y día para filtrar los eventos
input_year = int(input("Ingrese el año de los eventos que desea traer (YYYY): "))
input_month = int(input("Ingrese el mes de los eventos que desea traer (MM): "))
input_day = int(input("Ingrese el día de los eventos que desea traer (DD): "))

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

# Contador de eventos cargados
event_count = 0

while True:
    events = win32evtlog.ReadEventLog(handle, flags, 0)
    if not events:  # Salir si no hay más eventos
        break
    for event in events:
        # Filtrar eventos por año, mes y día
        if (
            event.TimeGenerated.year == input_year and
            event.TimeGenerated.month == input_month and
            event.TimeGenerated.day == input_day
        ):
            try:
                event_data = {
                    'NombreRegistro': 'Sistema',
                    'IdEvento': str(event.EventID),
                    'Origen': str(event.SourceName),
                    'Categoria': str(event.EventCategory),
                    'Detalle': " | ".join(event.StringInserts) if event.StringInserts else "Sin detalles",
                    'FechaGenerado': event.TimeGenerated.strftime('%Y-%m-%d %H:%M:%S'),
                    'Nivel': str(event.EventType),
                    'User': log_type,
                }
                
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
                    }.get(event.EventType, "Desconocido"),
                    'User': event_data['User']
                }
                new_item = list_obj.add_item(item_properties)
                ctx.execute_query()

                # Incrementar el contador de eventos cargados
                event_count += 1
                time.sleep(0.5)  # Reducir throttling
            except Exception as e:
                print(f"Error al cargar el evento {event.EventID}: {e}")

# Imprimir el total de eventos cargados al final
print(f"Procesamiento completado. Total de eventos cargados: {event_count}")
