import win32evtlog
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

# Configuración de SharePoint
site_url = 'https://idtsas.sharepoint.com/sites/Bk-Empresas'
list_name = 'Visor de Eventos Construsol'
username = 'info@idtsas.com'
password = '1D2022++'

# Preguntar al usuario por el año y mes para filtrar los eventos
input_year = int(input("Ingrese el año de los eventos que desea traer (YYYY): "))
input_month = int(input("Ingrese el mes de los eventos que desea traer (MM): "))

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
log_type = 'System'  # Puedes cambiar a 'Application', 'Security', etc.
server = 'localhost'  # Servidor local
handle = win32evtlog.OpenEventLog(server, log_type)

flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ

events = win32evtlog.ReadEventLog(handle, flags, 0)

print(events)

for event in events:
    # Filtrar eventos por año y mes
    if event.TimeGenerated.year == input_year and event.TimeGenerated.month == input_month:
        event_data = {
            'NombreRegistro': 'Sistema',
            'IdEvento': str(event.EventID),  # Título del evento
            'Origen': str(event.SourceName),       # Fuente del evento
            'Categoria': str(event.EventCategory),       # Categoría
            'Detalle': str(event.StringInserts),    # Detalle
            'FechaGenerado': event.TimeGenerated.strftime('%Y-%m-%d %H:%M:%S'),  # Fecha
            'Nivel': str(event.EventType),
            'User': log_type,
            'CategoriaTarea': event.EventCategory
        }
        
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

print("Eventos cargados con éxito en SharePoint")
