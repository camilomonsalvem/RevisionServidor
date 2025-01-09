import win32evtlog
from office365.sharepoint.listitems.listitem import ListItem
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

# Set the URL of the SharePoint site and the list name
site_url = 'https://idtsas.sharepoint.com/sites/Bk-Empresas'
list_name = 'Visor de Eventos Construsol'
username = 'info@idtsas.com'
password = '1D2022++'


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

# Crear un nuevo elemento de lista
item_properties = {'Title': 'Prueba Video'}
new_item = list_obj.add_item(item_properties)
ctx.execute_query()

print("Elemento creado con éxito")