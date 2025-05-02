# Sistema de Automatización de Monitoreo de Eventos y Backups

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Descripción

Este proyecto implementa una solución completa para la automatización del monitoreo de eventos críticos en servidores empresariales y la supervisión de respaldos (backups) de equipos. Utiliza tecnologías como Python, Power Automate, SharePoint y PowerApps para crear un sistema integrado que mejora significativamente la eficiencia del equipo de soporte técnico.

## Requisitos previos

- Python 3.6+
- Acceso a Microsoft 365 (SharePoint, Power Automate)
- Permisos administrativos en los servidores a monitorear
- Las siguientes bibliotecas de Python:
  - wmi
  - office365
  - python-dotenv
  - psutil

## Estructura del proyecto

```
.
├── .env.secure              # Variables de entorno seguras (credenciales)
├── .env                     # Variables de entorno configurables
├── serverRevision.py        # Script principal para la revisión de servidores
├── serverRevision.spec      # Configuración de empaquetado con PyInstaller
├── deleteEvents.py          # Script para limpiar eventos procesados
├── email_template.html      # Plantilla para envío de notificaciones
```

## Instalación

1. Clona el repositorio:
   ```bash
   git clone https://github.com/tu-usuario/monitoreo-automatizado.git
   cd monitoreo-automatizado
   ```

2. Instala las dependencias:
   ```bash
   pip install -r requirements.txt
   ```

3. Configura tus variables de entorno:
   - Crea un archivo `.env.secure` con tus credenciales (sigue el formato en `.env.example`)
   - Ajusta el archivo `.env` con la configuración específica de tu empresa

## Configuración

### Variables de entorno

El sistema utiliza dos archivos de configuración:

1. **`.env.secure`** (información sensible):
   ```
   idt_username=usuario_sharepoint
   idt_password=contraseña_sharepoint
   site_url=https://tu-tenant.sharepoint.com/sites/tu-sitio
   site_url_soporte=https://tu-tenant.sharepoint.com/sites/soporte
   smtp_server=smtp.office365.com
   smtp_port=587
   email_sender=correo@tuempresa.com
   email_password=contraseña_correo
   ```

2. **`.env`** (configuración general):
   ```
   empresa=Nombre_de_la_Empresa
   ruta_archivos_guardados=C:/Ruta/A/Archivos
   ruta_archivos_guardados_system_state=C:/Ruta/SystemState
   email_recipients=soporte@tuempresa.com,admin@tuempresa.com
   list_name_visor_eventos=Nombre_Lista_Visor
   list_name_inventario=Nombre_Lista_Inventario
   list_name_chequeo_servidor=Nombre_Lista_Chequeo
   sharepoint_folder=/sites/SoporteTecnico/Documentos/Empresa/VisorEventos
   powerapps_app_link=https://apps.powerapps.com/play/tu-app-id
   ```

### SharePoint

El sistema requiere tres listas en SharePoint:

1. **Lista de Visor de Eventos** con las siguientes columnas:
   - Title (texto)
   - Id_del_Evento (texto)
   - Origen (texto)
   - Categoria de Tarea (texto)
   - Detalle (texto múltiple)
   - Fecha (fecha/hora)
   - Nivel (texto)
   - User (texto)
   - No_de_Eventos (número)
   - ID Chequeo ServidorId (búsqueda)

2. **Lista de Inventario PC** con las columnas:
   - Title (nombre del equipo)
   - Otras columnas de inventario según necesidad

3. **Lista de Chequeo Servidor** con las columnas:
   - Title (texto)
   - Fecha (fecha/hora)
   - Ruta_Archivos_Guardados (texto)
   - Procesador_% (porcentaje)
   - Memoria_% (porcentaje)
   - Tamaño_Discos (número)
   - Espacio_Libre_Disco_C (número)
   - Espacio_Libre_Disco_D (número)
   - Espacio_Libre_Disco_I (número)
   - Version_de_la_Actualizacion (texto)
   - Virus_Detectados (número)
   - Sistemas_Operativo (texto)
   - ServidorId (búsqueda)

## Uso

### Monitoreo de eventos

1. Ejecuta el archivo `serverRevision.exe` en el servidor que deseas monitorear:
   ```bash
   ./serverRevision.exe
   ```

2. El proceso realizará automáticamente:
   - Recolección de eventos del visor de Windows (Application, System, Security, Setup)
   - Resumen y agrupación de eventos similares
   - Subida de eventos a SharePoint
   - Envío de notificación por correo
   - Limpieza del visor de eventos

3. Accede a la aplicación PowerApps utilizando el enlace proporcionado en el correo de notificación para visualizar los eventos.

### Limpieza manual de eventos

Si necesitas limpiar manualmente los eventos del sistema, puedes utilizar el script `deleteEvents.py`:

```bash
python deleteEvents.py
```

## Programación de tareas

Para una automatización completa, se recomienda programar la ejecución diaria del ejecutable utilizando el Programador de tareas de Windows:

1. Abre el Programador de tareas de Windows
2. Crea una tarea básica
3. Establece un nombre (ej. "Monitoreo de Eventos")
4. Selecciona frecuencia diaria
5. Selecciona "Iniciar un programa"
6. Navega hasta la ubicación de `serverRevision.exe`
7. Finaliza el asistente

### Registros de ejecución

El sistema genera un archivo `app.log` que contiene información detallada sobre cada ejecución. Consulta este archivo en caso de errores.

### Errores comunes

- **Error de autenticación en SharePoint**: Verifica tus credenciales en `.env.secure`
- **Error al enviar correo**: Comprueba la configuración SMTP
- **No se encuentran eventos**: Verifica los permisos administrativos en el servidor
