import subprocess

def clear_event_log(log_name):
    """
    Elimina todos los eventos de un registro específico en el Visor de Eventos.
    """
    try:
        print(f"Intentando eliminar eventos del registro: {log_name}")
        subprocess.run(f"wevtutil cl \"{log_name}\"", shell=True, check=True, text=True)
        print(f"Eventos del registro '{log_name}' eliminados correctamente.")
    except subprocess.CalledProcessError as e:
        print(f"Error al eliminar los eventos del registro '{log_name}': {e}")
    except Exception as e:
        print(f"Error inesperado al procesar el registro '{log_name}': {e}")

if __name__ == "__main__":
    # Vaciar registros de "Application" y "System"
    clear_event_log("Application")
    clear_event_log("System")