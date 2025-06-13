import sys
import subprocess
import ctypes
import socket
import os

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def log_message(message):
    log_path = 'C:\\hostname_change.log'
    with open(log_path, 'a', encoding='utf-8') as f:
        f.write(f"{message}\n")

def change_hostname_silently(new_hostname, domain_user, domain_password):
    try:
        if not is_admin():
            log_message("Error: El script no tiene privilegios de administrador.")
            return 1

        current = socket.gethostname()
        if current.lower() == new_hostname.lower():
            log_message(f"El hostname ya es '{new_hostname}'. No se requiere cambio.")
            return 0

        # Construir PowerShell en una sola línea
        powershell_command = (
            f"$securePassword = ConvertTo-SecureString '{domain_password}' -AsPlainText -Force; "
            f"$cred = New-Object System.Management.Automation.PSCredential('{domain_user}', $securePassword); "
            f"Rename-Computer -NewName '{new_hostname}' -DomainCredential $cred -Force"
        )

        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", powershell_command],
            capture_output=True,
            text=True,
            shell=True
        )

        if result.returncode == 0:
            log_message(f"Éxito: El nombre del equipo fue cambiado a '{new_hostname}' (sin reinicio).")
            return 0
        else:
            log_message(f"Error al cambiar nombre: {result.stderr.strip()}")
            log_message(f"Salida PowerShell: {result.stdout.strip()}")
            log_message(f"hostname: {nuevo_hostname}")
            log_message(f"usuario: {usuario_dominio}")
            log_message(f"password: {contraseña}")
            return 1

    except Exception as e:
        log_message(f"Excepción: {str(e)}")
        return 1


if __name__ == "__main__":
    if len(sys.argv) != 4:
        log_message("Uso incorrecto. Debe proporcionar: <nuevo_hostname> <usuario_dominio> <contraseña>")
        sys.exit(1)

    nuevo_hostname = sys.argv[1]
    usuario_dominio = sys.argv[2]
    contraseña = sys.argv[3]

    sys.exit(change_hostname_silently(nuevo_hostname, usuario_dominio, contraseña))
