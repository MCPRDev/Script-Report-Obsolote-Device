import os
import base64
from datetime import datetime
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from bs4 import BeautifulSoup
import re
import csv
from dateutil.relativedelta import relativedelta

# Configuración inicial
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
print("=== CONFIGURACIÓN DE CREDENCIALES ===")
CREDENTIALS_FILE = input("Ingrese la ruta del archivo credentials.json: ").strip()

if not os.path.exists(CREDENTIALS_FILE):
    print(f"Archivo no encontrado: {CREDENTIALS_FILE}")
    exit()

TOKEN_FILE = 'token.json'

# Expresiones regulares para extraer la información
PATTERNS = {
    'usuario': re.compile(r'Usuario\s*:\s*([^\n]+)', re.IGNORECASE),
    'cpu_nuevo': re.compile(r'CPU\s*:\s*([^\n]+)', re.IGNORECASE),
    'monitor_nuevo': re.compile(r'Monitor\s*:\s*([^\n]+)', re.IGNORECASE),
    'teclado_nuevo': re.compile(r'Teclado\s*:\s*([^\n]+)', re.IGNORECASE),
    'cpu_viejo': re.compile(r'CPU a reponer\s*:\s*([^\n]+)', re.IGNORECASE),
    'monitor_viejo': re.compile(r'Monitor a reponer\s*:\s*([^\n]+)', re.IGNORECASE),
    'teclado_viejo': re.compile(r'Teclado a reponer\s*:\s*([^\n]+)', re.IGNORECASE)
}

def get_gmail_service():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    
    return build('gmail', 'v1', credentials=creds)

def extract_info_from_email(body, is_html=False):
    """
    Extrae una o varias filas de inventario de un correo.
    - Si detecta HTML: recorre todas las <table>, todas sus filas de datos y devuelve
      una lista de dicts (uno por cada row).
    - Si no es HTML o no hay tablas válidas: usa findall() con regex para capturar
      múltiples ocurrencias de cada patrón y devuelve un único dict.
    """
    if is_html:
        soup   = BeautifulSoup(body, 'html.parser')
        tables = soup.find_all('table')
        resultados = []

        for tbl in tables:
            rows = tbl.find_all('tr')
            # Salta tablas sin datos (menos de 2 filas)
            if len(rows) < 2:
                continue

            # Para cada fila de datos (desde la segunda)
            for row in rows[1:]:
                cols = row.find_all('td')
                if len(cols) < 7:
                    continue
                info = {
                    'usuario'        : cols[0].get_text(strip=True),
                    'cpu_nuevo'      : cols[1].get_text(strip=True),
                    'monitor_nuevo'  : cols[2].get_text(strip=True),
                    'teclado_nuevo'  : cols[3].get_text(strip=True),
                    'cpu_viejo'      : cols[4].get_text(strip=True),
                    'monitor_viejo'  : cols[5].get_text(strip=True),
                    'teclado_viejo'  : cols[6].get_text(strip=True),
                }
                resultados.append(info)

        if resultados:
            return resultados

    # —————————————————————————————
    # Fallback: texto plano con regex (captura múltiples ocurrencias)
    
    info = { key: [] for key in PATTERNS }
    for key, pattern in PATTERNS.items():
        matches = pattern.findall(body)
        info[key] = [m.strip() for m in matches]

    # Convertir listas en comas o cadena vacía
    single = { key: (",".join(vals) if vals else "") for key, vals in info.items() }
    return [ single ]

def process_emails_by_date_range(service, start_date, end_date, max_results=None):
    start_str = start_date.strftime('%Y/%m/%d')
    end_str   = end_date.strftime('%Y/%m/%d')
    query = f'"Remplazo por obsolescencia" after:{start_str} before:{end_str}'

    messages = []
    resp = service.users().messages().list(userId='me', q=query,
                                           maxResults=max_results or 500).execute()
    messages.extend(resp.get('messages', []))
    while 'nextPageToken' in resp and (not max_results or len(messages) < max_results):
        resp = service.users().messages().list(
            userId='me', q=query,
            maxResults=max_results or 500,
            pageToken=resp['nextPageToken']
        ).execute()
        messages.extend(resp.get('messages', []))
        if max_results and len(messages) >= max_results:
            messages = messages[:max_results]
            break

    resultados = []
    for m in messages:
        try:
            msg = service.users().messages().get(
                userId='me', id=m['id'], format='full'
            ).execute()
            fecha = datetime.fromtimestamp(int(msg['internalDate'])/1000)
            fecha_str = fecha.strftime('%Y-%m-%d %H:%M:%S')

            # Decodificar cuerpo
            body, is_html = "", False
            for part in msg.get('payload', {}).get('parts', []):
                mt = part.get('mimeType','')
                data = part.get('body',{}).get('data')
                if mt == 'text/html' and data:
                    body = base64.urlsafe_b64decode(data).decode('utf-8')
                    is_html = True
                    break
                if mt == 'text/plain' and data and not body:
                    body = base64.urlsafe_b64decode(data).decode('utf-8')

            # Extraer N registros de este correo
            registros = extract_info_from_email(body, is_html)
            for info in registros:
                info['fecha'] = fecha_str
                resultados.append(info)

        except Exception as e:
            print(f"⚠️ Error procesando mensaje {m['id']}: {e}")
            continue

    return resultados

def save_to_csv(data, start_date, end_date):
    # Formatear nombre del archivo
    start_str = start_date.strftime('%d-%m-%Y')
    end_str = end_date.strftime('%d-%m-%Y')
    filename = f"Equipos cambiados por obsolecencia {start_str} hasta {end_str}.csv"
    
    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = [
            'usuario',
            'cpu_nuevo',
            'monitor_nuevo',
            'teclado_nuevo',
            'cpu_viejo',
            'monitor_viejo',
            'teclado_viejo',
            'fecha'
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
    
    return filename

def process_in_batches(service, initial_date, final_date, months_per_batch=6):
    current_start = initial_date
    batch_number = 1
    
    while current_start < final_date:
        # Calcular fecha final del lote
        current_end = current_start + relativedelta(months=months_per_batch)
        if current_end > final_date:
            current_end = final_date
        
        print(f"\nProcesando lote {batch_number}: {current_start.strftime('%d/%m/%Y')} - {current_end.strftime('%d/%m/%Y')}")
        
        # Procesar correos en este rango de fechas
        resultados = process_emails_by_date_range(service, current_start, current_end)
        
        if resultados:
            # Guardar resultados en CSV
            filename = save_to_csv(resultados, current_start, current_end)
            print(f"Se guardaron {len(resultados)} registros en el archivo: {filename}")
        else:
            print("No se encontraron correos en este rango de fechas.")
        
        # Preparar siguiente lote
        current_start = current_end + relativedelta(days=1)
        batch_number += 1

def main():
    print("Iniciando proceso de extracción de información de Gmail...")
    service = get_gmail_service()
    print("Conectado a Gmail API")
    
    # Solicitar fechas al usuario
    print("\nConfiguración del período a analizar:")
    start_date = input("Fecha inicial (DD/MM/YYYY): ")
    end_date = input("Fecha final (DD/MM/YYYY): ")
    
    try:
        initial_date = datetime.strptime(start_date, '%d/%m/%Y')
        final_date = datetime.strptime(end_date, '%d/%m/%Y')
    except ValueError:
        print("Formato de fecha incorrecto. Use DD/MM/YYYY.")
        return
    
    # Solicitar tamaño del lote
    months_per_batch = input("Cantidad de meses por lote (por defecto 6): ")
    months_per_batch = int(months_per_batch) if months_per_batch.isdigit() else 6
    
    # Procesar en lotes
    process_in_batches(service, initial_date, final_date, months_per_batch)
    
    print("\nProceso completado.")

if __name__ == '__main__':
    main()