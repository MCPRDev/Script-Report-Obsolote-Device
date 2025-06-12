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
CREDENTIALS_FILE = 'credentials.json'
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
    info = {
        'usuario': '',
        'cpu_nuevo': '',
        'monitor_nuevo': '',
        'teclado_nuevo': '',
        'cpu_viejo': '',
        'monitor_viejo': '',
        'teclado_viejo': '',
        'fecha': ''
    }

    if is_html:
        soup = BeautifulSoup(body, 'html.parser')
        table = soup.find('table')
        if table:
            rows = table.find_all('tr')
            if len(rows) > 1:  # Tiene al menos una fila de datos
                headers = [th.get_text(strip=True).lower() for th in rows[0].find_all(['td', 'th'])]
                values = [td.get_text(strip=True) for td in rows[1].find_all('td')]

                if len(headers) == 7 and len(values) == 7:
                    info['usuario'] = values[0]
                    info['cpu_nuevo'] = values[1]
                    info['monitor_nuevo'] = values[2]
                    info['teclado_nuevo'] = values[3]
                    info['cpu_viejo'] = values[4]
                    info['monitor_viejo'] = values[5]
                    info['teclado_viejo'] = values[6]
                    return info

    # Fallback: texto plano con regex
    for key, pattern in PATTERNS.items():
        match = pattern.search(body)
        if match:
            info[key] = match.group(1).strip()

    return info

def process_emails_by_date_range(service, start_date, end_date, max_results=None):
    # Formatear fechas para la consulta Gmail
    start_date_str = start_date.strftime('%Y/%m/%d')
    end_date_str = end_date.strftime('%Y/%m/%d')
    
    # Buscar correos con términos específicos y en el rango de fechas
    query = f'"Remplazo por obsolescencia" after:{start_date_str} before:{end_date_str}'
    
    print(f"Buscando correos entre {start_date_str} y {end_date_str}...")
    
    response = service.users().messages().list(
        userId='me',
        q=query,
        maxResults=max_results if max_results else 500,
        includeSpamTrash=False
    ).execute()
    
    messages = []
    if 'messages' in response:
        messages.extend(response['messages'])
    
    while 'nextPageToken' in response and (not max_results or len(messages) < max_results):
        page_token = response['nextPageToken']
        response = service.users().messages().list(
            userId='me',
            q=query,
            maxResults=max_results if max_results else 500,
            includeSpamTrash=False,
            pageToken=page_token
        ).execute()
        messages.extend(response['messages'])
        
        if max_results and len(messages) >= max_results:
            messages = messages[:max_results]
            break
    
    # Procesar cada mensaje
    resultados = []
    for message in messages:
        msg = service.users().messages().get(
            userId='me',
            id=message['id'],
            format='full'
        ).execute()
        
        # Extraer fecha
        fecha = datetime.fromtimestamp(int(msg['internalDate'])/1000)
        fecha_str = fecha.strftime('%Y-%m-%d %H:%M:%S')
        
        # Extraer cuerpo del mensaje
        body = ''

        body = ''
        is_html = False

        if 'parts' in msg['payload']:
            for part in msg['payload']['parts']:
                mime = part['mimeType']
                if mime == 'text/html':
                    data = part['body'].get('data')
                    if data:
                        body = base64.urlsafe_b64decode(data).decode('utf-8')
                        is_html = True
                        break
                elif mime == 'text/plain' and not body:
                    data = part['body'].get('data')
                    if data:
                        body = base64.urlsafe_b64decode(data).decode('utf-8')
        else:
            if 'body' in msg['payload'] and 'data' in msg['payload']['body']:
                data = msg['payload']['body']['data']
                body = base64.urlsafe_b64decode(data).decode('utf-8')

        info = extract_info_from_email(body, is_html=is_html)


        info['fecha'] = fecha_str
        resultados.append(info)
    
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