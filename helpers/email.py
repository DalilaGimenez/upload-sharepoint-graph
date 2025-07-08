import os
import json
import logging
import requests
from dotenv import load_dotenv
from typing import Optional, List

# Carrega variáveis do .env
load_dotenv()

# Configura o logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def validate_env_vars(*vars: str) -> bool:
    for var in vars:
        if not os.getenv(var):
            logging.critical(f"[ERRO CRÍTICO] Variável de ambiente '{var}' não definida.")
            return False
    return True

def get_access_token() -> Optional[str]:
    tenant_id = os.getenv('TENANT_ID')
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')
    
    if not validate_env_vars('TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET'):
        return None

    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    payload = {
        'client_id': client_id,
        'scope': 'https://graph.microsoft.com/.default',
        'client_secret': client_secret,
        'grant_type': 'client_credentials'
    }
    
    try:
        response = requests.post(url, headers=headers, data=payload)
        response.raise_for_status()
        return response.json().get('access_token')
    except requests.exceptions.HTTPError as e:
        logging.error(f"Erro HTTP ao obter token: {e}")
        logging.error(f"Resposta: {e.response.text}")
    except requests.exceptions.RequestException as e:
        logging.error(f"Erro de conexão ao obter token: {e}")
    return None

def send_email(access_token: str, subject: str, body: str, body_content: str, sender: str, recipients: Optional[List[str]] = None, is_html: bool = True):
    sender_email = os.getenv('SENDER_EMAIL')
    if not sender_email:
        logging.error("Variável SENDER_EMAIL não está definida no ambiente.")
        return

    if recipients is None:
        recipients_str = os.getenv('RECIPIENT_EMAILS', '')
        recipients = [email.strip() for email in recipients_str.split(',') if email.strip()]

    if not recipients:
        logging.error("Nenhum destinatário definido para envio de e-mail.")
        return

    to_recipients = [{'emailAddress': {'address': email}} for email in recipients]
    content_type = 'HTML' if is_html else 'Text'

    email_payload = {
        'message': {
            'subject': subject,
            'body': {
                'contentType': content_type,
                'content': body_content
            },
            'toRecipients': to_recipients
        },
        'saveToSentItems': True
    }

    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    try:
        response = requests.post(
            f'https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail',
            headers=headers,
            json=email_payload
        )

        if response.status_code == 202:
            logging.info("✅ E-mail enviado com sucesso!")
        else:
            logging.error(f"[ERRO] Falha no envio. Status: {response.status_code}")
            logging.error(f"[ERRO] Detalhes: {response.text}")
    except requests.exceptions.RequestException as e:
        logging.error(f"[ERRO DE CONEXÃO] Falha ao conectar ao Graph: {e}")

if __name__ == '__main__':
    try:
        logging.info("Iniciando processo de envio de e-mail...")
        token = get_access_token()
        if token:
            assunto_email = "Log de Upload para o SharePoint"
            corpo_email = "<h1>Relatório de Upload</h1><p>Segue o log em anexo.</p>"
            send_email(token, assunto_email, corpo_email)
        else:
            logging.error("Token de acesso não obtido. Abortando envio.")
    except Exception as e:
        logging.critical(f"Erro fatal inesperado: {e}", exc_info=True)
