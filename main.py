# main.py

from helpers.auth import get_access_token
from helpers.sharepoint import get_site_id, get_drive_id, upload_files
from helpers.email import send_email
from dotenv import load_dotenv
import os, sys
from datetime import datetime

def get_base_path():
    """ Retorna o caminho base para o executável ou script. """
    if getattr(sys, 'frozen', False):
        # Se estiver rodando como um executável (congelado pelo PyInstaller)
        return os.path.dirname(sys.executable)
    else:
        # Se estiver rodando como um script .py
        return os.path.dirname(os.path.abspath(__file__))

BASE_PATH = get_base_path()
dotenv_path = os.path.join(BASE_PATH, '.env')
load_dotenv(dotenv_path=dotenv_path)


# === VARIÁVEIS ===
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
DOMAIN = os.getenv("DOMAIN")
SITE_NAME = os.getenv("SITE_NAME")
DRIVE_NAME = os.getenv("DRIVE_NAME")
SOURCE_FOLDER = os.getenv("SOURCE_FOLDER")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
RECIPIENT_EMAILS = os.getenv("RECIPIENT_EMAILS", "").split(',')

# === Mapeamentos para subpastas ===
KEYWORD_MAP = {
    "cortes": "CORTE",
    "vendas": "VENDAS",
    "a_faturar": "AFATURAR"
}
PREFIX_MAP = {
    "rca_": "EQUIPE",
    "clientes_": "CLIENTES",
    "fornecedores_": "FORNECEDOR",
    "produtos_": "Produto linha",
    "clientes_rca_": "CLI RCA"
}

# === LOG SETUP ===
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_messages = []
def log(msg):
    print(msg)
    log_messages.append(msg)

def send_report_email(token: str, logs: list[str], is_error: bool = False):
    """Envia o e-mail de relatório."""
    if not (token and SENDER_EMAIL and RECIPIENT_EMAILS and RECIPIENT_EMAILS[0]):
        log("-> Variáveis de e-mail não configuradas. Pulando envio de relatório.")
        return

    date_str = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    if is_error:
        subject = f"ERRO FATAL - Upload SharePoint - {date_str}"
        body_intro = "Ocorreu um erro fatal durante o processo de upload de arquivos para o SharePoint.\n\n"
    else:
        subject = f"Relatório de Upload SharePoint - {date_str}"
        body_intro = "Processo de upload de arquivos para o SharePoint concluído.\n\n"

    body = body_intro + "Resumo da execução:\n" + "\n".join(logs)

    try:
        send_email(
            access_token=token,
            sender=SENDER_EMAIL,
            recipients=RECIPIENT_EMAILS,
            subject=subject,
            body=body,
            body_content=body,
            is_html=False
        )
        log(f"-> E-mail de {'erro' if is_error else 'relatório'} enviado com sucesso.")
    except Exception as email_exc:
        log(f"[ERRO AO ENVIAR E-MAIL] {str(email_exc)}")

# === EXECUÇÃO ===
token = None
try:
    token = get_access_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    site_id = get_site_id(token, DOMAIN, SITE_NAME)
    drive_id = get_drive_id(token, site_id, DRIVE_NAME)

    log(f"Processando arquivos modificados hoje... {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}")
    success, errors = upload_files(
        token, site_id, drive_id,
        SOURCE_FOLDER,
        keyword_map=KEYWORD_MAP,
        prefix_map=PREFIX_MAP,
        log=log
    )

    log(f"Total enviados com sucesso: {len(success)}")
    log(f"Total Arquivo ignorado: {len(errors)}")
    send_report_email(token, log_messages)

except Exception as e:
    log(f"[FATAL] {str(e)}")
    send_report_email(token, log_messages, is_error=True)

log_file_path = os.path.join(BASE_PATH, f"TaskSchedulerLog_{timestamp}.txt")
with open(log_file_path, "w", encoding="utf-8") as f:
    f.write("\n".join(log_messages))
