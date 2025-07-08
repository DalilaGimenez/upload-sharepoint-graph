# sharepoint.py

import requests
import os
from datetime import datetime

def get_site_id(access_token: str, domain: str, site_name: str) -> str:
    url = f"https://graph.microsoft.com/v1.0/sites/{domain}:/sites/{site_name}"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()["id"]

def get_drive_id(access_token: str, site_id: str, drive_name: str) -> str:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    drives = response.json()["value"]
    for drive in drives:
        if drive["name"] == drive_name:
            return drive["id"]
    raise ValueError(f"Drive '{drive_name}' não encontrado.")


def upload_files(access_token: str, site_id: str, drive_id: str, folder_path: str, keyword_map: dict, prefix_map: dict, log=None) -> tuple[list[str], list[str]]:
    success_logs = []
    error_logs = []
    today = datetime.now().date()

    for filename in os.listdir(folder_path):
        full_path = os.path.join(folder_path, filename)

        if os.path.isfile(full_path):
            file_modified = datetime.fromtimestamp(os.path.getmtime(full_path)).date()
            if file_modified != today:
                continue

            # Lógica de subpasta com base no prefixo ou palavras-chave
            subfolder = ""
            for prefix, mapped in prefix_map.items():
                if filename.startswith(prefix):
                    subfolder = mapped
                    break
            if not subfolder:
                for keyword, mapped in keyword_map.items():
                    if keyword in filename.lower():
                        subfolder = mapped
                        break

            # Se não encontrou subpasta, pula o arquivo
            if not subfolder:
                msg = f"[AVISO] {filename}: Nenhuma regra de mapeamento encontrada. Arquivo ignorado."
                if log: log(msg)
                error_logs.append(msg)
                continue

            destino_path = f"{subfolder}/{filename}"

            try:
                url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{destino_path}:/content"
                headers = {
                    "Authorization": f"Bearer {access_token}",
                    "Content-Type": "application/octet-stream"
                }
                with open(full_path, "rb") as file_stream:
                    response = requests.put(url, headers=headers, data=file_stream)
                    response.raise_for_status()
                    msg = f"[OK] {filename} → {subfolder}"
                    if log: log(msg)
                    success_logs.append(msg)
            except Exception as e:
                msg = f"[ERRO] {filename}: {str(e)}"
                if log: log(msg)
                error_logs.append(msg)

    return success_logs, error_logs
