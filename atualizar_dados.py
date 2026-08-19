import os
import requests
import pandas as pd
import json
from io import BytesIO

# Puxa do cofre de Secrets do GitHub
client_id = os.getenv("MEU_CLIENT_ID")
client_secret = os.getenv("MEU_CLIENT_SECRET")
tenant_id = os.getenv("MEU_TENANT_ID")

site_id = "231de1cf-c260-40f1-8c16-9ea0400b82e0" 

def obter_token():
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    response = requests.post(url, data=payload)
    response.raise_for_status()
    return response.json().get("access_token")

def extrair_dados():
    token = obter_token()
    headers = {"Authorization": f"Bearer {token}"}

    # 1. Descobre qual é o drive principal (biblioteca de documentos) do site da Logística
    url_drives = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    res_drives = requests.get(url_drives, headers=headers)
    res_drives.raise_for_status()
    
    # Pega o ID da primeira biblioteca de documentos (Documentos Compartilhados)
    drives = res_drives.json().get("value", [])
    if not drives:
        raise Exception("Nenhum drive encontrado neste site do SharePoint.")
    
    drive_id = drives[0]["id"]

    # 2. Busca o arquivo pelo caminho exato dentro da biblioteca de documentos
    # Caminho relativo dentro do Documentos Compartilhados: DasboardEstoque/Pasta.xlsx
    caminho_interno = "DasboardEstoque/Pasta.xlsx"
    url_item = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{caminho_interno}"
    
    res_item = requests.get(url_item, headers=headers)
    res_item.raise_for_status()
    
    # Pega o link direto para download da planilha
    download_url = res_item.json().get("@microsoft.graph.downloadUrl")
    
    response = requests.get(download_url)
    response.raise_for_status()

    # Lê o Excel e converte para JSON com as colunas que a TV lê
    df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
    dados_json = df.to_dict(orient='records')

    # Salva o arquivo dados.json na raiz do repositório
    with open('dados.json', 'w', encoding='utf-8') as f:
        json.dump(dados_json, f, ensure_ascii=False, indent=4)

if __name__ == "__main__":
    extrair_dados()
