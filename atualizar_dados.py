import os
import requests
import pandas as pd
import json
from io import BytesIO

# Puxa do cofre de Secrets do GitHub
client_id = os.getenv("MEU_CLIENT_ID")
client_secret = os.getenv("MEU_CLIENT_SECRET")
tenant_id = os.getenv("MEU_TENANT_ID")

# O site_id que pescamos agora pouco
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

    # Busca o arquivo Excel direto pelo caminho da pasta no SharePoint da Logística
    # (Note que usamos o nome exato da sua biblioteca e pastas)
    caminho_arquivo = "/sites/LOGISTICACORPORATIVO/Documentos Compartilhados/DasboardEstoque/Pasta.xlsx"
    
    # Rota da Graph API que localiza pelo caminho (evita dor de cabeça com drive_id)
    url_drive = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{caminho_arquivo}"
    res_drive = requests.get(url_drive, headers=headers)
    res_drive.raise_for_status()
    
    # Pega o link de conteúdo direto do arquivo encontrado
    download_url = res_drive.json().get("@microsoft.graph.downloadUrl")
    
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
