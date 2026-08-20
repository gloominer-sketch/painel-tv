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

    # =========================================================
    # ARQUIVO 1: Pasta.xlsx (Dados de Estoque originais)
    # =========================================================
    caminho_1 = "DasboardEstoque/Pasta.xlsx"
    url_item_1 = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{caminho_1}"
    
    res_item_1 = requests.get(url_item_1, headers=headers)
    res_item_1.raise_for_status()
    
    download_url_1 = res_item_1.json().get("@microsoft.graph.downloadUrl")
    
    response_1 = requests.get(download_url_1)
    response_1.raise_for_status()

    # Lê o Excel e salva como dados.json
    df_1 = pd.read_excel(BytesIO(response_1.content), engine='openpyxl')
    dados_json_1 = df_1.to_dict(orient='records')

    with open('dados.json', 'w', encoding='utf-8') as f:
        json.dump(dados_json_1, f, ensure_ascii=False, indent=4)
        
    print("✅ Arquivo 1 (Pasta.xlsx) processado com sucesso e salvo como dados.json")

    # =========================================================
    # ARQUIVO 2: faturamento.xlsx (Novos dados de faturamento)
    # =========================================================
    caminho_2 = "DasboardEstoque/faturamento.xlsx"
    url_item_2 = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{caminho_2}"
    
    res_item_2 = requests.get(url_item_2, headers=headers)
    res_item_2.raise_for_status()
    
    download_url_2 = res_item_2.json().get("@microsoft.graph.downloadUrl")
    
    response_2 = requests.get(download_url_2)
    response_2.raise_for_status()

    # Lê o Excel com as colunas (faturamento geral, faturamento dia, mes, centro)
    df_2 = pd.read_excel(BytesIO(response_2.content), engine='openpyxl')
    dados_json_2 = df_2.to_dict(orient='records')

    # Salva como um SEGUNDO arquivo JSON
    with open('dados_faturamento.json', 'w', encoding='utf-8') as f:
        json.dump(dados_json_2, f, ensure_ascii=False, indent=4)
        
    print("✅ Arquivo 2 (faturamento.xlsx) processado com sucesso e salvo como dados_faturamento.json")

if __name__ == "__main__":
    extrair_dados()
