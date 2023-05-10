import io
import os
from pathlib import Path
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from urllib.parse import urlparse
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import tempfile

#Credenciais do Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive']

# Arquivo de credinciais IMPORTANTE(O email nas credencias tem que ter acesso aos arquivos do csv)
SERVICE_ACCOUNT_FILE = r'auth\credentials.json'

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Receve um link do drive e retorna baixa o arquivo para pegar as informações      
def dowload_file(file_link):
    creds = None

    # Verificar se as credenciais existem
    if os.path.exists(r"auth\token.json"):
        creds = Credentials.from_authorized_user_file(r"auth\token.json", SCOPES)

    # Se as credenciais não existirem ou estiverem expiradas, solicitar ao usuário que faça login
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            r"auth\client_secret.json", SCOPES)
        creds = flow.run_local_server(port=0)

        # Salvar as credenciais para uso posterior
        with open(r"auth\token.json", 'w') as token:
            token.write(creds.to_json())

    # Crie o objeto da API do Google Drive
    service = build('drive', 'v3', credentials=creds)

    parsed_url = urlparse(file_link)
    file_id = ""
    # Verifica se a URL é do tipo "open?id="
    if "open?id=" in file_link:
        # Extrai o ID do arquivo
        file_id = parsed_url.query.split("=")[1]
        
    # Verifica se a URL é do tipo "d/"
    elif "/d/" in file_link:
        # Extrai o ID do arquivo
        file_id = parsed_url.path.split("/")[-2]

    # Recupere informações do arquivo
    file = service.files().get(fileId=file_id).execute()
   
    # Nome do arquivo
    filename = file['name']

    # Criar um objeto de arquivo para armazenar os dados baixados
    file_content = io.BytesIO()

    # Baixar o arquivo
    request = service.files().get_media(fileId=file_id)
    downloader = MediaIoBaseDownload(file_content, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()

    ext = Path(filename).suffix.lower()

    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as output:
        output.write(file_content.getbuffer())
    return output.name

    # # Salvar o arquivo no disco
    # with open(filename, 'wb') as f:
    #     f.write(file_content.getbuffer())
    # return filename