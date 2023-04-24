import io, docx, os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from urllib.parse import urlparse, parse_qs
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import tempfile


#Credenciais do Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive']

# Arquivo de credinciais IMPORTANTE(O email nas credencias tem que ter acesso aos arquivos do csv)
SERVICE_ACCOUNT_FILE = './credentials.env'

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Irá ler um arquivo word com base no caminho, procurar suas tabelas e organizar elas em uma lista
# Ex: lista= [{Coluna1:resposta1},{Coluna1:reposta1}]
def ler_word(file_link):
    doc = docx.Document(file_link)
    tabelas=[]
    for table in doc.tables:
        tabela = {}
        for i, row in enumerate(table.rows):
            # Se for a primeira linha, extrai os nomes das colunas
            if i == 0:
             col_names = [cell.text for cell in row.cells]
            else:
            # Se não for a primeira linha, adiciona os valores da linha atual como valores das colunas correspondentes
                for j, cell in enumerate(row.cells):
                    col_name = col_names[j]
                    tabela[col_name] = cell.text
        
    tabelas.append(tabela)
    return tabelas

# Recebe um dicionario dict1 (Dicionario original) e uma lista com dicionarios list2 (A lista de dicionarios com as tabelas do arquivo word)
def merge_dic(dict1, list2):

    result_dict = dict1.copy()
    for dict2 in list2:
        count = 0
        for key, value in dict2.items():
            #Caso exista alguma chave igual, vai renomear o valor dela
            if key in result_dict:
                new_key = f"{key}_{count}"
                while new_key in result_dict:
                    count += 1
                    new_key = f"{key}_{count}"
                result_dict[new_key] = value
                count = 1
            else:
                result_dict[key] = value
    return result_dict

# Receve um link do drive e retorna baixa o arquivo para pegar as informações      
def dowload_file(file_link):
    creds = None

    # Verificar se as credenciais existem
    if os.path.exists('token.env'):
        creds = Credentials.from_authorized_user_file('token.env', SCOPES)

    # Se as credenciais não existirem ou estiverem expiradas, solicitar ao usuário que faça login
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            'client_secret.env', SCOPES)
        creds = flow.run_local_server(port=0)

        # Salvar as credenciais para uso posterior
        with open('token.env', 'w') as token:
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

    # # ID do arquivo que você deseja baixar
    # url_components = urlparse(file_link)
    # query_params = parse_qs(url_components.query)

    # # Obtenha o valor do parâmetro 'id' da consulta
    # file_id = query_params['id'][0]


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

    # with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as output:
    #     output.write(file_content.getbuffer())
    
    # return output.name

    # Salvar o arquivo no disco
    with open(filename, 'wb') as f:
        f.write(file_content.getbuffer())
    return filename