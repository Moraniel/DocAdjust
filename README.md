# GePy

<div align="center">
		<img src="https://user-images.githubusercontent.com/118611278/233367698-d5da373d-b861-45fb-85c3-0f2747f50829.png" alt="Imagem">
</div>

GePy é um projeto para processar arquivos no formato ``.xlsx``, ``.csv`` e ``.ods`` que contêm respostas de um formulário do Google Forms e gerar um documento no formato ``.docx`` com as respostas.

## Instalação e execução:

1. Certifique-se de ter Python 3.7 ou superior instalado em sua máquina. Navegue até a pasta raiz do projeto GePy no terminal e digite o seguinte comando:

```s
pip install -r requirements.txt
```

2. Para utilizar a API do Google Drive nesse projeto, é necessário seguir alguns passos importantes. Primeiramente, realize o cadastro no Google Cloud Platform. Em seguida, crie um projeto e ative a API do Google Drive em seu painel de controle. Por fim, baixe os arquivos ``credentials`` e ``client_secret`` e salve-os na pasta "auth"

3. Digite o seguinte comando para executar o aplicativo Flask:

```s
python app.py
```
O servidor será iniciado e o aplicativo estará disponível em ``http://127.0.0.1:5000``

## Uso:

1. Acesse o aplicativo em ``http://127.0.0.1:5000``
2. Clique no botão "Escolher arquivo" para selecionar o arquivo ``.xlsx``, ``.csv`` ou ``.ods`` que contém as respostas do formulário.
3. Clique no botão "Enviar" para enviar o arquivo.
4. O arquivo será processado e um documento ``.docx`` será gerado com as respostas do formulário.
5. Clique no botão "Baixar" para baixar o documento ``.docx`` gerado.

## Limitações:
* O projeto suporta apenas arquivos ``.xlsx``, ``.csv`` e ``.ods`` gerados pelo Google Planilha contendo as respostas do formulário. Outros formatos de arquivo podem não ser processados corretamente.

## Conclusão:
O GePy é um projeto simples, mas eficiente para processar arquivos de respostas de formulários do Google e gerar um documento ``.docx`` com as respostas. Ele pode ser útil para gerar relatórios a partir de dados coletados por meio de formulários on-line.
