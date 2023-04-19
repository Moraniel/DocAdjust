## Introdução:
GePy é um projeto para processar arquivos no formato ``.xlsx`` e ``.csv`` que contêm respostas de um formulário do Google Forms e gerar um documento no formato ``.docx`` com as respostas.

## Instalação e execução:

1. Certifique-se de ter Python 3.7 ou superior instalado em sua máquina. Navegue até a pasta raiz do projeto GePy no terminal e instale o ``Flask``, ``pandas``, ``openpyxl`` e ``python-docx`` usando o pip. Abra o terminal e digite o seguinte comando:

```s
pip install flask pandas openpyxl python-docx
```

2. Digite o seguinte comando para executar o aplicativo Flask:

```s
python app.py
```
O servidor será iniciado e o aplicativo estará disponível em ``http://127.0.0.1:5000``

## Uso:

1. Acesse o aplicativo em ``http://127.0.0.1:5000``
2. Clique no botão "Escolher arquivo" para selecionar o arquivo .xlsx ou .csv que contém as respostas do formulário.
3. Clique no botão "Enviar" para enviar o arquivo.
4. O arquivo será processado e um documento ``.docx`` será gerado com as respostas do formulário.
5. Clique no botão "Baixar" para baixar o documento ``.docx`` gerado.

## Requisitos:
* Python 3.7 ou superior
* Flask
* Pandas
* openpyxl
* python-docx

## Limitações:
* O projeto suporta apenas arquivos ``.xlsx`` e ``.csv`` gerados pelo Google Forms. Outros formatos de arquivo podem não ser processados corretamente.

## Conclusão:
O GePy é um projeto simples, mas eficiente para processar arquivos de respostas de formulários do Google e gerar um documento ``.docx`` com as respostas. Ele pode ser útil para gerar relatórios a partir de dados coletados por meio de formulários on-line.
