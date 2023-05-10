from pdf2docx import Converter
from pathlib import Path
from flask import Flask, request, render_template, send_file
from auth.drive_dowload import dowload_file
from auxiliary.leituraEscritaWord import read_file
from auxiliary.rename import rename
from docx.shared import Cm
from io import BytesIO
import pandas as pd
import docx 
import tempfile

app = Flask(__name__)
processed_file = None

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file']

        # Tentando ler o tipo de arquivo atráves do content_type
        try:
            if f.content_type == 'text/csv':
                dados = pd.read_csv(BytesIO(f.read()))
            elif f.content_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                dados = pd.read_excel(BytesIO(f.read()))
            elif f.content_type == 'application/octet-stream':
                dados = pd.read_excel(BytesIO(f.read()), engine="odf")
            else:
                raise Exception(f'Tipo de arquivo "{f.content_type}" não suportado.')
        except Exception as e:
            print(e)
            return "Erro ao processar o arquivo."

        # transformando os dados em uma lista de dicinário
        dados_dict = dados.to_dict(orient="records")
                    
        rename(dados_dict)

        # A partir daqui é a escrita do arquivo  
        DOCUMENTO = docx.Document()
    
        sec = DOCUMENTO.sections
        for section in sec:
            section.top_margin = Cm(1)
            section.bottom_margin = Cm(1)
            section.left_margin = Cm(1)
            section.right_margin = Cm(1)

        for i in dados_dict:
            
            # removendo esse valor do dicinário
            if "Carimbo de data/hora" in i.keys():
                del i["Carimbo de data/hora"]

            # título para cada resposta do setor 
            DOCUMENTO.add_heading(f"Respostas da {i['Identificação da Unidade/Gerência:']}", 0)

            for k, v in i.items():
               
                if (str(v) != "nan"):
                    DOCUMENTO.add_heading(f"{k}", 1)
                    DOCUMENTO.add_paragraph(f"{v}")

                    if ("Arquivo em PDF ou Documento (word ou odf)" in k):
                        try:
                            DOCUMENTO.add_heading(f"Conteúdo do arquivo anexado pela {i['Identificação da Unidade/Gerência:']}", 1)
                            file = dowload_file(str(v))

                            # pegando a extensão do arquivo
                            ext = Path(file).suffix.lower()

                            if (ext == ".pdf"):
                                cv = Converter(file)
                                cv.convert("doc.docx")      
                                cv.close()
                                parent = docx.Document("doc.docx")
                            elif (ext == '.docx'):
                                parent = docx.Document(file)

                            read_file(parent, DOCUMENTO)
                        except Exception as e:
                            return "Erro ao realizar a autenticação"
        
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as output:
            DOCUMENTO.save(output)
            return send_file(output.name)
        
    return render_template('index.html')


@app.route('/download')
def download_file():
    global processed_file
    
    if processed_file is not None:
        try:
            return send_file(BytesIO(processed_file), as_attachment=True)
        except Exception as e:
            print(e)
            return "Erro ao fazer download do arquivo."
    else:
        return "Nenhum arquivo processado encontrado."
    
if __name__ == '__main__':
    app.run(debug=True)