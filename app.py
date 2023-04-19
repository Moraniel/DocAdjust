from flask import Flask, request, render_template, send_file
import pandas as pd
import docx
from docx.shared import Cm
from io import BytesIO
import tempfile

app = Flask(__name__)
processed_file = None

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file']
        print(f.content_type)


        # Tentando ler o tipo de arquivo atráves do content_type
        try:
            if f.content_type == 'text/csv':
                dados = pd.read_csv(BytesIO(f.read()))
            elif f.content_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                dados = pd.read_excel(BytesIO(f.read()))
            else:
                raise Exception(f'Tipo de arquivo "{f.content_type}" não suportado.')
        except Exception as e:
            print(e)
            return "Erro ao processar o arquivo."

        # transformando os dados em uma lista de dicinário
        dados_dict = dados.to_dict(orient="records")
        print(type(dados_dict))
        
        # essa parte do código se refere à uma alteração que o GEPLAN solicitou
        for i in dados_dict:
            for chave in list(i.keys()):
                if chave == 'Deliberações gerenciais;':
                    i['Atribuições do setor:'] = i.pop(chave)
                if chave == 'Diretor (a)/Coordenador (a) e ou Gerente:':
                    i['Nome Diretor (a)/ Coordenador(a) e ou Gerente:'] = i.pop(chave)

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
                    DOCUMENTO.add_heading(f"{k}", level=1)
                    DOCUMENTO.add_paragraph(f"{v}")
        
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