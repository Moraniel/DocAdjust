from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import docx

def read_file(parent, DOCUMENTO):
    # Verifica o tipo do elemento pai e seleciona o elemento apropriado
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("algo não está certo")

    # Itera sobre os elementos filhos do elemento pai
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):  # Se o elemento filho for um parágrafo, adiciona o texto ao documento
            if(len(Paragraph(child, parent).text) != 0): # se não for linha vazia
                DOCUMENTO.add_paragraph(Paragraph(child, parent).text)
        elif isinstance(child, CT_Tbl):  # Se o elemento filho for uma tabela, extrai os dados e cria uma nova tabela no documento
            table = Table(child, parent)
            tabelas = []
            col_names = []

            # Itera sobre as linhas da tabela original
            for row in table.rows:
                # Se for a primeira linha, extrai os nomes das colunas
                if not col_names:
                    col_names = [cell.text for cell in row.cells]
                else:
                    # Se não for a primeira linha, adiciona os valores da linha atual como valores das colunas correspondentes
                    tabela = {}
                    for j, cell in enumerate(row.cells):
                        col_name = col_names[j]
                        tabela[col_name] = cell.text
                    tabelas.append(tabela)

            # Cria uma nova tabela no documento e preenche com os dados da tabela original
            table_new = DOCUMENTO.add_table(rows=len(tabelas)+1, cols=len(col_names))
            table_new.style = 'Table Grid'
            hdr_cells = table_new.rows[0].cells
            for j, k in enumerate(col_names):
                hdr_cells[j].text = k

            for i, row in enumerate(table_new.rows):
                if i != 0:  # pula a primeira linha (já preenchida com os cabeçalhos)
                    if i <= len(tabelas):
                        for j, k in enumerate(col_names):
                            row.cells[j].text = tabelas[i-1].get(k, '')
                    else:
                        break