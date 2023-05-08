def rename(dict):
    # essa parte do código se refere à uma alteração que o GEPLAN solicitou
    for i in dict:
        for chave in list(i.keys()):
            if chave == 'Deliberações gerenciais;':
                i['Atribuições do setor:'] = i.pop(chave)
            if chave == 'Diretor (a)/Coordenador (a) e ou Gerente:':
                i['Nome Diretor (a)/ Coordenador(a) e ou Gerente:'] = i.pop(chave)
    