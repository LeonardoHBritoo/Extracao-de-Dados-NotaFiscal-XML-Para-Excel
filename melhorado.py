import xmltodict
import os
import pandas as pd

def pegar_informacoes(nome_arquivo, valores, colunas):
    print(f"Peguei o arquivo {nome_arquivo}")
    with open(f'nfs/{nome_arquivo}', 'rb') as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)

        if 'NFe' in dic_arquivo:
            infos_nf = dic_arquivo["NFe"]["infNFe"]
        else:
            infos_nf = dic_arquivo['nfeProc']["NFe"]["infNFe"]

        numero_nota = infos_nf['@Id']
        empresa_emissora = infos_nf['emit']['xNome']
        nome_cliente = infos_nf['dest']['xNome']
        
        # Acessar diretamente os valores do endereço
        endereçoV = list(infos_nf['dest']['enderDest'].values())
        
        # Adicionar as chaves do dicionário de endereço às colunas
        endereçoC = list(infos_nf['dest']['enderDest'].keys())
        
        if 'vol' in infos_nf['transp']:
            peso = infos_nf['transp']['vol']['pesoB']
        else:
            peso = 'Não informado'
        
        # Verifique se as colunas de endereço já foram adicionadas
        if 'xLgr' not in colunas:
            colunas.extend(endereçoC)
        
        valores.append([numero_nota, empresa_emissora, nome_cliente, peso] + endereçoV )

lista_arquivos = os.listdir('nfs')
colunas = ['numero_nota', 'empresa_emissora', 'nome_cliente', 'peso']
valores = []
for arquivo in lista_arquivos:
    pegar_informacoes(arquivo, valores, colunas)
    
tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel('NotasFiscais.xlsx', index=False)
print('Processo concluído')
