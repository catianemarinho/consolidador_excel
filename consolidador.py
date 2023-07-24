import pandas as pd
import os
import datetime

# geração e formatação da data 
data = datetime.datetime.now()
data = data.strftime('%d-%m-%Y')

colunas = [
    'Segmento',
    'País',
    'Produto',
    'Qtde de Unidades Vendidas',
    'Preço Unitário',
    'Valor Total',	
    'Desconto',
    'Valor Total c/ Desconto',
    'Custo Total',
    'Lucro',
    'Data',
    'Mês',
    'Ano'
]

# cria um DataFrame vazio nomeando as colunas
consolidado = pd.DataFrame(columns=colunas)

# busca o nome dos arquivos na pasta a serem consolidados
arquivos = os.listdir('planilhas')

# realiza a consolidação de arquivos .xlsx
for arquivo in arquivos:

    if arquivo.endswith('.xlsx'):
        segmento = arquivo.split('-')[0]
        pais = arquivo.split('-')[1].replace('.xlsx', '')

        try:
            df = pd.read_excel(f'planilhas\\{arquivo}')
            df.insert(0, 'Segmento', segmento)
            df.insert(1, 'País', pais)

            consolidado = pd.concat([consolidado, df])
        except:
            with open('log_erros.txt', 'a') as file:
                file.write(f'Erro ao tentar consolidar o arquivo {arquivo}.')
    else:
        with open('log_erros.txt', 'a') as file:
                file.write(f'{data} - O arquivo {arquivo} não é um arquivo Excel válido!')
    
# exporta o arquivo consolidado para uma arquivo excel
consolidado.to_excel(f'Report consolidado em {data}.xlsx', sheet_name='dados', index=False)