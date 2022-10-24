import pandas as pd
import numpy as np
import os


# Carregar o arquivo CSV
# dtype = str : carrega todas as coluna do dataset como str
arquivo = pd.read_csv('ENDEREÇO DO ARQUIVO.csv', sep=";", dtype=str, encoding = 'UTF8')

# Carregar planilha excel com todas colunas como texto
arquivo = pd.read_excel('endereço do arquivo', 'planilha', dtype=str)  


# Selecionar colunas
arquivo = arquivo[["col1", "col2", "col3"]]



# Renomear colunas
arquivo = arquivo.rename(columns={'col1':'coluna1'})



# Transformar dados das colunas de texto em letras maiúscula
arquivo['col1'] = arquivo['col1'].str.upper()

# Substituir texto
# Caso queira substituir caractere especial, regex =  True
arquivo['col1'] = arquivo['col1'].str.replace(",", " ", regex=False)

# Dividir série em uma lista, separando por espaço
arquivo['col1'] = arquivo['col1'].str.split()

# Remover de texto palavra inteira, sem pegar partes de outras palavras, lembre-se de realizar split antes
# Remover LTDA
for lista in arquivo['col1']:
    try:
        try:
            lista.remove('palavra a ser removida')
        except AttributeError:
            pass
    except ValueError:
        pass

# Unir texto novamente para dividindo por espaço cada palavra da lista
# Ao realizar o procedimento de str.split e logo após o str.join removerá todos os tipos de espaçamentos, 
# simplificando o texto e removendo o caracteres especiais de espaçamento
arquivo['col1'] = arquivo['col1'].str.join(' ')

# Remover caracteres abaixo do fim do texto
arquivo['col1'] = arquivo['col1'].str.rstrip("0123456789")

# Remover espaços em branco em excesso no texto
arquivo['col1'] = arquivo['col1'].str.strip()


# Remover tudo que estiver filtro ou limitado na coluna col1
arquivo = arquivo[~(arquivo['col1'].str.contains('filtro|limitado', na= False))]

# Remover do dataset tudo está vazio na coluna col1
arquivo = arquivo[~(arquivo['col1'].isna())]

# Listando todas as duplicatas do dataset arquivo
Duplicatas_arquivo = arquivo['col1'][arquivo['col1'].duplicated()]
Duplicatas_arquivo = arquivo.reset_index()


Duplicatas_arquivo_lista = []


for lista in Duplicatas_arquivo['col1']:
    lista_repetidos =  Duplicatas_arquivo_lista.append(lista)

Duplicatas_arquivo_lista = '|'.join(Duplicatas_arquivo_lista)


Duplicatas_arquivo = arquivo[(arquivo['col1'].str.contains(
    Duplicatas_arquivo_lista, na= False))]

# Organizando Dataset por ordem crescente
arquivo['col1'] = arquivo['col1'].sort_values(by = 'col1', ascending = True)

# Apaga linhas duplicadas
arquivo = arquivo.drop_duplicates()

# Agrupa pela coluna col1, contando a quantidade de linhas
arquivo = arquivo.groupby("col").count()

# Agrupa pela coluna col1, somando as outras colunas
arquivo = arquivo.groupby('col1').aggregate([np.sum])

# resetando o index
arquivo = arquivo.reset_index()

# Dividindo série pelo delimitador "(" a direta e criando novas colunas, n=1 : quantidade de colunas que quero como resulto
arquivo = arquivo['col1'].str.rsplit('(', n= 1, expand = True)


# Join simples, usando o index como união
arquivo3 = arquivo.join(arquivo2)

# Deletar variáveis
del Cliente

# Alterando tipo da série para datetime
arquivo['col1'] = pd.to_datetime(arquivo['col1'])

# Comparando dado de linha com a linha anterior
arquivo['comparacao'] = arquivo['col1'] == arquivo['col1'].shift(1)                          

# Extraindo dias de uma coluna de timedelta
# errors = 'ignore' : para ignorar linhas em branco
arquivo['dias'] = arquivo['col1'].astype('timedelta64[D]', errors = 'ignore').astype(int, errors = 'ignore')

# Convertendo a série em numérico
arquivo['col1'] = pd.to_numeric(arquivo['col1'])

# Utilizando clásula where para criar uma nova coluna
arquivo['col1'] = np.where((arquivo['col1'] == True) & (arquivo['col2'] > 330) & (EMISSOES['Renovado2'] < (arquivo['col3'] + 330)),1,0)

# Excluindo colunas do dataset e sobescrevendo dataset
arquivo.drop(['col1', 'col2', 'col3',
              'col4', 'col5'], axis = 1, inplace= True)


# Unir dois datasets pela coluna
arquivo3 = arquivo.append(arquivo2, ignore_index=True)


# Salvando dataset no formato csv
# index =  false : não salvar index
arquivo_CSV = arquivo.to_csv("endereço o qual deseja salvar.csv", sep=";", index=False, encoding='ANSI')
arquivo_CSV

# Renomeando e salvando na pasta correta
download = "C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões"
os.chdir(download)
os.getcwd()
    
list_of_files = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\*.csv')
arquivo = max(list_of_files , key=os.path.getctime)

new =  nome + '.csv'
os.replace(arquivo, new)
time.sleep(3)
    
# Excluir arquivo da pasta 

pasta = "C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Cancelados\\Nosso"
os.chdir(pasta)
os.getcwd()

if os.path.exists(new):
    os.remove(new)
time.sleep(1)

    
# Salvar novo arquivo na pasta
os.chdir(download)
os.getcwd()

shutil.move( new , pasta)
time.sleep(1)


# Carregando arquivos do mesmo tipo localizado em uma pasta em um único dataset
arquivos = glob.glob('PASTA DOS ARQUIVOS\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
CVS_arquivos = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    CVS_arquivos.append(temp_df)

CVS_arquivos = pd.concat(CVS_arquivos, axis=0)


# Apagando as duplicatas de uma coluna
arquivos.drop_duplicates(subset=['col1'], inplace=True)

# Convertendo coluna em um datetime, quando o formato está como dd-mm-yyyy
arquivos['col1'] = pd.to_datetime(arquivos['col1'],  errors='ignore', dayfirst = False )

# Fazendo join especificando as colunas e transformando-as em index
arquivos3 = arquivos1.set_index('ID').join(arquivos2.set_index('ID'))

