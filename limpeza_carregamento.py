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

# Substituir na por "vazio"
arquivo['col1'] = arquivo['col1'].fillna("vazio")


##############################################################################################################################################
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


############################ OUTRA FORMA ###################################3

from nltk.tokenize import sent_tokenize, word_tokenize, wordpunct_tokenize
from nltk.stem import PorterStemmer
from nltk.corpus import stopwords

# Lista de stopwords do pacote NLTK
stops = set(stopwords.words('portuguese'))

stops_criadas = {'ficar', 'alex', 'caique', 'bianca', 'rayane', 'voces', 'flavia', 'fontenele',
             'tá', 'ta', 'pra', 'mim', 'paula', 'lopes', 'rayane/brunna', 'caíque', 'aline', 'gonçalves', 'bernado',
            'vitor', 'sei', '欄'}

# Tokenizar a série
arquivo['col1'] = arquivo['col1'].apply(lambda x: word_tokenize(x))

# Retirar Stops
arquivo['col1'] = arquivo['col1'].apply(lambda x :[i for i in x if not i.lower() in stops])

# Retirar stops_criadas
arquivo['col1'] = arquivo['col1'].apply(lambda x :[i for i in x if not i.lower() in stops_criadas])

# Juntar reverter tTokenização 
arquivo['col1'] = arquivo['col1'].str.join(' ')

######################################################################################################################

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

# Agrupar selecionando colunas e modo de agrupamento
arquivo = arquivo.groupby(by = ['col1']).agg({'col2': 'sum', 'col3': 'sum', 'col4': 'sum',
                                                'col5': 'sum', 'col6': 'sum', 'col7': 'sum',
                                                'col8': 'sum', 'col9': 'sum'})

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

# Salvando dataset e definindo separador decimal e formato de datas
arquivo_CSV = arquivo.to_csv("endereço e nome do arquivo.csv", date_format = '%m/%d/%Y', decimal= ',', sep=";", index=False, encoding='ANSI')

# Renomeando e salvando na pasta correta
download = "endereço da pasta"
os.chdir(download)
os.getcwd()
    
list_of_files = glob.glob('download\\*.csv')
arquivo = max(list_of_files , key=os.path.getctime)

new =  nome + '.csv'
os.replace(arquivo, new)
time.sleep(3)
    
# Excluir arquivo da pasta 

pasta = "endereço da pasta de destino"
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

##############################################################
import pandas as pd
import numpy as np
import pymsteams
import functools
import operator

def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

# Listando todas as duplicatas do dataset arquivo e enviando recado em canal do teams
# É utilizado a extensão Incoming Webhook para envios de mensagens para o canal
Duplicatas_arquivo = arquivos['col1'][arquivos['col1']duplicated()]
Duplicatas_arquivo = Duplicatas_arquivo.reset_index()

if len(Duplicatas_arquivo) > 0:

    Duplicatas_arquivo_lista = []
    Duplicatas_teams_lista = []

    for lista in Duplicatas_arquivo['col1']:
        lista_repetidos =  Duplicatas_arquivo_lista.append(lista)

    for lista in Duplicatas_arquivo_lista:
        lista_repetidos =  Duplicatas_teams_lista.append(str("* ") + str(lista) + str(' \n \n '))

    texto = functools.reduce(operator.add, (Duplicatas_teams_lista))
    a = str("TEXTO OPCIONAL: \n \n") + texto
    
    myTeamsMessage = pymsteams.connectorcard("LINK CANAL")
    myTeamsMessage.text("Informativo Setor de Inteligência:")
    myTeamsMessage.color('#A0D7C2')


    myMessageSection = pymsteams.cardsection()

    # Activity Elements
    myMessageSection.activityTitle("TÍTULO")
    myMessageSection.activitySubtitle(a)
    myMessageSection.activityText("TEXTO OPCIONAL")


    # Texto da Seção
    myMessageSection.text("TEXTO")


    # Adicione sua seção ao objeto do cartão do conector antes de enviar
    myTeamsMessage.addSection(myMessageSection)

    myTeamsMessage.send()
