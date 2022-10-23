#!/usr/bin/env python
# coding: utf-8

# In[ ]:




#!/usr/bin/env python
# coding: utf-8

# ! pip install webdriver-manager

import selenium
import time
import datetime
import os
import shutil

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from datetime import datetime, date, timedelta
import pyautogui #controlar o mouse
import pyperclip #controlar o teclado
import urllib #módul para trabalhar um URLs
import pandas as pd
import glob
from plyer import notification #notificações
from win32com.client import Dispatch




options = webdriver.ChromeOptions()
preferences = {"download.default_directory": "C:\OneDrive\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\Apuração de Resultados\Dashs datastudio\Base para Gsheets\Arquivos Bitrix", "safebrowsing.enabled": "false"}
options.add_experimental_option("prefs", preferences)

driver = webdriver.Chrome(executable_path=r'C:\chromedriver.exe', options=options)

# Maximizando a tela
driver.maximize_window()



#ACESSANDO O SITE DO BITRIX → NEGÓCIOS
driver.get("https://ic3.bitrix24.com.br/crm/deal/category/0/")
#ENTRANDO COM USERNAME E OK
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[1]/div/div/div[3]/div/form/div/div[1]/div/input').send_keys('apuracao5@nossocertificado.com.br')
time.sleep(2)
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[1]/div/div/div[3]/div/form/div/div[5]/button[1]').click()
time.sleep(5)

#ENTRANDO COM SENHA E OK
driver.find_element_by_xpath('//*[@id="password"]').send_keys('98124515')
time.sleep(2)
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[1]/div/div/div[3]/div/form/div/div[3]/button[1]').click()
time.sleep(30)







# Todos os negócios
try:
    driver.get("https://ic3.bitrix24.com.br/crm/deal/list/")
    time.sleep(10)

except WebDriverException:
    driver.get("https://ic3.bitrix24.com.br/crm/deal/list/")
    time.sleep(10)





#REMOVER FILTRO "Negócios em andamento"
button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="CRM_DEAL_LIST_V12_search_container"]/div[1]/div[2]')))
button.click()
time.sleep(2)
sel2 = driver.find_element_by_xpath('//*[@id="CRM_DEAL_LIST_V12_search"]')
sel2.click()






# Aplicar filtro
buc = driver.find_element_by_xpath('//*[@id="popup-window-content-CRM_DEAL_LIST_V12_search_container"]/div/div/div[1]/div[2]/div[7]')
buc.click()
time.sleep(10)







#ACESSAR EXPORTAÇÃO DE NEGÓCIOS PARA CSV
driver.find_element_by_xpath('//*[@id="uiToolbarContainer"]/div[4]/button').click()
time.sleep(1)
driver.find_element_by_xpath('//*[@id="popup-window-content-toolbar_deal_list_settings_menu"]/div/div/span[3]').click()







#SELECIONAR 'Exportar todos os campos do negócio' E 'Exportar SKU detalhadas'
driver.find_element_by_id('EXPORT_DEAL_CSV_opt_EXPORT_ALL_FIELDS_inp').click()
driver.find_element_by_id('EXPORT_DEAL_CSV_opt_EXPORT_PRODUCT_FIELDS_inp').click()




#CLICAR EM EXECUTAR PARA CARREGAR DADOS
driver.find_element_by_xpath('//*[@id="EXPORT_DEAL_CSV"]/div[3]/button[1]').click()
#CLICAR EM "DOWNLOAD EXPORT FILE"
button = WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[*]/div[4]/a"))) 
button.click()
#WAIT UNTIL FILE IS DOWNLOADED
time.sleep(15)

#ALTERAR NOME DO ARQUIVO DEAL (ÚLTIMO ARQUIVO BAIXADO)
Initial_path = "C:\OneDrive\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\Apuração de Resultados\Dashs datastudio\Base para Gsheets\Arquivos Bitrix"
new_name = '{}'.format(r"DEAL_NEGÓCIOS_NOVO_BITRIX.csv")
filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
shutil.move(filename,os.path.join(Initial_path, new_name))







#ACESSANDO O SITE DO BITRIX → CONTATOS
driver.get("https://ic3.bitrix24.com.br/crm/contact/list/")
time.sleep(5)

#ACESSAR EXPORTAÇÃO DE CONTATOS PARA CSV
botao = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="uiToolbarContainer"]/div[4]/button'))) 

botao.click()
time.sleep(2)
driver.find_element_by_xpath('//*[@id="popup-window-content-toolbar_contact_list_settings_menu"]/div/div/span[3]').click()
#SELECIONAR 'Exportar informações' E 'Exportar todos os campos da empresa'
driver.find_element_by_id('EXPORT_CONTACT_CSV_opt_REQUISITE_MULTILINE_inp').click()
driver.find_element_by_id('EXPORT_CONTACT_CSV_opt_EXPORT_ALL_FIELDS_inp').click()
#CLICAR EM EXECUTAR PARA CARREGAR DADOS
driver.find_element_by_xpath('//*[@id="EXPORT_CONTACT_CSV"]/div[3]/button[1]').click()
#CLICAR EM "DOWNLOAD EXPORT FILE"
button = WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="EXPORT_CONTACT_CSV"]/div[4]/a'))) 
button.click()
#WAIT UNTIL FILE IS DOWNLOADED
time.sleep(15)
#ALTERAR NOME DO ARQUIVO DEAL (ÚLTIMO ARQUIVO BAIXADO)
Initial_path = "C:\OneDrive\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\Apuração de Resultados\Dashs datastudio\Base para Gsheets\Arquivos Bitrix"
new_name = '{}'.format(r"CONTACT_CONTATOS_NOVO_BITRIX.csv")
filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
shutil.move(filename,os.path.join(Initial_path, new_name))


#ACESSANDO O SITE DO BITRIX → EMPRESAS
driver.get("https://ic3.bitrix24.com.br/crm/company/list/")
time.sleep(5)
#ACESSAR EXPORTAÇÃO DE CONTATOS PARA CSV
botao = WebDriverWait(driver, 180).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="uiToolbarContainer"]/div[4]/button'))) 
botao.click()
time.sleep(1)
driver.find_element_by_xpath('//*[@id="popup-window-content-toolbar_company_list_settings_menu"]/div/div/span[2]').click()
#SELECIONAR 'Exportar informações' E 'Exportar todos os campos da empresa'
driver.find_element_by_id('EXPORT_COMPANY_CSV_opt_REQUISITE_MULTILINE_inp').click()
driver.find_element_by_id('EXPORT_COMPANY_CSV_opt_EXPORT_ALL_FIELDS_inp').click()
#CLICAR EM EXECUTAR PARA CARREGAR DADOS
driver.find_element_by_xpath('//*[@id="EXPORT_COMPANY_CSV"]/div[3]/button[1]').click()
#CLICAR EM "DOWNLOAD EXPORT FILE"
button = WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[*]/div[4]/a"))) 
button.click()
#WAIT UNTIL FILE IS DOWNLOADED
time.sleep(15)
#ALTERAR NOME DO ARQUIVO DEAL (ÚLTIMO ARQUIVO BAIXADO)
Initial_path = "C:\OneDrive\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\Apuração de Resultados\Dashs datastudio\Base para Gsheets\Arquivos Bitrix"
new_name = '{}'.format(r"COMPANY_EMPRESAS_NOVO_BITRIX.csv")
filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
shutil.move(filename,os.path.join(Initial_path, new_name))






# Acessando a GFIS Nosso certificado
link = "https://nossocertificado.gfsis.com.br/gestaofacil/login/Index"
driver.get(link)

usuario = "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[1]/input"
time.sleep(2)
driver.find_element_by_xpath(usuario).send_keys("VICTOR")
senha = "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[2]/input"
time.sleep(2)
driver.find_element_by_xpath(senha).send_keys("123456")
time.sleep(2)
button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[3]/div/input"))) 
button.click()
time.sleep(10)




# Baixando o arquivo
link1 = "https://nossocertificado.gfsis.com.br/gestaofacil/login/faturamento/crud/PontoAtendimento?ACAO=listagem"
driver.get(link1)
time.sleep(10)
driver.find_element_by_xpath('//*[@id="btn_exportar"]').click()
time.sleep(2)




# Renomeando arquivo
try:
    download = "C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\Base para Gsheets\\Arquivos Bitrix"
    os.chdir(download)
    os.getcwd()
    
    list_of_files = glob.glob('C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\Base para Gsheets\\Arquivos Bitrix\\*.csv')
    arquivo = max(list_of_files , key=os.path.getctime)

    new =  'CadastroUnidadesNossoCertificado.csv'
    os.replace(arquivo, new)
    time.sleep(3)
    
    # Excluir arquivo da pasta 

    pasta = "C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\GFSIS"
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
except ValueError:
    notification.notify(
    title='BOT BUC',
    message='Arquivos não baixados corretamente, execute o Bot novamente ou baixe manualmente',
)




# Acessando a GFIS Certifica Brasil
link2 = "https://certificabrasil.gfsis.com.br/gestaofacil/login/Index"
driver.get(link2)

usuario = "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[1]/input"
time.sleep(2)
driver.find_element_by_xpath(usuario).send_keys("VICTOR")
senha = "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[2]/input"
time.sleep(2)
driver.find_element_by_xpath(senha).send_keys("123456")
time.sleep(2)
button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[3]/div/input"))) 
button.click()
time.sleep(10)




# Baixando o arquivo
link1 = "https://certificabrasil.gfsis.com.br/gestaofacil/login/faturamento/crud/PontoAtendimento?ACAO=listagem"
driver.get(link1)
time.sleep(10)
driver.find_element_by_xpath('//*[@id="btn_exportar"]').click()
time.sleep(2)




# Renomeando arquivo
try:
    download = "C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\Arquivos Bitrix"
    os.chdir(download)
    os.getcwd()

    list_of_files = glob.glob('C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\Arquivos Bitrix\\*.csv')
    arquivo = max(list_of_files , key=os.path.getctime)

    new =  'CadastroUnidadesCertificaBrasil.csv'
    os.replace(arquivo, new)
    time.sleep(3)
    
    # Excluir arquivo da pasta 

    pasta = "C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\GFSIS"
    os.chdir(pasta)
    os.getcwd()

    if os.path.exists(new):
        os.remove(new)
    time.sleep(3)
    
    # Salvar novo arquivo na pasta
    os.chdir(download)
    os.getcwd()

    shutil.move( new , pasta)
    time.sleep(3)
except ValueError:
    notification.notify(
    title='BOT BUC',
    message='Arquivos não baixados corretamente, execute o Bot novamente ou baixe manualmente',
)


driver.quit()


# In[ ]:



import pandas as pd

import numpy as np
import os


# # Cadastro PA



# Carregar o arquivo CSV
COMPANY_EMPRESAS_NOVO_BITRIX = pd.read_csv('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\Arquivos Bitrix\\COMPANY_EMPRESAS_NOVO_BITRIX.csv', sep=";", dtype=str)




CADASTRO_PA = COMPANY_EMPRESAS_NOVO_BITRIX[["ID", "Nome da Empresa", "Tipo da empresa", "Telefone de trabalho", "Celular", "Email de trabalho", "Criado", "DOCUMENTO PA", "Endereço", "Complemento", "CEP", "UF", "Bairro", "Cidade", "Agente de Expansão", "CNAE PA", "Tipo de Pessoa", "Pessoa Responsável (CS)"]]




CADASTRO_PA = CADASTRO_PA.rename(columns={'Nome da Empresa':'Nome_da_Empresa'})




CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.upper()
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace(",", " ", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("-", " ", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("(", " ", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace(".", " ", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace(")", " ", regex=False)




CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.split()




### Remover ####
# Remover LTDA
for lista in CADASTRO_PA['Nome_da_Empresa']:
    try:
        try:
            lista.remove('LTDA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover ME
for lista in CADASTRO_PA['Nome_da_Empresa']:
    try:
        try:
            lista.remove('ME')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/S
for lista in CADASTRO_PA['Nome_da_Empresa']:
    try:
        try:
            lista.remove('S/S')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover CIA
for lista in CADASTRO_PA['Nome_da_Empresa']:
    try:
        try:
            lista.remove('CIA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover SS
for lista in CADASTRO_PA['Nome_da_Empresa']:
    try:
        try:
            lista.remove('SS')
        except AttributeError:
            pass
    except ValueError:
        pass
    

    
# Remover LT
for lista in CADASTRO_PA['Nome_da_Empresa']:
    try:
        try:
            lista.remove('LT')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/C
for lista in CADASTRO_PA['Nome_da_Empresa']:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover SC
for lista in CADASTRO_PA['Nome_da_Empresa']:
    try:
        try:
            lista.remove('SC')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover S/C
for lista in CADASTRO_PA['Nome_da_Empresa']:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass



CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.join(' ')




CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("Á", "A", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("É", "E", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("Ã", "A", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("Õ", "O", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("Í", "I", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("Ç", "C", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("Ú", "U", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("Â", "A", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("Ô", "O", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("Ó", "O", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("&", "E", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace(".", "", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.replace("  ", "", regex=False)
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.rstrip("0123456789")
CADASTRO_PA['Nome_da_Empresa'] = CADASTRO_PA['Nome_da_Empresa'].str.strip()




CADASTRO_PA = CADASTRO_PA[~(CADASTRO_PA['Nome_da_Empresa'].str.contains('INATIV|Inati', na= False))]




CADASTRO_PA = CADASTRO_PA[~(CADASTRO_PA['Nome_da_Empresa'].isna())]


# # Duplicatas CADASTRO_PA



# Localizando duplicatas no Nome da Empresa em Cadastro_PA
Duplicatas_CADASTRO_PA = CADASTRO_PA['Nome_da_Empresa'][CADASTRO_PA['Nome_da_Empresa'].duplicated()]
Duplicatas_CADASTRO_PA = Duplicatas_CADASTRO_PA.reset_index()

Duplicatas_CADASTRO_PA['Nome_da_Empresa']
Duplicatas_CADASTRO_PA_lista = []


for lista in Duplicatas_CADASTRO_PA['Nome_da_Empresa']:
    lista_repetidos =  Duplicatas_CADASTRO_PA_lista.append(lista)

Duplicatas_CADASTRO_PA_lista = '|'.join(Duplicatas_CADASTRO_PA_lista)


Duplicatas_CADASTRO_PA = CADASTRO_PA[(CADASTRO_PA['Nome_da_Empresa'].str.contains(
    Duplicatas_CADASTRO_PA_lista, na= False))]

Duplicatas_CADASTRO_PA = Duplicatas_CADASTRO_PA.sort_values(by = 'Nome_da_Empresa', ascending = True)


# In[ ]:



# CONTATOS DE EMPRESAS


CONTATOS = pd.read_csv('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\Arquivos Bitrix\\CONTACT_CONTATOS_NOVO_BITRIX.csv', sep=";", encoding = 'UTF8')




# Limpeza em NOME
CONTATOS['Nome'] = CONTATOS['Nome'].str.upper()




CONTATOS['Nome'] = CONTATOS['Nome'].str.replace("Á", "A", regex=False)
CONTATOS['Nome'] = CONTATOS['Nome'].str.replace("É", "E", regex=False)
CONTATOS['Nome'] = CONTATOS['Nome'].str.replace("Ã", "A", regex=False)
CONTATOS['Nome'] = CONTATOS['Nome'].str.replace("Í", "I", regex=False)
CONTATOS['Nome'] = CONTATOS['Nome'].str.replace("Õ", "O", regex=False)
CONTATOS['Nome'] = CONTATOS['Nome'].str.replace("Ç", "C", regex=False)
CONTATOS['Nome'] = CONTATOS['Nome'].str.replace("Â", "A", regex=False)
CONTATOS['Nome'] = CONTATOS['Nome'].str.replace("Ú", "U", regex=False)
CONTATOS['Nome'] = CONTATOS['Nome'].str.replace("Ô", "O", regex=False)
CONTATOS['Nome'] = CONTATOS['Nome'].str.replace("&", "E", regex=False)
CONTATOS['Nome'] = CONTATOS['Nome'].str.replace(".", "", regex=False)
CONTATOS['Nome'] = CONTATOS['Nome'].str.replace("-", "", regex=False)
CONTATOS['Nome'] = CONTATOS['Nome'].str.rstrip("0123456789")
CONTATOS['Nome'] = CONTATOS['Nome'].str.strip()

CONTATOS = CONTATOS[~(CONTATOS['Nome'].str.contains('INATIV|Inati', na= False))]


# Limpeza em EMPRESA
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.upper()
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace(",", " ", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("-", " ", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("(", " ", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace(".", " ", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace(")", " ", regex=False)




CONTATOS["Empresa"] = CONTATOS["Empresa"].str.split()




### Remover ####
# Remover LTDA
for lista in CONTATOS["Empresa"]:
    try:
        try:
            lista.remove('LTDA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover ME
for lista in CONTATOS["Empresa"]:
    try:
        try:
            lista.remove('ME')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/S
for lista in CONTATOS["Empresa"]:
    try:
        try:
            lista.remove('S/S')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover CIA
for lista in CONTATOS["Empresa"]:
    try:
        try:
            lista.remove('CIA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover SS
for lista in CONTATOS["Empresa"]:
    try:
        try:
            lista.remove('SS')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover LT
for lista in CONTATOS["Empresa"]:
    try:
        try:
            lista.remove('LT')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/C
for lista in CONTATOS["Empresa"]:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover SC
for lista in CONTATOS["Empresa"]:
    try:
        try:
            lista.remove('SC')
        except AttributeError:
            pass
    except ValueError:
        pass




CONTATOS["Empresa"] = CONTATOS["Empresa"].str.join(' ')





CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("Á", "A", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("É", "E", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("Ã", "A", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("Õ", "O", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("Í", "I", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("Ú", "U", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("Ç", "C", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("Â", "A", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("Ô", "O", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("Ó", "O", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("&", "E", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace(".", "", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.replace("  ", "", regex=False)
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.rstrip("0123456789")
CONTATOS["Empresa"] = CONTATOS["Empresa"].str.strip()





CONTATOS = CONTATOS.drop_duplicates()



CONTATOS = CONTATOS[['Nome', 'Empresa', 'Cargo','CPF']]





# # CONTATOS_INFO_CADASTRO



CONTATOS_INFO_CADASTRO = pd.read_csv('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\Arquivos Bitrix\\CONTACT_CONTATOS_NOVO_BITRIX.csv', sep=";", dtype=str)





# Limpeza em NOME
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.upper()



CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.replace("Á", "A", regex=False)
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.replace("É", "E", regex=False)
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.replace("Ã", "A", regex=False)
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.replace("Í", "I", regex=False)
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.replace("Õ", "O", regex=False)
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.replace("Ç", "C", regex=False)
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.replace("Â", "A", regex=False)
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.replace("Ú", "U", regex=False)
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.replace("Ô", "O", regex=False)
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.replace("&", "E", regex=False)
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.replace(".", "", regex=False)
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.replace("-", "", regex=False)
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.rstrip("0123456789")
CONTATOS_INFO_CADASTRO['Nome'] = CONTATOS_INFO_CADASTRO['Nome'].str.strip()




# Limpeza em EMPRESA
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.upper()
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace(",", " ", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("-", " ", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("(", " ", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace(".", " ", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace(")", " ", regex=False)





CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.split()





### Remover ####
# Remover LTDA
for lista in CONTATOS_INFO_CADASTRO['Empresa']:
    try:
        try:
            lista.remove('LTDA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover ME
for lista in CONTATOS_INFO_CADASTRO['Empresa']:
    try:
        try:
            lista.remove('ME')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/S
for lista in CONTATOS_INFO_CADASTRO['Empresa']:
    try:
        try:
            lista.remove('S/S')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover CIA
for lista in CONTATOS_INFO_CADASTRO['Empresa']:
    try:
        try:
            lista.remove('CIA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover SS
for lista in CONTATOS_INFO_CADASTRO['Empresa']:
    try:
        try:
            lista.remove('SS')
        except AttributeError:
            pass
    except ValueError:
        pass

        pass
# Remover LT
for lista in CONTATOS_INFO_CADASTRO['Empresa']:
    try:
        try:
            lista.remove('LT')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/C
for lista in CONTATOS_INFO_CADASTRO['Empresa']:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover SC
for lista in CONTATOS_INFO_CADASTRO['Empresa']:
    try:
        try:
            lista.remove('SC')
        except AttributeError:
            pass
    except ValueError:
        pass





CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.join(' ')




CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("Á", "A", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("É", "E", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("Ã", "A", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("Õ", "O", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("Ú", "U", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("Í", "I", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("Ç", "C", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("Â", "A", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("Ô", "O", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("&", "E", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace(".", "", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.replace("  ", "", regex=False)
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.rstrip("0123456789")
CONTATOS_INFO_CADASTRO['Empresa'] = CONTATOS_INFO_CADASTRO['Empresa'].str.strip()




CONTATOS_INFO_CADASTRO = CONTATOS_INFO_CADASTRO[CONTATOS_INFO_CADASTRO['Cargo'].str.contains('Prop', na= False)]




CONTATOS_INFO_CADASTRO = CONTATOS_INFO_CADASTRO[["Nome", "Empresa", "Telefone de trabalho", "Email de trabalho"]]




CONTATOS_INFO_CADASTRO = CONTATOS_INFO_CADASTRO[~(CONTATOS_INFO_CADASTRO['Nome'].str.contains('INATIV|Inati', na= False))]




CONTATOS_INFO_CADASTRO = CONTATOS_INFO_CADASTRO.rename(columns={'Telefone de trabalho':'Telefone do Proprietário'})
CONTATOS_INFO_CADASTRO = CONTATOS_INFO_CADASTRO.rename(columns={'Email de trabalho':'Email do Proprietário'})
CONTATOS_INFO_CADASTRO = CONTATOS_INFO_CADASTRO.rename(columns={'Nome':'Nome do Proprietário'})


# # Identificando Duplicatas CONTATOS_INFO_CADASTRO



# Localizando duplicatas no nome do Proprietário em Duplicatas_CONTATOS_INFO_CADASTRO
Duplicatas_CONTATOS_INFO_CADASTRO = CONTATOS_INFO_CADASTRO['Nome do Proprietário'][CONTATOS_INFO_CADASTRO['Nome do Proprietário'].duplicated()]
Duplicatas_CONTATOS_INFO_CADASTRO = Duplicatas_CONTATOS_INFO_CADASTRO.reset_index()

Duplicatas_CONTATOS_INFO_CADASTRO['Nome do Proprietário']
Duplicatas_CONTATOS_INFO_CADASTRO_lista = []


for lista in Duplicatas_CONTATOS_INFO_CADASTRO['Nome do Proprietário']:
    lista_repetidos =  Duplicatas_CONTATOS_INFO_CADASTRO_lista.append(lista)

Duplicatas_CONTATOS_INFO_CADASTRO_lista = '|'.join(Duplicatas_CONTATOS_INFO_CADASTRO_lista)


Duplicatas_CONTATOS_INFO_CADASTRO = CONTATOS_INFO_CADASTRO[(CONTATOS_INFO_CADASTRO['Nome do Proprietário'].str.contains(
    Duplicatas_CONTATOS_INFO_CADASTRO_lista, na= False))]

Duplicatas_CONTATOS_INFO_CADASTRO = Duplicatas_CONTATOS_INFO_CADASTRO.sort_values(by = 'Nome do Proprietário', ascending = True)


# # DEAL_NEGOCIOS_NOVO_BITRIX

# # DEAL_NEGOCIOS_NOVO_BITRIX



DEAL_NEGOCIOS_NOVO_BITRIX = pd.read_csv('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\Arquivos Bitrix\\DEAL_NEGÓCIOS_NOVO_BITRIX.csv', sep=";", dtype=str)







DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.upper()




DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("-", " ", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace(",", " ", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("(", " ", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace(".", " ", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace(")", " ", regex=False)




DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.split()




### Remover ####
# Remover LTDA
for lista in DEAL_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('LTDA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover ME
for lista in DEAL_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('ME')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/S
for lista in DEAL_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('S/S')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover CIA
for lista in DEAL_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('CIA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover SS
for lista in DEAL_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('SS')
        except AttributeError:
            pass
    except ValueError:
        pass
    
    
# Remover LT
for lista in DEAL_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('LT')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/C
for lista in DEAL_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover SC
for lista in DEAL_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('SC')
        except AttributeError:
            pass
    except ValueError:
        pass




DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.join(' ')




DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Á", "A", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("É", "E", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Ã", "A", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Õ", "O", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Í", "I", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Ç", "C", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Ú", "U", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Â", "A", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Ô", "O", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Ó", "O", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("&", "E", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace(".", "", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("  ", "", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.rstrip("0123456789")
DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'] = DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].str.strip()




DEAL_NEGOCIOS_NOVO_BITRIX = DEAL_NEGOCIOS_NOVO_BITRIX[~(DEAL_NEGOCIOS_NOVO_BITRIX['Empresa'].isna())]

DEAL_NEGOCIOS_NOVO_BITRIX = DEAL_NEGOCIOS_NOVO_BITRIX.rename(columns={'Produto.1':'Produto_1'})

DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.upper()


DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace("Á", "A", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace("É", "E", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace("Ã", "A", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace("Í", "I", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace("Õ", "O", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace("Ó", "O", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace("Ç", "C", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace("Ú", "U", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace("Â", "A", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace("Ô", "O", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace("&", "E", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace(".", "", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.replace("-", "", regex=False)
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.rstrip("0123456789")
DEAL_NEGOCIOS_NOVO_BITRIX["Contato"] = DEAL_NEGOCIOS_NOVO_BITRIX["Contato"].str.strip()

DEAL_NEGOCIOS_NOVO_BITRIX = DEAL_NEGOCIOS_NOVO_BITRIX.sort_values(by = 'Fase', ascending = False)
DEAL_NEGOCIOS_NOVO_BITRIX = DEAL_NEGOCIOS_NOVO_BITRIX.sort_values(by = 'Modificado', ascending = False)
DEAL_NEGOCIOS_NOVO_BITRIX = DEAL_NEGOCIOS_NOVO_BITRIX.drop_duplicates(subset=['Contato', 'Nome do negócio'])
DEAL_NEGOCIOS_NOVO_BITRIX = DEAL_NEGOCIOS_NOVO_BITRIX.dropna(subset=['Contato'])



# # CONSULTA_NEGOCIOS_NOVO_BITRIX



CONSULTA_NEGOCIOS_NOVO_BITRIX = pd.read_csv('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\Arquivos Bitrix\\DEAL_NEGÓCIOS_NOVO_BITRIX.csv', sep=";", dtype=str)




CONSULTA_NEGOCIOS_NOVO_BITRIX = CONSULTA_NEGOCIOS_NOVO_BITRIX[["Fase", "Renda", "Empresa", "Criado", "Produto.1"]]




CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.upper()




CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("-", " ", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace(",", " ", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("(", " ", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace(".", " ", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace(")", " ", regex=False)



CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.split()




### Remover ####
# Remover LTDA
for lista in CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('LTDA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover ME
for lista in CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('ME')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/S
for lista in CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('S/S')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover CIA
for lista in CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('CIA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover SS
for lista in CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('SS')
        except AttributeError:
            pass
    except ValueError:
        pass
    
    
# Remover LT
for lista in CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('LT')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/C
for lista in CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover SC
for lista in CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa']:
    try:
        try:
            lista.remove('SC')
        except AttributeError:
            pass
    except ValueError:
        pass




CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.join(' ')




CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Á", "A", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("É", "E", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Ã", "A", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Õ", "O", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Í", "I", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Ç", "C", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Ú", "U", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Â", "A", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Ô", "O", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("Ó", "O", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("&", "E", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace(".", "", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.replace("  ", "", regex=False)
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.rstrip("0123456789")
CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'] = CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].str.strip()




CONSULTA_NEGOCIOS_NOVO_BITRIX = CONSULTA_NEGOCIOS_NOVO_BITRIX[~(CONSULTA_NEGOCIOS_NOVO_BITRIX['Empresa'].isna())]

CONSULTA_NEGOCIOS_NOVO_BITRIX = CONSULTA_NEGOCIOS_NOVO_BITRIX.rename(columns={'Produto.1':'Produto_1'})


# Ordena por ordem alfabética o nome da empresa
CONSULTA_NEGOCIOS_NOVO_BITRIX = CONSULTA_NEGOCIOS_NOVO_BITRIX.sort_values(by = 'Criado', ascending = True)
CONSULTA_NEGOCIOS_NOVO_BITRIX = CONSULTA_NEGOCIOS_NOVO_BITRIX.sort_values(by = 'Empresa', ascending = True)

# # CONTATOS_CONT_AGR



CONTATOS_CONT_AGR = pd.read_csv('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\Arquivos Bitrix\\CONTACT_CONTATOS_NOVO_BITRIX.csv', sep=";", dtype=str)




CONTATOS_CONT_AGR = CONTATOS_CONT_AGR[CONTATOS_CONT_AGR['Cargo'].str.contains('Proprietário/AGR|Agente', na= False)]





CONTATOS_CONT_AGR = CONTATOS_CONT_AGR[["Nome", "Empresa", "Cargo"]]




# Limpeza em NOME
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.upper()




CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace("Á", "A", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace("É", "E", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace("Ã", "A", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace("Í", "I", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace("Õ", "O", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace("Ó", "O", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace("Ç", "C", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace("Ú", "U", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace("Â", "A", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace("Ô", "O", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace("&", "E", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace(".", "", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.replace("-", "", regex=False)
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.rstrip("0123456789")
CONTATOS_CONT_AGR["Nome"] = CONTATOS_CONT_AGR["Nome"].str.strip()





# Limpeza em EMPRESA
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.upper()
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace(",", " ", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("-", " ", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("(", " ", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace(".", " ", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace(")", " ", regex=False)





CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.split()





### Remover ####
# Remover LTDA
for lista in CONTATOS_CONT_AGR["Empresa"]:
    try:
        try:
            lista.remove('LTDA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover ME
for lista in CONTATOS_CONT_AGR["Empresa"]:
    try:
        try:
            lista.remove('ME')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/S
for lista in CONTATOS_CONT_AGR["Empresa"]:
    try:
        try:
            lista.remove('S/S')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover CIA
for lista in CONTATOS_CONT_AGR["Empresa"]:
    try:
        try:
            lista.remove('CIA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover SS
for lista in CONTATOS_CONT_AGR["Empresa"]:
    try:
        try:
            lista.remove('SS')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover LT
for lista in CONTATOS_CONT_AGR["Empresa"]:
    try:
        try:
            lista.remove('LT')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/C
for lista in CONTATOS_CONT_AGR["Empresa"]:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover SC
for lista in CONTATOS_CONT_AGR["Empresa"]:
    try:
        try:
            lista.remove('SC')
        except AttributeError:
            pass
    except ValueError:
        pass





CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.join(' ')





CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("Á", "A", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("É", "E", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("Ã", "A", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("Õ", "O", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("Í", "I", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("Ú", "U", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("Ç", "C", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("Â", "A", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("Ô", "O", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("Ó", "O", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("&", "E", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace(".", "", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.replace("  ", "", regex=False)
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.rstrip("0123456789")
CONTATOS_CONT_AGR["Empresa"] = CONTATOS_CONT_AGR["Empresa"].str.strip()





CONTATOS_CONT_AGR = CONTATOS_CONT_AGR.drop_duplicates()


CONTATOS_CONT_AGR = CONTATOS_CONT_AGR[~(CONTATOS_CONT_AGR['Empresa'].str.contains('INATIV|Inati', na= False))]



CONTATOS_CONT_AGR = CONTATOS_CONT_AGR.groupby("Empresa").count()





CONTATOS_CONT_AGR = CONTATOS_CONT_AGR[["Nome"]]





CONTATOS_CONT_AGR = CONTATOS_CONT_AGR.rename(columns={'Nome':'Quantidade de AGRs'})


# # NEGOCIOS_SUM_VALOR




NEGOCIOS_SUM_VALOR = pd.read_csv('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\Arquivos Bitrix\\DEAL_NEGÓCIOS_NOVO_BITRIX.csv', sep=";")





NEGOCIOS_SUM_VALOR = NEGOCIOS_SUM_VALOR[[ "Empresa","Renda"]]




# Limpeza em EMPRESA
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.upper()
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace(",", " ", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("-", " ", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("(", " ", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace(".", " ", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace(")", " ", regex=False)





NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.split()





### Remover ####
# Remover LTDA
for lista in NEGOCIOS_SUM_VALOR["Empresa"]:
    try:
        try:
            lista.remove('LTDA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover ME
for lista in NEGOCIOS_SUM_VALOR["Empresa"]:
    try:
        try:
            lista.remove('ME')
        except AttributeError:
            pass
    except ValueError:
        pass
    
# Remover S/S
for lista in NEGOCIOS_SUM_VALOR["Empresa"]:
    try:
        try:
            lista.remove('S/S')
        except AttributeError:
            pass
    except ValueError:
        pass
    
# Remover CIA
for lista in NEGOCIOS_SUM_VALOR["Empresa"]:
    try:
        try:
            lista.remove('CIA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover SS
for lista in NEGOCIOS_SUM_VALOR["Empresa"]:
    try:
        try:
            lista.remove('SS')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover LT
for lista in NEGOCIOS_SUM_VALOR["Empresa"]:
    try:
        try:
            lista.remove('LT')
        except AttributeError:
            pass
    except ValueError:
        pass
    
# Remover S/C
for lista in NEGOCIOS_SUM_VALOR["Empresa"]:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass
    
# Remover SC
for lista in NEGOCIOS_SUM_VALOR["Empresa"]:
    try:
        try:
            lista.remove('SC')
        except AttributeError:
            pass
    except ValueError:
        pass




NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.join(' ')





NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("Á", "A", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("É", "E", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("Ã", "A", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("Õ", "O", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("Í", "I", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("Ú", "U", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("Ç", "C", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("Â", "A", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("Ô", "O", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("Ó", "O", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("&", "E", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace(".", "", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("  ", "", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.replace("INATIVO", "", regex=False)
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.rstrip("0123456789")
NEGOCIOS_SUM_VALOR["Empresa"] = NEGOCIOS_SUM_VALOR["Empresa"].str.strip()





NEGOCIOS_SUM_VALOR = NEGOCIOS_SUM_VALOR.groupby('Empresa').aggregate([np.sum])





NEGOCIOS_SUM_VALOR = NEGOCIOS_SUM_VALOR.reset_index()




NEGOCIOS_SUM_VALOR = NEGOCIOS_SUM_VALOR.reset_index()





NEGOCIOS_SUM_VALOR = NEGOCIOS_SUM_VALOR[['Empresa', 'Renda']]




NEGOCIOS_SUM_VALOR = NEGOCIOS_SUM_VALOR.rename(columns={'Renda':'Valor dos Produtos'})



# # Emissões



EMISSOES = pd.read_csv('C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Financeiro CD\\APURAÇAO DE EMISSOES 4.0.csv', sep="," , dtype=str, encoding='ANSI') 






# EMISSOES = EMISSOES[['Identificador', 'Data', 'Data de aprovação', 'Situação', 'Vendedor','Cliente', 'E-mail', 'Telefone', 'Indicação', 'Valor total', 'Valor Total Nota', 'Valor Total Delivery', 'Observação','Itens do pedido de venda', 'Formas de pagamento do pedido de venda','Validação de Videoconferência', 'A quem cobrar?', 'TABELA','PREÇO VENDA', 'DATA BASE', 'TIPO', 'PERIODO DE COBRANÇA','Código AE ou PE', 'AE ou PE', '% AE ou PE', 'REPASSE AE ou PE','REPASSE AE ou PE LIQ', 'REPASSE EFETIVO AE ou PE', 'GE', '% GE','REPASSE GE', 'CUSTO\n(PE, AE e GE)', 'Código UE', 'UE', 'CUSTO (UE)','% UE', 'REPASSE UE', 'REPASSE DISTRIBUIDOR', 'SITUAÇÃO DE PAGAMENTO','DESPESA', 'DESPESA IMPOSTOS', 'DATA DINAMICA','DATA DINAMICA COMISSÃO', 'NOME FAIXA', 'AR', 'QUINZENA', 'AE ou PE2','CUSTO ULT FAIXA', 'CBO', 'LIQUIDO', 'Status Soluti', 'Conc Midias','% PE 2', 'CUSTO CENTRAL DE EMISSÃO', 'RESULTADO NOSSO CERTIFICADO','CPF AGR', 'CUSTO PARCEIRO INDICADDO', 'REPASSE PARCEIRO INDICADOR','CUSTAS PARCEIRO INDICADO', 'REPASSE PARCEIRO INDICADO', 'Franquia NTW','Validade', 'Tempo de Validade', 'Renovação?', 'Renovado?', 'Retorno','Já venceu?', 'Tipo de Produto', 'Critério Remuneração CBO','Tipo de Ponto', '% renovação do PA', 'PA faz Renovação?']]

EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.upper()
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("Á", "A", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("É", "E", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("Ã", "A", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("Í", "I", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("Õ", "O", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("Ç", "C", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("Ú", "U", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("Â", "A", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("Ô", "O", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("Ó", "O", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("&", "E", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace(".", "", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("-", "", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("*", "", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.rstrip("0123456789")
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.strip()


# # Limpeza feita pela inteligência:



# MUDANÇAS QUE AINDA NÃO FIZERAM EM EMISSÕES, APAGAR CÓDIGO DEPOIS
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("LUZIA LEITE RODRIGUES", "LUZIA RODRIGUES LEITE", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("AQUILA JESSICA FERREIRA DE OLIVEIRA", "AQUILA JESSICA", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("JOAO CARLOS BRITES ESPINOSA", "JOAO CARLOS BRITES", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("FRANCISCO DE CHAGAS MAGALHAES LIMA", "FRANCISCO DAS CHAGAS MAGALHAES LIMA", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("FERNANDA A ROSA ANDREOLI", "FERNANDA APARECIDA ROSA ANDREOLI", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("GLAUCIAMILESI", "GLAUCIA MILESI", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("HEVELLYNG ARAUJO", "HEVELLYNG ARAUJO SILVA", regex=False)
EMISSOES['Vendedor'] = EMISSOES['Vendedor'].str.replace("HEVELLYNG ARAUJO SILVA SILVA", "HEVELLYNG ARAUJO SILVA", regex=False)


EMISSOES['Cliente'] = EMISSOES['Cliente'].str.upper()
EMISSOES['Cliente'] = EMISSOES['Cliente'].str.replace("(", " (", regex=True)
EMISSOES['Cliente'] = EMISSOES['Cliente'].str.replace(")", " )", regex=True)
EMISSOES['Cliente'] = EMISSOES['Cliente'].str.split()
EMISSOES['Cliente'] = EMISSOES['Cliente'].str.join(' ')


# In[4]:


Cliente = EMISSOES['Cliente'].str.rsplit('(', n= 1, expand = True)


# In[6]:


Cliente.rename(columns = {0:'Nome do Cliente', 1: 'Documento do Cliente'}, inplace = True)


# In[7]:


Cliente['Documento do Cliente'] = Cliente['Documento do Cliente'].str.replace(")", "", regex=True)
Cliente['Documento do Cliente'] = Cliente['Documento do Cliente'].str.replace("/", "", regex=True)
Cliente['Documento do Cliente'] = Cliente['Documento do Cliente'].str.replace(".", "", regex=True)
Cliente['Documento do Cliente'] = Cliente['Documento do Cliente'].str.replace("-", "", regex=True)
Cliente['Documento do Cliente'] = Cliente['Documento do Cliente'].str.strip()


# In[8]:


EMISSOES = Cliente.join(EMISSOES)

del Cliente


# In[10]:


EMISSOES['Data de aprovação'] = pd.to_datetime(EMISSOES['Data de aprovação'])


# In[11]:


EMISSOES.sort_values(by = ['Documento do Cliente','Data de aprovação'], 
                      ascending = True, inplace= True)


# In[12]:


EMISSOES.columns


# In[13]:


EMISSOES.reset_index(inplace= True)


# In[14]:


EMISSOES['Renovado1'] = EMISSOES['Documento do Cliente'] == EMISSOES['Documento do Cliente'].shift(1)                          


# In[15]:


EMISSOES['Renovado2'] = EMISSOES['Data de aprovação'] - EMISSOES['Data de aprovação'].shift(1) 
EMISSOES['Validade_ant'] = EMISSOES['Tempo de Validade'].shift(1)


# In[16]:


EMISSOES['Renovado2'] = EMISSOES['Renovado2'].astype('timedelta64[D]', errors = 'ignore').astype(int, errors = 'ignore')


# In[17]:


EMISSOES['Renovado2'] = np.where(EMISSOES['Renovado2'].isna(), 0, EMISSOES['Renovado2'])
EMISSOES['Validade_ant'] = pd.to_numeric(EMISSOES['Validade_ant'])


# In[25]:


EMISSOES['Renovação1'] = np.where((EMISSOES['Renovado1'] == True) & (EMISSOES['Renovado2'] > 330) & (EMISSOES['Renovado2'] < (EMISSOES['Validade_ant'] + 330)),1,0)

EMISSOES['Renovado1'] = np.where(EMISSOES['Renovação1'].shift(-1) == 1 ,1,0)
# In[29]:


EMISSOES.drop(['index', 'Nome do Cliente', 'Documento do Cliente',
              'Renovado2', 'Validade_ant'], axis = 1, inplace= True)

EMISSOES.sort_values(by = ['Data de aprovação'], 
                      ascending = True, inplace= True)

AGR = pd.read_excel('C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Financeiro CD\\APURAÇAO DE EMISSOES 4.0 - Run.xlsm', 'AGR', dtype=str) 




# Limpeza em PA
AGR['PA'] = AGR['PA'].str.upper()
AGR['PA'] = AGR['PA'].str.replace(",", " ", regex=False)
AGR['PA'] = AGR['PA'].str.replace("-", " ", regex=False)
AGR['PA'] = AGR['PA'].str.replace("(", " ", regex=False)
AGR['PA'] = AGR['PA'].str.replace(".", " ", regex=False)
AGR['PA'] = AGR['PA'].str.replace(")", " ", regex=False)





AGR['PA'] = AGR['PA'].str.split()




### Remover ####
# Remover LTDA
for lista in AGR['PA']:
    try:
        try:
            lista.remove('LTDA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover ME
for lista in AGR['PA']:
    try:
        try:
            lista.remove('ME')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/S
for lista in AGR['PA']:
    try:
        try:
            lista.remove('S/S')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover CIA
for lista in AGR['PA']:
    try:
        try:
            lista.remove('CIA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover SS
for lista in AGR['PA']:
    try:
        try:
            lista.remove('SS')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover LT
for lista in AGR['PA']:
    try:
        try:
            lista.remove('LT')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/C
for lista in AGR['PA']:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover SC
for lista in AGR['PA']:
    try:
        try:
            lista.remove('SC')
        except AttributeError:
            pass
    except ValueError:
        pass
    
# Remover SC
for lista in AGR['PA']:
    try:
        try:
            lista.remove('PE')
        except AttributeError:
            pass
    except ValueError:
        pass





AGR['PA'] = AGR['PA'].str.join(' ')





AGR['PA'] = AGR['PA'].str.replace("Á", "A", regex=False)
AGR['PA'] = AGR['PA'].str.replace("É", "E", regex=False)
AGR['PA'] = AGR['PA'].str.replace("Ã", "A", regex=False)
AGR['PA'] = AGR['PA'].str.replace("Õ", "O", regex=False)
AGR['PA'] = AGR['PA'].str.replace("Í", "I", regex=False)
AGR['PA'] = AGR['PA'].str.replace("Ú", "U", regex=False)
AGR['PA'] = AGR['PA'].str.replace("Ç", "C", regex=False)
AGR['PA'] = AGR['PA'].str.replace("Â", "A", regex=False)
AGR['PA'] = AGR['PA'].str.replace("Ô", "O", regex=False)
AGR['PA'] = AGR['PA'].str.replace("Ó", "O", regex=False)
AGR['PA'] = AGR['PA'].str.replace("&", "E", regex=False)
AGR['PA'] = AGR['PA'].str.replace(".", "", regex=False)
AGR['PA'] = AGR['PA'].str.replace("  ", "", regex=False)
AGR['PA'] = AGR['PA'].str.replace("INATIVO", "", regex=False)
AGR['PA'] = AGR['PA'].str.rstrip("0123456789")
AGR['PA'] = AGR['PA'].str.strip()





# Limpeza em NOME
AGR['AGR'] = AGR['AGR'].str.upper()
AGR['AGR'] = AGR['AGR'].str.replace("Á", "A", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace("É", "E", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace("Ã", "A", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace("Í", "I", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace("Õ", "O", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace("Ç", "C", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace("Ú", "U", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace("Â", "A", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace("Ô", "O", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace("Ó", "O", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace("&", "E", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace(".", "", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace("-", "", regex=False)
AGR['AGR'] = AGR['AGR'].str.replace("*", "", regex=False)
AGR['AGR'] = AGR['AGR'].str.rstrip("0123456789")
AGR['AGR'] = AGR['AGR'].str.strip()


# # PackControleDeVendas



PackControleDeVendas = pd.read_excel('C:\OneDrive\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\Financeiro CD\Controle de Vendas C.xlsx', 'Pack', dtype=str)  



PackControleDeVendas = PackControleDeVendas.rename(columns={'CNPJPA':'CNPJ PA'})




PackControleDeVendas = PackControleDeVendas[["x", "UE", "AE / PE", "GE", "PA", "CNPJ PA", "AGR", "PACK", "PARCELA COMISSÃO", "VALOR TOTAL", "DATA EMAIL", "DATA FICHA", "VENCIMENTO", "FORMA", "COBRADO", "DESPESA", "RECEBIDO", "DATA REC", "SITUAÇÃO ENVIOS", "CÓD RASTREIO", "CADASTRO", "CUSTO PACK", "CUSTO NF", "DIFERENÇA CUSTO", "VALOR COMISSÃO AE", "VALOR COMISSÃO GE", "VALOR COMISSÃO EX INTER", "DATA PGMT COMISSÃO", "Data 1º Pgto", "Data 1º Venc", "OBSERVAÇÃO"]]




PackControleDeVendas = PackControleDeVendas[~(PackControleDeVendas['UE'].str.contains('CADASTRAR NA PLANILHA DE EMISSÕES', na= False))]





# Limpeza em PA
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.upper()
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace(",", " ", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("-", " ", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("(", " ", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace(".", " ", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace(")", " ", regex=False)





PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.split()





### Remover ####
# Remover LTDA
for lista in PackControleDeVendas['PA']:
    try:
        try:
            lista.remove('LTDA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover ME
for lista in PackControleDeVendas['PA']:
    try:
        try:
            lista.remove('ME')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/S
for lista in PackControleDeVendas['PA']:
    try:
        try:
            lista.remove('S/S')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover CIA
for lista in PackControleDeVendas['PA']:
    try:
        try:
            lista.remove('CIA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover SS
for lista in PackControleDeVendas['PA']:
    try:
        try:
            lista.remove('SS')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover LT
for lista in PackControleDeVendas['PA']:
    try:
        try:
            lista.remove('LT')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/C
for lista in PackControleDeVendas['PA']:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover SC
for lista in PackControleDeVendas['PA']:
    try:
        try:
            lista.remove('SC')
        except AttributeError:
            pass
    except ValueError:
        pass



# In[79]:


PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.join(' ')





PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("Á", "A", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("É", "E", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("Ã", "A", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("Õ", "O", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("Í", "I", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("Ú", "U", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("Ç", "C", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("Â", "A", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("Ô", "O", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("Ó", "O", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("&", "E", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace(".", "", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("  ", "", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("INATIVO", "", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.rstrip("0123456789")
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.strip()

PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("RKF CERTIFICADO DIGITAL E APOIO ADMINISTRATIVO", "RKF CERTIFICACAO DIGITAL E APOIO ADMINISTRATIVO", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("JOSEFO GENILDA PINTO DE OLIVEIRA", "JOSEFA GENILDA PINTO DE OLIVEIRA", regex=False)



PackControleDeVendas = PackControleDeVendas[['x', 'UE', 'AE / PE', 'GE', 'PA', 'CNPJ PA', 'AGR', 'PACK',
                      'PARCELA COMISSÃO', 'VALOR TOTAL', 'DATA EMAIL', 'DATA FICHA',
                      'VENCIMENTO', 'FORMA', 'COBRADO', 'DESPESA', 'RECEBIDO', 'DATA REC',
                      'SITUAÇÃO ENVIOS', 'CÓD RASTREIO', 'CADASTRO', 'CUSTO PACK', 'CUSTO NF',
                      'DIFERENÇA CUSTO', 'VALOR COMISSÃO AE', 'VALOR COMISSÃO GE',
                      'VALOR COMISSÃO EX INTER', 'DATA PGMT COMISSÃO', 'Data 1º Pgto','Data 1º Venc']]


# # CursosAvulsosControleDeVendas



CursosAvulsosControleDeVendas = pd.read_excel('C:\OneDrive\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\Financeiro CD\Controle de Vendas C.xlsx', 'Cursos Avulsos', dtype=str)  



# Limpeza em RECEBIDO DE
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.upper()
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace(",", " ", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("-", " ", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("(", " ", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace(".", " ", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace(")", " ", regex=False)





CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.split()





### Remover ####
# Remover LTDA
for lista in CursosAvulsosControleDeVendas['RECEBIDO DE']:
    try:
        try:
            lista.remove('LTDA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover ME
for lista in CursosAvulsosControleDeVendas['RECEBIDO DE']:
    try:
        try:
            lista.remove('ME')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/S
for lista in CursosAvulsosControleDeVendas['RECEBIDO DE']:
    try:
        try:
            lista.remove('S/S')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover CIA
for lista in CursosAvulsosControleDeVendas['RECEBIDO DE']:
    try:
        try:
            lista.remove('CIA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover SS
for lista in CursosAvulsosControleDeVendas['RECEBIDO DE']:
    try:
        try:
            lista.remove('SS')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover LT
for lista in CursosAvulsosControleDeVendas['RECEBIDO DE']:
    try:
        try:
            lista.remove('LT')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/C
for lista in CursosAvulsosControleDeVendas['RECEBIDO DE']:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover SC
for lista in CursosAvulsosControleDeVendas['RECEBIDO DE']:
    try:
        try:
            lista.remove('SC')
        except AttributeError:
            pass
    except ValueError:
        pass





CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.join(' ')





CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("Á", "A", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("É", "E", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("Ã", "A", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("Õ", "O", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("Í", "I", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("Ú", "U", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("Ç", "C", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("Â", "A", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("Ô", "O", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("Ó", "O", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("&", "E", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace(".", "", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("  ", "", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("INATIVO", "", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.rstrip("0123456789")
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.strip()


# CORREÇÃO DE NA PLANILHA DE CURSOS
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("RKF CERTIFICADO DIGITAL E APOIO ADMINISTRATIVO", "RKF CERTIFICACAO DIGITAL E APOIO ADMINISTRATIVO", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("WANDERSON HELDER DE LIMASILVA", "WANDERSON HELDER DE LIMA SILVA", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("RD FIGUEIREDO ASSESSORIA DOCUMENTAL", "R D FIGUEIREDO ASSESSORIA DOCUMENTAL", regex=False)




CursosAvulsosControleDeVendas = CursosAvulsosControleDeVendas[~(CursosAvulsosControleDeVendas['AGR'].isna())]

CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.split()
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.join(' ')
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.upper()


CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("MATTHEOS KOGUT", "MATHEOS FERNANDO SUTIL KOGUT", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("JONATHAN MAX COUTO", "JONATHAN MAX DO NASCIMENTO COUTO", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("MICHELLE STEPHANY", "MICHELLE STEPHANY DE LIMA MELLO SERRA", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("FERNANDA CRISTINA AMBROSINI SILVA", "FERNANDA CRISTINA AMBROSINI SILVA DOS SANTOS", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("CICERO COSMO ME", "CICERO COSMO", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("EMERSON DEL CONTI BATISTA LIMA", "EMERSON DEL COLI BATISTA LIMA", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("RAFAELA CRISTINA SANTOS", "RAFAELA CRISTINA SANTOS DE OLIVEIRA", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("MARIANNE COIMBRA / EDSON FRANCO ALVES", "MARIANNE COIMBRA DE OLIVEIRA", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("KAIO ALLEN / INDRIDY VITORIA", "KAIO ALLEN CHRISOSTOMO BERNARDINO", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("MATHIAS PEREIRA FRANÇA", "MATHIAS PEREIRA FRANCA", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("KELLY CRISTINA", "KELLY CRISTINA MARTINI VITTO", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("TATIELE SOARES GONÇALVE", "TATIELE SOARES GONCALVES", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("GUSTAVO MAGALHES LOCATELI", "GUSTAVO MAGALHAES LOCATELLI", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("ANA PAULA CANTALICE", "ANA PAULA CANTALICE DOS SANTOS", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("JHONNY FERNANDE", "JHONNY FERNANDES", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("MARIANNE COIMBRA", "MARIANNE COIMBRA DE OLIVEIRA", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("NILCIA LA SCALA / RONALDO MOREIRA", "NILCIA LA SCALA", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("LUANA DE CASTRO FEDOR", "LUANA DE CASTRO BARROS FEDOR", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("ADAUTON BECKER / JOÃO LUIZ GIOVANELLI", "ADAUTON BECKER", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("KLEBER EMERSON NAVARRO / MARCOS WANDSON PEREIRA", "KLEBER EMERSON NAVARRO", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("ROSA MELONI", "ROSANA MELONI TEIXEIRA", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("NILCIA LA SCALA / RONALDO MOREIRA", "NILCIA LA SCALA", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("LADSON JULIO CRUZ RAIOL", "LADSON JULIO DA CRUZ RAIOL", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("SAULO JOSÉ FARIAS", "SAULO JOSE FARIAS", regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.replace("MARIANNE COIMBRA DE OLIVEIRA DE OLIVEIRA", "MARIANNE COIMBRA DE OLIVEIRA", regex=False)



CursosAvulsosControleDeVendas = CursosAvulsosControleDeVendas[[ 'UE', 'AE / PE', 'RECEBIDO DE', 'CNPJ', 'AGR','PACK/CURSO', 'DATA EMAIL', 'DATA FICHA',
                               'FORMA PGMT', 'VENCIMENTO','PAGAMENTO', 'VALOR', 'RECEBIDO', 'AR', 'OBSERVAÇÕES',
                               'Valor em Aberto']]


# # 1. CadastroUnidadesCertificaBrasil




CadastroUnidadesCertificaBrasil = pd.read_csv('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\GFSIS\\CadastroUnidadesCertificaBrasil.csv', sep=";", encoding='ANSI')
CadastroUnidadesCertificaBrasil['Identificador'] = "CERTIFICA"


# # 2. CadastroUnidadesNossoCertificado




CadastroUnidadesNossoCertificado = pd.read_csv('C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\GFSIS\\CadastroUnidadesNossoCertificado.csv', sep=";", encoding='ANSI')
CadastroUnidadesNossoCertificado['Identificador'] = "NOSSO"


# # GFSIS


GFSIS = CadastroUnidadesCertificaBrasil.append(CadastroUnidadesNossoCertificado, ignore_index=True)


# In[92]:


# Limpeza em Nome
GFSIS['Nome'] = GFSIS['Nome'].str.upper()
GFSIS['Nome'] = GFSIS['Nome'].str.replace(",", " ", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("-", " ", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("(", " ", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace(".", " ", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace(")", " ", regex=False)





GFSIS['Nome'] = GFSIS['Nome'].str.split()


# In[94]:


### Remover ####
# Remover LTDA
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('LTDA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover ME
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('ME')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/S
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('S/S')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover CIA
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('CIA')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover SS
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('SS')
        except AttributeError:
            pass
    except ValueError:
        pass

# Remover LT
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('LT')
        except AttributeError:
            pass
    except ValueError:
        pass
# Remover S/C
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass
    
# Remover S/C
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('S/C')
        except AttributeError:
            pass
    except ValueError:
        pass
    
# Remover SC
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('SC')
        except AttributeError:
            pass
    except ValueError:
        pass
    
# Remover INATIVADA
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('INATIVADA')
        except AttributeError:
            pass
    except ValueError:
        pass
    
# Remover INATIVA
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('INATIVA')
        except AttributeError:
            pass
    except ValueError:
        pass
    
# Remover INATIVO
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('INATIVO')
        except AttributeError:
            pass
    except ValueError:
        pass
    
# Remover DESATIVADO
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('DESATIVADO')
        except AttributeError:
            pass
    except ValueError:
        pass



GFSIS['Nome'] = GFSIS['Nome'].str.join(' ')




GFSIS['Nome'] = GFSIS['Nome'].str.replace("Á", "A", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("Â", "A", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("É", "E", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("Ã", "A", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("Õ", "O", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("Í", "I", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("Ú", "U", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("Ç", "C", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("Ô", "O", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("Ó", "O", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("&", "E", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace(".", "", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("  ", "", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("INATIVO", "", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.replace("INATI", "", regex=False)
GFSIS['Nome'] = GFSIS['Nome'].str.rstrip("0123456789")
GFSIS['Nome'] = GFSIS['Nome'].str.strip()




GFSIS = GFSIS.drop_duplicates()

# Localizando duplicatas no nome no GFSIS

Duplicatas_GFSIS = GFSIS[(GFSIS['Ativo'].str.contains('Sim'))]
Duplicatas_GFSIS = Duplicatas_GFSIS[Duplicatas_GFSIS['Nome'].duplicated()]
Duplicatas_GFSIS = Duplicatas_GFSIS[~(Duplicatas_GFSIS['Nome'].str.contains('PE '))]
Duplicatas_GFSIS = Duplicatas_GFSIS[Duplicatas_GFSIS['Nome'].duplicated()]





Duplicatas_GFSIS = Duplicatas_GFSIS[Duplicatas_GFSIS['Nome'].duplicated()]

Duplicatas_GFSIS['Nome']
Duplicatas_GFSIS_lista = []


for lista in Duplicatas_GFSIS['Nome']:
    lista_repetidos =  Duplicatas_GFSIS_lista.append(lista)

Duplicatas_GFSIS_lista = '|'.join(Duplicatas_GFSIS_lista)


Duplicatas_GFSIS = Duplicatas_GFSIS[(Duplicatas_GFSIS['Nome'].str.contains(
    Duplicatas_GFSIS_lista, na= False))]

Duplicatas_GFSIS = Duplicatas_GFSIS.sort_values(by = 'Nome', ascending = True)
Duplicatas_GFSIS





GFSIS['Nome'] = GFSIS['Nome'].str.split()

# Remover PE
for lista in GFSIS['Nome']:
    try:
        try:
            lista.remove('PE')
        except AttributeError:
            pass
    except ValueError:
        pass
    
GFSIS['Nome'] = GFSIS['Nome'].str.join(' ')





GFSIS['DOCUMENTO PA'] = GFSIS['CNPJ'].fillna(GFSIS['CPF'])
GFSIS['DOCUMENTO PA'] = GFSIS['CNPJ Unidade'].fillna(GFSIS['DOCUMENTO PA'])





GFSIS = GFSIS[['Identificador', 'Tipo', 'Nome', 'Razão social', 'Tipo.1', 'CPF',
               'CNPJ', 'Limite de crédito', 'Percentual de comissão',
               'Tabela de preço geral', 'Contato', 'E-mail principal',
               'E-mails secundários', 'Gestor comercial', 'Gestor marketing',
               'Vendedor responsável', 'Data de credenciamento',
               'Próxima auditoria operacional', 'Última auditoria de manutenção',
               'Vencimento contrato de cessão de espaço',
               'Vencimento contrato de cessão de equipamento', 'Vencimento do alvará',
               'Tabela de credenciamento', 'Segmento', 'Consignação de mídia',
               'Valor do Start Fee', 'Outros valores', 'Descrição dos outros valores',
               'Telefone principal', 'Outros telefones','Telefones fornecidos pela AR',
               'Logradouro', 'Número', 'Complemento','Bairro', 'CEP', 'Município', 'UF',
               'Ativo','Possui integração renovação online', 'Sócio administrador',
               'Custo operacional', 'Origem do lead', 'Local de estoque','Forma de acerto', 'Banco',
               'Tipo de conta', 'Agência', 'Dvagencia','Conta', 'Dvconta', 'Chave Pix',
               'Tipo da chave pix', 'CNPJ Unidade', 'DOCUMENTO PA']]


# # Mesclagem Cadastro PA




CADASTRO_PA_Mescla1 = CADASTRO_PA.set_index('Nome_da_Empresa').join(CONTATOS_INFO_CADASTRO.set_index('Empresa'))





CADASTRO_PA_Mescla1.reset_index(inplace=True)





CADASTRO_PA_Mescla1 = CADASTRO_PA_Mescla1[~(CADASTRO_PA_Mescla1[['index', 'DOCUMENTO PA']].duplicated())]




CONTATOS_CONT_AGR.reset_index(inplace=True)





CADASTRO_PA_Mescla2 = CADASTRO_PA_Mescla1.set_index('index').join(CONTATOS_CONT_AGR.set_index('Empresa'))





CADASTRO_PA_Mescla2 = CADASTRO_PA_Mescla2[['ID','Tipo da empresa', 'Telefone de trabalho', 'Celular',
       'Email de trabalho', 'Criado', 'DOCUMENTO PA', 'Endereço',
       'Complemento', 'CEP', 'UF', 'Bairro', 'Cidade', 'Agente de Expansão',
       'CNAE PA', 'Tipo de Pessoa', 'Pessoa Responsável (CS)',
       'Nome do Proprietário', 'Telefone do Proprietário',
       'Email do Proprietário', 'Quantidade de AGRs']]





CADASTRO_PA_Mescla2.reset_index(inplace=True)





CADASTRO_PA_Mescla2 = CADASTRO_PA_Mescla2[~CADASTRO_PA_Mescla2[['index', 'DOCUMENTO PA']].duplicated()]





CADASTRO_PA_Mescla3 = CADASTRO_PA_Mescla2.set_index('index').join(NEGOCIOS_SUM_VALOR.set_index('Empresa'))




CADASTRO_PA_Mescla3.reset_index(inplace=True)





CADASTRO_PA_Mescla3 = CADASTRO_PA_Mescla3.rename(columns={'index':'Nome da Empresa'})



## Adicionando a classificação do Perfil do PA
Perfil_PA = pd.read_csv('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\Perfil do PA.csv', sep=";", header=3,dtype=str, encoding='utf-8')
Perfil_PA = Perfil_PA[~(Perfil_PA['Rótulos de Linha'].isna())]
Perfil_PA = Perfil_PA[['Rótulos de Linha','LETRA','NUMERO','CS','ID']]
CADASTRO_PA_Mescla4 = CADASTRO_PA_Mescla3.set_index('ID').join(Perfil_PA.set_index('ID'))
CADASTRO_PA_Mescla4 = CADASTRO_PA_Mescla4.rename(columns={'LETRA':'Classificação PAs'})
CADASTRO_PA_Mescla4.reset_index(inplace=True)
CADASTRO_PA_Mescla4 = CADASTRO_PA_Mescla4.rename(columns={'index':'DOCUMENTO PA'})
CADASTRO_PA_Mescla4 = CADASTRO_PA_Mescla4[['ID','Nome da Empresa','Tipo da empresa',
                                           'Telefone de trabalho','Celular',
                                           'Email de trabalho','Criado',
                                           'DOCUMENTO PA','Endereço',
                                           'Complemento','CEP',
                                           'UF','Bairro',
                                           'Cidade','Agente de Expansão',
                                           'CNAE PA','Tipo de Pessoa',
                                           'Pessoa Responsável (CS)','Nome do Proprietário',
                                           'Telefone do Proprietário','Email do Proprietário',
                                           'Quantidade de AGRs', ('Valor dos Produtos', 'sum'), 'Classificação PAs',
                                           'NUMERO','CS']]


### Adicionando CNAE - atividade
Atividade_PA = pd.read_csv('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\ATIVIDADE.csv', sep=";", encoding='utf-8')
Atividade_PA = Atividade_PA[['DOCUMENTO','ATIVIDADE']]
CADASTRO_PA_Mescla5 = CADASTRO_PA_Mescla4.set_index('DOCUMENTO PA').join(Atividade_PA.set_index('DOCUMENTO'))
CADASTRO_PA_Mescla5.reset_index(inplace=True)
CADASTRO_PA_Mescla5 = CADASTRO_PA_Mescla5.rename(columns={'index':'DOCUMENTO PA'})
CADASTRO_PA_Mescla5 = CADASTRO_PA_Mescla5[['ID','Nome da Empresa','Tipo da empresa',
                                           'Telefone de trabalho','Celular',
                                           'Email de trabalho','Criado',
                                           'DOCUMENTO PA','Endereço',
                                           'Complemento','CEP',
                                           'UF','Bairro',
                                           'Cidade','Agente de Expansão',
                                           'CNAE PA','Tipo de Pessoa',
                                           'Pessoa Responsável (CS)','Nome do Proprietário',
                                           'Telefone do Proprietário','Email do Proprietário',
                                           'Quantidade de AGRs', ('Valor dos Produtos', 'sum'), 'Classificação PAs',
                                          'NUMERO','CS','ATIVIDADE']]

# # Arquivos



# GFSIS
GFSIS_CSV = GFSIS.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\GFSIS.csv", sep=";", index=False, encoding='ANSI')
GFSIS_CSV
Duplicatas_GFSIS_CSV = Duplicatas_GFSIS.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\Duplicatas_GFSIS.csv", sep=";", index=False, encoding='ANSI')
Duplicatas_GFSIS_CSV





# CADASTRO PA
CADASTRO_PA_CSV = CADASTRO_PA_Mescla5.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\CADASTRO_PA.csv", sep=";", index=False, encoding='UTF-8')
CADASTRO_PA_CSV
Duplicatas_CADASTRO_PA_CSV = Duplicatas_CADASTRO_PA.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\Duplicatas_CADASTRO_PA.csv", sep=";", index=False, encoding='ANSI')
Duplicatas_CADASTRO_PA_CSV




# CONTATOS_INFO_CADASTRO
CONTATOS_INFO_CADASTRO_CSV = CONTATOS_INFO_CADASTRO.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\CONTATOS_INFO_CADASTRO.csv", sep=";", index=False, encoding='UTF-8')
CONTATOS_INFO_CADASTRO_CSV





Duplicatas_CONTATOS_INFO_CADASTRO_CSV = Duplicatas_CONTATOS_INFO_CADASTRO.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\Duplicatas_CONTATOS_INFO_CADASTRO.csv", sep=";", index=False, encoding='ANSI')
Duplicatas_CONTATOS_INFO_CADASTRO_CSV




# DEAL_NEGOCIOS_NOVO_BITRIX_CSV
DEAL_NEGOCIOS_NOVO_BITRIX_CSV = DEAL_NEGOCIOS_NOVO_BITRIX.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\DEAL_NEGOCIOS_NOVO_BITRIX.csv", sep=";", index=False, encoding='ANSI')
DEAL_NEGOCIOS_NOVO_BITRIX_CSV


# CONSULTA_NEGOCIOS_NOVO_BITRIX
CONSULTA_NEGOCIOS_NOVO_BITRIX_CSV = CONSULTA_NEGOCIOS_NOVO_BITRIX.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\CONSULTA_NEGOCIOS_NOVO_BITRIX.csv", sep=";", index=False, encoding='ANSI')
CONSULTA_NEGOCIOS_NOVO_BITRIX_CSV




# CONTATOS_CONT_AGR
CONTATOS_CONT_AGR_CSV = CONTATOS_CONT_AGR.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\CONTATOS_CONT_AGR.csv", sep=";", index=False, encoding='ANSI')
CONTATOS_CONT_AGR_CSV







# NEGOCIOS_SUM_VALOR
NEGOCIOS_SUM_VALOR_CSV = NEGOCIOS_SUM_VALOR.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\NEGOCIOS_SUM_VALOR.csv", sep=";", index=False, encoding='ANSI')
NEGOCIOS_SUM_VALOR_CSV




# EMISSOES
EMISSOES_CSV = EMISSOES.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\EMISSOES.csv",
                               decimal= '.',date_format = '%m/%d/%Y', sep=";", index=False, encoding='ANSI')
EMISSOES_CSV





# AGR
AGR_CSV = AGR.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\AGR.csv", sep=";", index=False, encoding='ANSI')
AGR_CSV



# PackControleDeVendas
PackControleDeVendas_CSV = PackControleDeVendas.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\PackControleDeVendas.csv", sep="|", index=False, encoding='ANSI')
PackControleDeVendas_CSV


# In[ ]:



# CONTATOS DE EMPRESAS
CONTATOS_CSV = CONTATOS.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\CONTATOS.csv", sep=";", index=False, encoding='UTF8')
CONTATOS_CSV

# CursosAvulsosControleDeVendas
CursosAvulsosControleDeVendas_CSV = CursosAvulsosControleDeVendas.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\CursosAvulsosControleDeVendas.csv", sep=";", index=False, encoding='ANSI')
CursosAvulsosControleDeVendas_CSV



## CENTRAL ##


import selenium
import time
import os
import shutil

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from datetime import datetime, date, timedelta
from selenium.webdriver.common.keys import Keys
import pyautogui #controlar o mouse
import pyperclip #controlar o teclado
import urllib #módul para trabalhar um URLs
import pandas as pd
import glob
import win32com.client
import datetime




options = webdriver.ChromeOptions()
preferences = {"download.default_directory": "C:\ONEDRIVE\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\Apuração de Resultados\Dashs datastudio\Central de Emissões", "safebrowsing.enabled": "false"}
options.add_experimental_option("prefs", preferences)

driver = webdriver.Chrome(executable_path=r'C:\chromedriver.exe', options=options)




dia = date.today()
# dia = sp_feriados.rollback(dia - pd.tseries.offsets.BusinessDay(n=1))

primeiroDia = dia.replace(day=1)
primeiroDia= primeiroDia.strftime("%d/%m/%Y")

# nxt_mnth = dia.replace(day=28) + datetime.timedelta(days=4)
# ultimoDia = nxt_mnth - datetime.timedelta(days=nxt_mnth.day)
# ultimoDia = ultimoDia.strftime("%d/%m/%Y")
# nxt_mnth = dia.replace(day=28) + datetime.timedelta(days=4)
ultimoDia = dia - datetime.timedelta(days=1)
ultimoDia = ultimoDia.strftime("%d/%m/%Y")

nome = dia.strftime('%m%Y')
print(nome)




# Acessando a GFSIS Nosso certificado
link = "https://nossocertificado.gfsis.com.br/gestaofacil/login/Index"
driver.get(link)

usuario = "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[1]/input"
time.sleep(2)
driver.find_element_by_xpath(usuario).send_keys("VICTOR.SOARES")
senha = "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[2]/input"
time.sleep(2)
driver.find_element_by_xpath(senha).send_keys("123456")
time.sleep(2)
button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[3]/div/input"))) 
button.click()
time.sleep(5)


# # Meus pedidos Nosso

# Acessando meus pedidos
link = "https://nossocertificado.gfsis.com.br/gestaofacil/login/faturamento/crud/PedidoVenda?ACAO=listagem"
driver.get(link)
time.sleep(5)

# Baixando Pedidos Aprovados    
        # Período de aprovação
data2 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/input[1]'
driver.find_element_by_xpath(data2).click()
driver.find_element_by_xpath(data2).send_keys(primeiroDia)
data3 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/input[3]'
driver.find_element_by_xpath(data3).click()
driver.find_element_by_xpath(data3).send_keys(ultimoDia)

        # Baixando Pedidos Aprovados
driver.find_element_by_xpath('//*[@id="aguardando_chk"]').click()
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys(Keys.DELETE)
        # Exportar aprovados
button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_exportar"]'))) 
button.click()
time.sleep(20)


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

pasta = "C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Aprovados\\Nosso"
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

# Baixando Pedidos Cancelados
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys('01/01/2020')
driver.find_element_by_xpath('//*[@id="confirmado_chk"]').click()
driver.find_element_by_xpath('//*[@id="cancelado_chk"]').click()
        # Limpar filtros
data3 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/input[3]'
driver.find_element_by_xpath(data3).click()
driver.find_element_by_xpath(data3).send_keys(Keys.DELETE)
data2 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/input[1]'
driver.find_element_by_xpath(data2).click()
driver.find_element_by_xpath(data2).send_keys(Keys.DELETE)
        # Filtros Avançados
driver.find_element_by_xpath('//*[@id="link_filtros_avancados"]').click()

        # Período de cancelamento
time.sleep(0.5)
data4 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[7]/td/div/table/tbody/tr[15]/td[2]/input[1]'
driver.find_element_by_xpath(data4).click()
driver.find_element_by_xpath(data4).send_keys(primeiroDia)
time.sleep(0.5)
data5 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[7]/td/div/table/tbody/tr[15]/td[2]/input[3]'
driver.find_element_by_xpath(data5).click()
driver.find_element_by_xpath(data5).send_keys(ultimoDia)

button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_exportar"]'))) 
button.click()
time.sleep(5)



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




# Videoconferência Nosso



# Acessando dados de atendimentos de Videoconferência
link = "https://nossocertificado.gfsis.com.br/gestaofacil/login/videoconferencia/crud/AtendimentoVideoconferencia?ACAO=listagem"
driver.get(link)




# Baixando os Videoconferência Aprovada
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys("01/10/2021")

driver.find_element_by_xpath('/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[4]/select/option[3]').click()
driver.find_element_by_xpath('//*[@id="link_filtros_avancados"]').click()

time.sleep(0.5)
        # período de Aprovação
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[8]/td/div/table/tbody/tr[3]/td[4]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys(primeiroDia)
data2 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[8]/td/div/table/tbody/tr[3]/td[4]/input[3]'
driver.find_element_by_xpath(data2).click()
driver.find_element_by_xpath(data2).send_keys(ultimoDia)

button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_exportar"]')))
button.click()

time.sleep(20)

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

pasta = "C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Aprovados\\Video\\Nosso"
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

# Videoconferência Nosso Cancelados
link = "https://nossocertificado.gfsis.com.br/gestaofacil/login/videoconferencia/crud/AtendimentoVideoconferencia?ACAO=listagem"
driver.get(link)

driver.find_element_by_xpath('//*[@id="btn_limpar"]').click()

# Baixando os Videoconferência Cancelados

data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys("01/10/2021")
driver.find_element_by_xpath('/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[4]/select/option[4]').click()

button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_exportar"]')))
button.click()

time.sleep(20)

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

pasta = "C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Cancelados\\Video\\Nosso"
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




# Acessando a GFSIS Certifica
# Acessando a GFSIS Certifica Brasil
link = "https://certificabrasil.gfsis.com.br/gestaofacil/login/Index"
driver.get(link)

usuario = "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[1]/input"
time.sleep(2)
driver.find_element_by_xpath(usuario).send_keys("VICTOR.SOARES")
senha = "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[2]/input"
time.sleep(2)
driver.find_element_by_xpath(senha).send_keys("123456")
time.sleep(2)
button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[3]/div/input"))) 
button.click()
time.sleep(5)


# # Meus pedidos Certifica

# Acessando meus pedidos
link = "https://certificabrasil.gfsis.com.br/gestaofacil/login/faturamento/crud/PedidoVenda?ACAO=listagem"
driver.get(link)
time.sleep(5)


# Baixando Pedidos Aprovados    
        # Período de aprovação
data2 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/input[1]'
driver.find_element_by_xpath(data2).click()
driver.find_element_by_xpath(data2).send_keys(primeiroDia)
data3 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/input[3]'
driver.find_element_by_xpath(data3).click()
driver.find_element_by_xpath(data3).send_keys(ultimoDia)

        # Baixando Pedidos Aprovados
driver.find_element_by_xpath('//*[@id="aguardando_chk"]').click()
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys(Keys.DELETE)
        # Exportar aprovados
button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_exportar"]'))) 
button.click()
time.sleep(20)


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

pasta = "C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Aprovados\\Certifica"
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

# Baixando Pedidos Cancelados
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys('01/01/2020')
driver.find_element_by_xpath('//*[@id="confirmado_chk"]').click()
driver.find_element_by_xpath('//*[@id="cancelado_chk"]').click()
        # Limpar filtros
data3 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/input[3]'
driver.find_element_by_xpath(data3).click()
driver.find_element_by_xpath(data3).send_keys(Keys.DELETE)
data2 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/input[1]'
driver.find_element_by_xpath(data2).click()
driver.find_element_by_xpath(data2).send_keys(Keys.DELETE)
        # Filtros Avançados
driver.find_element_by_xpath('//*[@id="link_filtros_avancados"]').click()

        # Período de cancelamento
time.sleep(0.5)
data4 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[7]/td/div/table/tbody/tr[15]/td[2]/input[1]'
driver.find_element_by_xpath(data4).click()
driver.find_element_by_xpath(data4).send_keys(primeiroDia)
time.sleep(0.5)
data5 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[7]/td/div/table/tbody/tr[15]/td[2]/input[3]'
driver.find_element_by_xpath(data5).click()
driver.find_element_by_xpath(data5).send_keys(ultimoDia)

button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_exportar"]'))) 
button.click()
time.sleep(5)



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

# Videoconferência Nosso



# Acessando dados de atendimentos de Videoconferência
link = "https://certificabrasil.gfsis.com.br/gestaofacil/login/videoconferencia/crud/AtendimentoVideoconferencia?ACAO=listagem"
driver.get(link)




# Baixando os Videoconferência Aprovada
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys("01/10/2021")

driver.find_element_by_xpath('/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[4]/select/option[3]').click()
driver.find_element_by_xpath('//*[@id="link_filtros_avancados"]').click()


time.sleep(0.5)
        # período de Aprovação
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[8]/td/div/table/tbody/tr[3]/td[4]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys(primeiroDia)
data2 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[8]/td/div/table/tbody/tr[3]/td[4]/input[3]'
driver.find_element_by_xpath(data2).click()
driver.find_element_by_xpath(data2).send_keys(ultimoDia)

button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_exportar"]')))
button.click()

time.sleep(20)

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

pasta = "C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Aprovados\\Video\\Certifica"
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

# Videoconferência Nosso Cancelados
link = "https://certificabrasil.gfsis.com.br/gestaofacil/login/videoconferencia/crud/AtendimentoVideoconferencia?ACAO=listagem"
driver.get(link)

driver.find_element_by_xpath('//*[@id="btn_limpar"]').click()

# Baixando os Videoconferência Cancelados
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys("01/10/2021")
driver.find_element_by_xpath('/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[4]/select/option[4]').click()

button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_exportar"]')))
button.click()

time.sleep(20)

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

pasta = "C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Cancelados\\Video\\Certifica"
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




# Acessando a GFSIS Digtec
# Acessando a GFSIS DIGITEC
link = "https://digtec.gfsis.com.br/gestaofacil/login/faturamento/crud/PontoAtendimento?ACAO=listagem"
driver.get(link)

usuario = "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[1]/input"
time.sleep(2)
driver.find_element_by_xpath(usuario).send_keys("VICTOR.RAYOL")
senha = "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[2]/input"
time.sleep(2)
driver.find_element_by_xpath(senha).send_keys("123456")
time.sleep(2)
button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td/div/div/div/div[2]/form/div[3]/div/input"))) 
button.click()
time.sleep(5)


# # Meus pedidos DIGTEC


# Acessando meus pedidos
link = "https://digtec.gfsis.com.br/gestaofacil/login/faturamento/crud/PedidoVenda?ACAO=listagem"
driver.get(link)
time.sleep(5)


# Baixando Pedidos Aprovados    
        # Período de aprovação
data2 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/input[1]'
driver.find_element_by_xpath(data2).click()
driver.find_element_by_xpath(data2).send_keys(primeiroDia)
data3 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/input[3]'
driver.find_element_by_xpath(data3).click()
driver.find_element_by_xpath(data3).send_keys(ultimoDia)

        # Baixando Pedidos Aprovados
driver.find_element_by_xpath('//*[@id="aguardando_chk"]').click()
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys(Keys.DELETE)
        # Exportar aprovados
button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_exportar"]'))) 
button.click()
time.sleep(20)


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

pasta = "C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Aprovados\\Digtec"
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

# Baixando Pedidos Cancelados
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys('01/01/2020')
driver.find_element_by_xpath('//*[@id="confirmado_chk"]').click()
driver.find_element_by_xpath('//*[@id="cancelado_chk"]').click()
        # Limpar filtros
data3 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/input[3]'
driver.find_element_by_xpath(data3).click()
driver.find_element_by_xpath(data3).send_keys(Keys.DELETE)
data2 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/input[1]'
driver.find_element_by_xpath(data2).click()
driver.find_element_by_xpath(data2).send_keys(Keys.DELETE)
        # Filtros Avançados
driver.find_element_by_xpath('//*[@id="link_filtros_avancados"]').click()

        # Período de cancelamento
time.sleep(0.5)
data4 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[7]/td/div/table/tbody/tr[15]/td[2]/input[1]'
driver.find_element_by_xpath(data4).click()
driver.find_element_by_xpath(data4).send_keys(primeiroDia)
time.sleep(0.5)
data5 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[7]/td/div/table/tbody/tr[15]/td[2]/input[3]'
driver.find_element_by_xpath(data5).click()
driver.find_element_by_xpath(data5).send_keys(ultimoDia)

button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_exportar"]'))) 
button.click()
time.sleep(5)



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

pasta = "C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Cancelados\\Digtec"
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

# Videoconferência Nosso



# Acessando dados de atendimentos de Videoconferência
link = "https://digtec.gfsis.com.br/gestaofacil/login/videoconferencia/crud/AtendimentoVideoconferencia?ACAO=listagem"
driver.get(link)



# Baixando os Videoconferência Aprovada
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys("01/10/2021")

driver.find_element_by_xpath('/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[4]/select/option[3]').click()
driver.find_element_by_xpath('//*[@id="link_filtros_avancados"]').click()


time.sleep(0.5)
        # período de Aprovação
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[8]/td/div/table/tbody/tr[3]/td[4]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys(primeiroDia)
data2 = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[8]/td/div/table/tbody/tr[3]/td[4]/input[3]'
driver.find_element_by_xpath(data2).click()
driver.find_element_by_xpath(data2).send_keys(ultimoDia)

button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_exportar"]')))
button.click()

time.sleep(20)

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

pasta = "C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Aprovados\\Video\\Digtec"
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

# Videoconferência Nosso Cancelados
link = "https://digtec.gfsis.com.br/gestaofacil/login/videoconferencia/crud/AtendimentoVideoconferencia?ACAO=listagem"
driver.get(link)

driver.find_element_by_xpath('//*[@id="btn_limpar"]').click()

# Baixando os Videoconferência Cancelados
data = '/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input[1]'
driver.find_element_by_xpath(data).click()
driver.find_element_by_xpath(data).send_keys("01/10/2021")
driver.find_element_by_xpath('/html/body/div[2]/div/form/div/div[2]/table[1]/tbody/tr/td/table[1]/tbody/tr[3]/td[4]/select/option[4]').click()

button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_exportar"]')))
button.click()

time.sleep(20)

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

pasta = "C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Cancelados\\Video\\Digtec"
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


driver.quit()
time.sleep(10)



# Nosso

# In[125]:


# Meus pedidos NOSSO - Aprovados
arquivos = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Aprovados\\Nosso\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
Pedidos_Nosso = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    Pedidos_Nosso.append(temp_df)

Pedidos_Nosso = pd.concat(Pedidos_Nosso, axis=0)

# Video NOSSO - Aprovados
arquivos = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Aprovados\\Video\\Nosso\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
Video_Nosso = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    Video_Nosso.append(temp_df)

Video_Nosso = pd.concat(Video_Nosso, axis=0)
Video_Nosso.drop(['Cliente', 'Indicação', 'Itens do pedido de venda', 'Data'], axis = 1, inplace=True)
Pedidos_Nosso['Cliente'] = Pedidos_Nosso['Nome'] + ' (' + Pedidos_Nosso['CPF/CNPJ'] + ')'
Video_Nosso = Video_Nosso.set_index('Pedido de venda').join(Pedidos_Nosso.set_index('Identificador'))
Video_Nosso.reset_index(inplace=True)
Video_Nosso = Video_Nosso[['Id', 'Pedido de venda', 'Etapa', 'Cliente', 'Indicação', 'Itens do pedido de venda', 'Tipo de emissão',
              'Situação', 'Responsável pré-atendimento', 'Data', 'Status', 'Responsável atendimento', 'Data.1', 'Status.1',
              'Data e hora criação do pedido', 'Data de cancelamento', 'Data de aprovação', 'Formas de pagamento do pedido de venda',
                           'Situação do documento','Responsável Pós-atendimento', 'Data.2', 'Status.2']]
# Video_Certifica.rename(columns = {'index':'Pedido de venda'}, inplace=True)
Video_Nosso['ConcID'] = Video_Nosso['Pedido de venda'] + "NOSSO"


# In[126]:


# Meus pedidos NOSSO - Cancelados
arquivos = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Cancelados\\Nosso\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
Pedidos_Nosso_Can = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    Pedidos_Nosso_Can.append(temp_df)
    
Pedidos_Nosso_Can = pd.concat(Pedidos_Nosso_Can, axis=0)

Pedidos_Nosso_Can['Cliente'] = Pedidos_Nosso_Can['Nome'] + ' (' + Pedidos_Nosso_Can['CPF/CNPJ'] + ')'

# Video NOSSO - Cancelados
arquivos = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Cancelados\\Video\\Nosso\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
Video_Nosso_Can = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    Video_Nosso_Can.append(temp_df)

Video_Nosso_Can = pd.concat(Video_Nosso_Can, axis=0)
Video_Nosso_Can.drop(['Cliente', 'Indicação', 'Itens do pedido de venda', 'Data'], axis = 1, inplace=True)

Video_Nosso_Can = Video_Nosso_Can.set_index('Pedido de venda').join(Pedidos_Nosso_Can.set_index('Identificador'))
Video_Nosso_Can.reset_index(inplace=True)

Video_Nosso_Can.rename(columns = {'index':'Pedido de venda'}, inplace=True)

Video_Nosso_Can = Video_Nosso_Can[['Id', 'Pedido de venda', 'Etapa', 'Cliente', 'Indicação', 'Itens do pedido de venda', 'Tipo de emissão',
              'Situação', 'Responsável pré-atendimento', 'Data', 'Status', 'Responsável atendimento', 'Data.1', 'Status.1',
              'Data e hora criação do pedido', 'Data de cancelamento', 'Data de aprovação', 'Formas de pagamento do pedido de venda',
                           'Situação do documento','Responsável Pós-atendimento', 'Data.2', 'Status.2']]
Video_Nosso_Can['ConcID'] = Video_Nosso_Can['Pedido de venda'] + "NOSSO"

Video_Nosso_Can = Video_Nosso_Can.dropna(subset=['Data de cancelamento'])

Video_Nosso = Video_Nosso.append(Video_Nosso_Can)


# Certifica

# In[127]:


# Meus pedidos Certifica - Aprovados
arquivos = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Aprovados\\Certifica\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
Pedidos_Certifica = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    Pedidos_Certifica.append(temp_df)

Pedidos_Certifica = pd.concat(Pedidos_Certifica, axis=0)

# Video Certifica - Aprovados
arquivos = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Aprovados\\Video\\Certifica\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
Video_Certifica = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    Video_Certifica.append(temp_df)

Video_Certifica = pd.concat(Video_Certifica, axis=0)
Video_Certifica.drop(['Cliente', 'Indicação', 'Itens do pedido de venda', 'Data'], axis = 1, inplace=True)
Pedidos_Certifica['Cliente'] = Pedidos_Certifica['Nome'] + ' (' + Pedidos_Certifica['CPF/CNPJ'] + ')'
Video_Certifica = Video_Certifica.set_index('Pedido de venda').join(Pedidos_Certifica.set_index('Identificador'))
Video_Certifica.reset_index(inplace=True)
Video_Certifica = Video_Certifica[['Id', 'Pedido de venda', 'Etapa', 'Cliente', 'Indicação', 'Itens do pedido de venda', 'Tipo de emissão',
              'Situação', 'Responsável pré-atendimento', 'Data', 'Status', 'Responsável atendimento', 'Data.1', 'Status.1',
              'Data e hora criação do pedido', 'Data de cancelamento', 'Data de aprovação', 'Formas de pagamento do pedido de venda',
                           'Situação do documento','Responsável Pós-atendimento', 'Data.2', 'Status.2']]
# Video_Certifica.rename(columns = {'index':'Pedido de venda'}, inplace=True)
Video_Certifica['ConcID'] = Video_Certifica['Pedido de venda'] + "CERTIFICA"


# In[128]:


# Meus pedidos Certifica - Cancelados
arquivos = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Cancelados\\Certifica\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
Pedidos_Certifica_Can = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    Pedidos_Certifica_Can.append(temp_df)
    
Pedidos_Certifica_Can = pd.concat(Pedidos_Certifica_Can, axis=0)

Pedidos_Certifica_Can['Cliente'] = Pedidos_Certifica_Can['Nome'] + ' (' + Pedidos_Certifica_Can['CPF/CNPJ'] + ')'

# Video Certifica - Cancelados
arquivos = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Cancelados\\Video\\Certifica\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
Video_Certifica_Can = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    Video_Certifica_Can.append(temp_df)

Video_Certifica_Can = pd.concat(Video_Certifica_Can, axis=0)
Video_Certifica_Can.drop(['Cliente', 'Indicação', 'Itens do pedido de venda', 'Data'], axis = 1, inplace=True)

Video_Certifica_Can = Video_Certifica_Can.set_index('Pedido de venda').join(Pedidos_Certifica_Can.set_index('Identificador'))
Video_Certifica_Can.reset_index(inplace=True)

Video_Certifica_Can.rename(columns = {'index':'Pedido de venda'}, inplace=True)

Video_Certifica_Can = Video_Certifica_Can[['Id', 'Pedido de venda', 'Etapa', 'Cliente', 'Indicação', 'Itens do pedido de venda', 'Tipo de emissão',
              'Situação', 'Responsável pré-atendimento', 'Data', 'Status', 'Responsável atendimento', 'Data.1', 'Status.1',
              'Data e hora criação do pedido', 'Data de cancelamento', 'Data de aprovação', 'Formas de pagamento do pedido de venda',
                           'Situação do documento','Responsável Pós-atendimento', 'Data.2', 'Status.2']]
Video_Certifica_Can['ConcID'] = Video_Certifica_Can['Pedido de venda'] + "CERTIFICA"

Video_Certifica_Can = Video_Certifica_Can.dropna(subset=['Data de cancelamento'])

Video_Certifica = Video_Certifica.append(Video_Certifica_Can)


# DIGTEC

# In[132]:


# Meus pedidos Digtec - Aprovados
arquivos = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Aprovados\\Digtec\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
Pedidos_Digtec = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    Pedidos_Digtec.append(temp_df)

Pedidos_Digtec = pd.concat(Pedidos_Digtec, axis=0)

# Video Digtec - Aprovados
arquivos = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Aprovados\\Video\\Digtec\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
Video_Digtec = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    Video_Digtec.append(temp_df)

Video_Digtec = pd.concat(Video_Digtec, axis=0)
Video_Digtec.drop(['Cliente', 'Indicação', 'Itens do pedido de venda', 'Data'], axis = 1, inplace=True)
Pedidos_Digtec['Cliente'] = Pedidos_Digtec['Nome'] + ' (' + Pedidos_Digtec['CPF/CNPJ'] + ')'
Video_Digtec = Video_Digtec.set_index('Pedido de venda').join(Pedidos_Digtec.set_index('Identificador'))
Video_Digtec.reset_index(inplace=True)
Video_Digtec = Video_Digtec[['Id', 'Pedido de venda', 'Etapa', 'Cliente', 'Indicação', 'Itens do pedido de venda', 'Tipo de emissão',
              'Situação', 'Responsável pré-atendimento', 'Data', 'Status', 'Responsável atendimento', 'Data.1', 'Status.1',
              'Data e hora criação do pedido', 'Data de cancelamento', 'Data de aprovação', 'Formas de pagamento do pedido de venda',
                           'Situação do documento','Responsável Pós-atendimento', 'Data.2', 'Status.2']]
# Video_Certifica.rename(columns = {'index':'Pedido de venda'}, inplace=True)
Video_Digtec['ConcID'] = Video_Digtec['Pedido de venda'] + "DIGTEC"


# In[133]:


# Meus pedidos Digtec - Cancelados
arquivos = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Cancelados\\Digtec\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
Pedidos_Digtec_Can = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    Pedidos_Digtec_Can.append(temp_df)
    
Pedidos_Digtec_Can = pd.concat(Pedidos_Digtec_Can, axis=0)

Pedidos_Digtec_Can['Cliente'] = Pedidos_Digtec_Can['Nome'] + ' (' + Pedidos_Digtec_Can['CPF/CNPJ'] + ')'

# Video Digtec - Cancelados
arquivos = glob.glob('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Cancelados\\Video\\Digtec\\*.csv')
# 'arquivos' agora é um array com o nome de todos os .csv que começam com 'arquivo'
Video_Digtec_Can = []

for x in arquivos:
    temp_df = pd.read_csv(x, sep=";", dtype=str, encoding='ANSI' )
    Video_Digtec_Can.append(temp_df)

Video_Digtec_Can = pd.concat(Video_Digtec_Can, axis=0)
Video_Digtec_Can.drop(['Cliente', 'Indicação', 'Itens do pedido de venda', 'Data'], axis = 1, inplace=True)

Video_Digtec_Can = Video_Digtec_Can.set_index('Pedido de venda').join(Pedidos_Digtec_Can.set_index('Identificador'))
Video_Digtec_Can.reset_index(inplace=True)

Video_Digtec_Can.rename(columns = {'index':'Pedido de venda'}, inplace=True)

Video_Digtec_Can = Video_Digtec_Can[['Id', 'Pedido de venda', 'Etapa', 'Cliente', 'Indicação', 'Itens do pedido de venda', 'Tipo de emissão',
              'Situação', 'Responsável pré-atendimento', 'Data', 'Status', 'Responsável atendimento', 'Data.1', 'Status.1',
              'Data e hora criação do pedido', 'Data de cancelamento', 'Data de aprovação', 'Formas de pagamento do pedido de venda',
                           'Situação do documento','Responsável Pós-atendimento', 'Data.2', 'Status.2']]
Video_Digtec_Can['ConcID'] = Video_Digtec_Can['Pedido de venda'] + "DIGTEC"

Video_Digtec_Can = Video_Digtec_Can.dropna(subset=['Data de cancelamento'])

Video_Digtec = Video_Digtec.append(Video_Digtec_Can)


# In[134]:



Videcoferencia = pd.concat([Video_Certifica, Video_Nosso, Video_Digtec], ignore_index=True)


# In[107]:



# In[49]:


EMISSOES = pd.read_csv("C:\\ONEDRIVE\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Base para Gsheets\\CSV_BUC\\EMISSOES.csv", sep=";", dtype=str, encoding="ANSI")


# In[50]:


EMISSOES['ConcID'] = EMISSOES['Identificador'] + EMISSOES['AR']


# In[135]:

EMISSOES2 = EMISSOES[[ "Data de aprovação", "Situação", "A quem cobrar?", "PREÇO VENDA", 
                     "REPASSE AE ou PE", "REPASSE GE", "DESPESA BOLETO", "DESPESA IMPOSTOS", "CUSTO ULT FAIXA", 
                     "REPASSE PARCEIRO INDICADOR", "REPASSE PARCEIRO INDICADO", "ConcID", 'AR']]

EMISSOES = EMISSOES[[ "A quem cobrar?", "PREÇO VENDA", 
                     "REPASSE AE ou PE", "REPASSE GE", "DESPESA BOLETO", "DESPESA IMPOSTOS", "CUSTO ULT FAIXA", 
                     "REPASSE PARCEIRO INDICADOR", "REPASSE PARCEIRO INDICADO", "ConcID", 'AR']]


EMISSOES2.drop_duplicates(subset=['ConcID'], inplace=True)

# In[52]:


EMISSOES.drop_duplicates(subset=['ConcID'], inplace=True)


# In[53]:


# In[108]:


EMISSOES2 = EMISSOES2.rename(columns = {"Data de aprovação": "Aprovação"})
EMISSOES2['Aprovação'] = pd.to_datetime(EMISSOES2['Aprovação'],  errors='ignore', dayfirst = False )

EMISSOES2 = EMISSOES2.rename(columns = {"Formas de pagamento do pedido de venda": "Formas de Pag",
                                        "A quem cobrar?": "Unidade", "REPASSE AE ou PE": "Custo AE/PE",
                                        "REPASSE GE": "Custo GE", "DESPESA": "Custo Boleto", 
                                        "DESPESA IMPOSTOS": "Custo Imposto", "CUSTO ULT FAIXA": "Custo Certificado",
                                        "REPASSE PARCEIRO INDICADOR": "Custo Franquia NTW", 
                                        "REPASSE PARCEIRO INDICADO": "Custo Franqueado"})


# In[109]:



Videcoferencia2 = Videcoferencia.set_index('ConcID').join(EMISSOES.set_index('ConcID'))


# In[54]:


Videcoferencia2.reset_index(inplace=True)
Videcoferencia2 = Videcoferencia2.rename(columns = {"Data de aprovação": "Aprovação" ,
                                      "Formas de pagamento do pedido de venda": "Formas de Pag", 
                                      "A quem cobrar?": "Unidade", "REPASSE AE ou PE": "Custo AE/PE",
                                      "REPASSE GE": "Custo GE", "DESPESA": "Custo Boleto", 
                                      "DESPESA IMPOSTOS": "Custo Imposto", "CUSTO ULT FAIXA": "Custo Certificado",
                                      "REPASSE PARCEIRO INDICADOR": "Custo Franquia NTW", 
                                      "REPASSE PARCEIRO INDICADO": "Custo Franqueado"})


# In[110]:


Videcoferencia2['Aprovação'] = pd.to_datetime(Videcoferencia2['Aprovação'],  errors='ignore', dayfirst = True )


# In[111]:


EMISSOES2 = EMISSOES2[EMISSOES2['AR'] == 'DIGTEC']


# In[112]:


Videcoferencia3 = Videcoferencia2.append(EMISSOES2, ignore_index=True)


# In[114]:


Videcoferencia3.drop_duplicates(subset = ['ConcID'], inplace= True)


# In[115]:


Videcoferencia3 = Videcoferencia3[["ConcID" , "Id" , "Pedido de venda" , "Etapa" , "Cliente" , "Indicação" , 
                           "Itens do pedido de venda" , "Tipo de emissão" , "Situação" , 
                           "Responsável pré-atendimento" , "Data" , "Status" , "Responsável atendimento" , 
                           "Data.1" , "Status.1" , "Data e hora criação do pedido" , "Data de cancelamento" , 
                           "Aprovação" , "Formas de Pag" , "Situação do documento" , "Responsável Pós-atendimento" , 
                           "Data.2" , "Status.2" , "Unidade" , "PREÇO VENDA" , "Custo AE/PE" , "Custo GE" , 
                           "DESPESA BOLETO" , "Custo Imposto" , "Custo Certificado" , "Custo Franquia NTW" , 
                           "Custo Franqueado"]]


# In[116]:


Video_CSV = Videcoferencia3.to_csv("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Arquivos\\Videcoferencia.csv", sep=";", index=False, encoding='ANSI')
Video_CSV


# In[11]:


### Abrir Apuração de emissões
xlapp = win32com.client.DispatchEx("Excel.Application")
xlapp.Visible = True
xlapp.DisplayAlerts = False
wb = xlapp.Workbooks.Open("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Dashs datastudio\\Central de Emissões\\Análise Central de Emissões.xlsx")

time.sleep(40)


# In[41]:


wb.RefreshAll()
time.sleep(120)
wb.Save()
time.sleep(30)





wb.Close()
time.sleep(5)
xlapp.Quit()





#ABRIR BUC
os.startfile(r'C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\Dashs datastudio\\Base para Gsheets\\BUC - Cadastros.xlsm')




# In[ ]:




