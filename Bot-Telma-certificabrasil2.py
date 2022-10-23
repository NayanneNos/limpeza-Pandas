#!/usr/bin/env python
# coding: utf-8

# In[1]:


#!/usr/bin/env python
# coding: utf-8

# In[5]:


#!/usr/bin/env python
# coding: utf-8
# !pip install autoit
# !pip3 install webdriver_manager
# In[1]:

import os

os.system("TASKKILL /F /IM Nosso3.exe")
os.system("TASKKILL /F /IM CERTIFICABRASIL3.exe")
os.system("TASKKILL /F /IM TokenCertificado.exe")
os.system("TASKKILL /F /IM TokenNosso.exe")

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException
import pyautogui
import time 
import autoit
import re
import imaplib
import email
from email.header import decode_header
import webbrowser
import functools
import operator
import datetime
import pygetwindow as gw



# In[2]:



t_agora = datetime.datetime.now()

print("Loop executado:", t_agora)  


# In[26]:


# entrada = str(input("Digite um horário para encerrar o processo de libração de inventário de máquina (HH:MM)\n"))
# hr = entrada.split(':')
hr = str('19:00')
hr = hr.split(':')
t_desp = datetime.datetime.combine( datetime.datetime.now().date(),
                                    datetime.time( int(hr[0]), int(hr[1])) )


# In[27]:

print("Parar execução: ", t_desp )

# In[6]:


# In[3]:



options = webdriver.ChromeOptions()
preferences = {}
options.add_experimental_option("prefs", preferences)
driver = webdriver.Chrome(executable_path=r'C:\chromedriver.exe', options=options)


# In[4]:



# In[7]:


driver.get("https://arcertificabrasil3.acsoluti.com.br/certdig/fechamento")


# In[8]:


time.sleep(5)


# In[9]:


driver.get("https://arcertificabrasil3.acsoluti.com.br/certdig/fechamento")


# In[10]:


time.sleep(5)
title = 'AR Nosso - Google Chrome'

window = gw.getWindowsWithTitle('AR Nosso - Google Chrome')[0]
window.minimize()
window.maximize()

time.sleep(0.5)

autoit.run("C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Script_Python\\LiberacaoDeMaquina\\CERTIFICABRASIL3.exe")



time.sleep(7)

title = 'AR Nosso - Google Chrome'

window = gw.getWindowsWithTitle('AR Nosso - Google Chrome')[0]
window.minimize()
window.maximize()
time.sleep(0.5)

autoit.run("C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Script_Python\\LiberacaoDeMaquina\\TokenCertificado.exe")







time.sleep(5)


# In[13]:





# In[ ]:


# In[14]:


time.sleep(10)


# ## Código

# In[10]:


# renviar código, e quando já logado ir para a página
try:

    reenviar = driver.find_element(By.XPATH, value ='/html/body/div[2]/div/div[1]/div/div/div/div[4]/div[2]/a')
    reenviar.click()
    
    time.sleep(20)
    ## ACESSAR O E-MAIL
    
    # account credentials
    username = "apuracao.nosso@outlook.com"
    password = "Nosso@2021"

    def clean(text):
        # clean text for creating a folder
        return "".join(c if c.isalnum() else "_" for c in text)

    # create an IMAP4 class with SSL 
    imap = imaplib.IMAP4_SSL("outlook.office365.com")
    # authenticate
    imap.login(username, password)
    
    status, messages = imap.select("INBOX")
    # number of top emails to fetch
    N = 1
    # total number of emails
    messages = int(messages[0])

    for i in range(messages, messages-N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                print("Subject:", subject)
                print("From:", From)
                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            print(body)
                        elif "attachment" in content_disposition:
                            # baixar anexo
                            filename = part.get_filename()
                            if filename:
                                folder_name = clean(subject)
                                if not os.path.isdir(folder_name):
                                    # crie uma pasta para este e-mail (com o nome do assunto)
                                    os.mkdir(folder_name)
                                filepath = os.path.join(folder_name, filename)
                                # baixe o anexo e salve
                                open(filepath, "wb").write(part.get_payload(decode=True))
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # print only text email parts
                        print(body)
                if content_type == "text/html":
                    # if it's HTML, create a new HTML file and open it in browser
                    folder_name = clean(subject)
                    if not os.path.isdir(folder_name):
                        # make a folder for this email (named after the subject)
                        os.mkdir(folder_name)
                    filename = "index.html"
                    filepath = os.path.join(folder_name, filename)
                    # write the file
                    open(filepath, "w").write(body)
                print("="*100)
    # close the connection and logout
    imap.close()
    imap.logout()
    
    html_text = open(filepath).read()
    text_filtered = re.sub(r'<(.*?)>', '', html_text)
    
   
    codigo = text_filtered.split()[-1]
    codigo = codigo[-6:]
    print(codigo)
    token = driver.find_element(By.XPATH, value ='/html/body/div[2]/div/div[1]/div/div/div/div[2]/div/div/input')
    token.send_keys(codigo)
    
    # Deletar e-mail baixado
    os.remove(filepath)
    os.removedirs(folder_name)
    
    Confirma = driver.find_element(By.XPATH, value ='/html/body/div[2]/div/div[1]/div/div/div/div[4]/div[1]/a')
    Confirma.click()
    
except NoSuchElementException:
    pass


# In[13]:


time.sleep(10)



# In[ ]:


# ### Liberação dos AGR´s

# In[16]:


driver.get("https://arcertificabrasil3.acsoluti.com.br/inventario-maquina/autorizar-inventario")


# In[17]:


time.sleep(7)


# In[25]:








def executeSomething():

    try:

        while True:

            agora = datetime.datetime.now()
            agr = driver.find_element(By.XPATH, value ='//*[@id="row_0"]/td[2]')
            AGR= agr.text
            texto = ("Máquina de ", AGR, " liberada com sucesso pelo processo do BOT no dia ", agora.strftime("%d/%m/%Y"), " às ", agora.strftime("%H:%M"))
            arquivo = open(os.path.join('C:\\OneDrive\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Script_Python\\LiberacaoDeMaquina\\CertificaBrasil.txt'), 'a')
            texto = functools.reduce(operator.add, (texto))
            arquivo.write(texto)
            arquivo.write("\n")
            
            Autorizar = driver.find_element(By.XPATH, value = '/html/body/div[2]/div/div[1]/div/div[1]/div[2]/table/tbody/tr/td[1]/button[1]')
            Loop = Autorizar.click()
            Loop
            time.sleep(2)
            #seleções
            sel1 = driver.find_element(By.XPATH, value ='//*[@id="check1"]')
            sel2 = driver.find_element(By.XPATH, value ='//*[@id="check2"]')
            sel3 = driver.find_element(By.XPATH, value ='//*[@id="check3"]')
            sel4 = driver.find_element(By.XPATH, value ='//*[@id="check4"]')
            sel5 = driver.find_element(By.XPATH, value ='//*[@id="check5"]')
            sel6 = driver.find_element(By.XPATH, value ='//*[@id="check6"]')
            sel7 = driver.find_element(By.XPATH, value ='//*[@id="check7"]')
            sel8 = driver.find_element(By.XPATH, value ='//*[@id="check8"]')
            sel9 = driver.find_element(By.XPATH, value ='//*[@id="check9"]')
            sel10 = driver.find_element(By.XPATH, value ='//*[@id="check10"]')
            sel11 = driver.find_element(By.XPATH, value ='//*[@id="check11"]')

            sel1.click()
            sel2.click()
            sel3.click()
            sel4.click()
            sel5.click()
            sel6.click()
            sel7.click()
            sel8.click()
            sel9.click()
            sel10.click()
            sel11.click()
            time.sleep(1)

            SIM = driver.find_element(By.XPATH, value ='//*[@id="dialogInformacaoMaquina"]/div[2]/div[1]/div/div[2]/button')
            SIM.click()

            time.sleep(3)

            driver.refresh()


    except NoSuchElementException:
        time.sleep(3)
        driver.refresh()


# In[29]:


while datetime.datetime.now() < t_desp:

    executeSomething()

else:
    print("Fim do processo de liberação de máquina!")
    driver.quit()


