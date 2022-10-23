import imaplib
import email
from email.header import decode_header
import os


    ## ACESSAR O E-MAIL
    
    # credenciais da conta
    username = "EMAIL"
    password = "SENHA"

    def clean(text):
        # texto limpo para criar uma pasta
        return "".join(c if c.isalnum() else "_" for c in text)

    # Crie uma classe IMAP4 com SSL
    imap = imaplib.IMAP4_SSL("outlook.office365.com")
    # Autenticar conta
    imap.login(username, password)
    
    status, messages = imap.select("INBOX")
    # Número dos e-mails mais recentes a serem buscados
    N = 1
    # Baixar e-mails
    messages = int(messages[0])

    for i in range(messages, messages-N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # Analisar um e-mail de bytes em um objeto de mensagem
                msg = email.message_from_bytes(response[1])
                # Decodificar o assunto do e-mail
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # Decodificar remetente de e-mail
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                print("Subject:", subject)
                print("From:", From)
                # Se a mensagem de e-mail tiver várias partes
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # Extrair tipo de conteúdo de e-mail
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # Obter o corpo do e-mail
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
                    # Extrair tipo de conteúdo de e-mail
                    content_type = msg.get_content_type()
                    # Obter o corpo do e-mail
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # Imprimir apenas partes de e-mail de texto
                        print(body)
                if content_type == "text/html":
                    # Se for HTML, crie um novo arquivo HTML e abra-o no navegador
                    folder_name = clean(subject)
                    if not os.path.isdir(folder_name):
                        # make a folder for this email (named after the subject)
                        os.mkdir(folder_name)
                    filename = "index.html"
                    filepath = os.path.join(folder_name, filename)
                    # Escreva o arquivo
                    open(filepath, "w").write(body)
                print("="*100)
    # Encerre a conexão e saia
    imap.close()
    imap.logout()
    
    html_text = open(filepath).read()
    text_filtered = re.sub(r'<(.*?)>', '', html_text)
    
   # Código extrair os dígitos do TOKEN
    codigo = text_filtered.split()[-1]
    codigo = codigo[-6:]
    print(codigo)
    token = driver.find_element(By.XPATH, value ='/html/body/div[2]/div/div[1]/div/div/div/div[2]/div/div/input')
    token.send_keys(codigo)
    
    # Deletar e-mail baixado
    os.remove(filepath)
    os.removedirs(folder_name)
