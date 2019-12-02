import imaplib # Pacote para instalar o IMAP
import base64 # pacote para codificar e decodificar dados binarios para ASCII
import os # pacote usado para manipular diretorios localmente
import email # pacote utilizado para ler, escrever e enviar e-mails

print('Script utilizado para automatizar o download dos relatorios diarios da Amil.')
email_user = input('Login: ')
email_pass = input('Password: ')


mail = imaplib.IMAP4_SSL("imap-mail.outlook.com",993)
mail.login(email_user, email_pass)
mail.list()
mail.select('inbox')

print('Digite o endereco de e-mail da pessoa que voce quer baixar os arquivos')
email_adress = input('E-mail: ')
result, data = mail.search(None, '(FROM "' + email_adress + '")')
 
ids = data[0] 
id_list = ids.split()
latest_email_id = id_list[-1] #Pega o ultimo email

typ, data = mail.fetch(latest_email_id, '(RFC822)' )
raw_email = data[0][1]

# converte de byte para string
raw_email_string = raw_email.decode('utf-8')
email_message = email.message_from_string(raw_email_string)


print('Digite o caminho para salvar o arquivo')
path_files = input('Caminho: ')
# downloading attachments
for part in email_message.walk():
    # this part comes from the snipped I don't understand yet... 
    if part.get_content_maintype() == 'multipart':
        continue
    if part.get('Content-Disposition') is None:
        continue
    fileName = part.get_filename()
    if bool(fileName):
        filePath = os.path.join(path_files, fileName)
        
        fp = open(filePath, 'wb')
        fp.write(part.get_payload(decode=True))
        fp.close()
        subject = str(email_message).split("Subject: ", 1)[1].split("\nTo:", 1)[0]
        print('Downloaded "{file}" from email titled "{subject}".'.format(file=fileName, subject=subject))
                
for response_part in data:
        if isinstance(response_part, tuple):
            msg = email.message_from_string(response_part[1].decode('utf-8'))
            email_subject = msg['subject']
            email_from = msg['from']
            print ('From : ' + email_from + '\n')
            print ('Subject : ' + email_subject + '\n')
            print(msg.get_payload(decode=True))


"""for num in data[0].split():
    typ, data = mail.fetch(num, '(RFC822)' )
    raw_email = data[0][1]
# converts byte literal to string removing b''
    raw_email_string = raw_email.decode('utf-8')
    email_message = email.message_from_string(raw_email_string)
# downloading attachments
    for part in email_message.walk():
        # this part comes from the snipped I don't understand yet... 
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()
        if bool(fileName):
            filePath = os.path.join(path_files, fileName)
            
            fp = open(filePath, 'wb')
            fp.write(part.get_payload(decode=True))
            fp.close()
            subject = str(email_message).split("Subject: ", 1)[1].split("\nTo:", 1)[0]
            print('Downloaded "{file}" from email titled "{subject}".'.format(file=fileName, subject=subject))
                
for response_part in data:
        if isinstance(response_part, tuple):
            msg = email.message_from_string(response_part[1].decode('utf-8'))
            email_subject = msg['subject']
            email_from = msg['from']
            print ('From : ' + email_from + '\n')
            print ('Subject : ' + email_subject + '\n')
            print(msg.get_payload(decode=True))
"""