import pandas as pd
import os
import mimetypes
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# ler a tabela Excel
df = pd.read_excel("nomes_emails.xlsx")
print(df.head())

for index, row in df.iterrows():
    nome = row['Nome']
    email = row['Email']
    filename = f"{nome}.pdf"

    if not os.path.exists(filename):
        print(f"O certificado de {nome} não foi encontrado.")
        continue

    # Defina as informações do servidor SMTP
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587

    # Defina suas credenciais de e-mail
    username = 'reynanvt@gmail.com'
    password = 'wkgkulqlzcvfesur'

    # Defina o destinatário do e-mail e o caminho do arquivo de anexo
    email_destino = email
    attachment = f'./{filename}'

    # Crie um objeto de conexão SMTP
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(username, password)

    # Leia o arquivo de anexo e adicione-o à mensagem do e-mail
    with open(attachment, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(attachment)
        file_type = mimetypes.guess_type(file_name)[0]

    attachment = MIMEBase(file_type, file_type)
    attachment.set_payload(file_data)
    attachment.add_header('Content-Disposition',
                          'attachment', filename=file_name)
    encoders.encode_base64(attachment)

    # Adicione o corpo do e-mail
    body = f'{nome}, segue em anexo o seu certificado.'
    text = MIMEText(body)
    message = MIMEMultipart()
    message.attach(text)

    # Adicione o anexo ao e-mail
    message['From'] = username
    message['To'] = email
    message['Subject'] = 'Assunto do email'
    message.attach(attachment)

    # Envie o e-mail
    text = message.as_string()
    server.sendmail(username, email_destino, text)
    print('email enviado com sucesso')

    # Encerre a conexão SMTP
    server.quit()
