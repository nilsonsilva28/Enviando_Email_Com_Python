import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from segredos import senha 
from segredos import email
import pandas as pd

#pip install openpyxl
#pip install pandas 

clientes = pd.read_excel('./emails.xlsx')
#ITERAR SOBRE OS DADOS NA PLANILHA
for index, cliente in clientes.iterrows():
    # CRIAÇÃO DO EMAIL
    msg = MIMEMultipart()
    msg['Subject'] = 'Seu titulo'
    msg['From'] = email
    msg['To'] = cliente['Email']
    message = "Seu conteudo a ser enviado"
    msg.attach(MIMEText(message,'html'))

    #CAMINHO DO ARQUIVO DO TIPO ANEXO : 
    cam_aquivo = (r'C:\Users\susan\Desktop\Curriculo\OPERADOR DE PRODUCAO.pdf')
    attchment = open(cam_aquivo,'rb')

    att= MIMEBase('aplication','octet-stream')
    att.set_payload(attchment.read())
    encoders.encode_base64(att)

    att.add_header('content-Disposition',f'attchemt;filename=OPERADOR DE PRODUCAO.pdf')
    attchment.close()
    msg.attach(att)

    # CONFIGURAÇÃO DO SERVIDOR SMTP
    server  = smtplib.SMTP('smtp.gmail.com',port=587)
    server.starttls()
    server.login(email,senha)
    server.sendmail(msg['From'],msg['To'],msg.as_string())
    server.quit()