import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import csv

# Configurações do servidor de e-mail
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL = 'seu_email@gmail.com'
PASSWORD = 'sua_senha_de_app'

# Função para ler contatos de um arquivo CSV
def ler_contatos(arquivo_csv):
    contatos = []
    with open(arquivo_csv, mode='r') as file:
        reader = csv.reader(file)
        next(reader)  # Ignora o cabeçalho do CSV
        for row in reader:
            nome, email = row
            contatos.append({'nome': nome, 'email': email})
    return contatos

# Função para enviar o e-mail
def enviar_email(destinatario, assunto, mensagem):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL
        msg['To'] = destinatario
        msg['Subject'] = assunto

        msg.attach(MIMEText(mensagem, 'plain'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL, PASSWORD)
            server.sendmail(EMAIL, destinatario, msg.as_string())
            print(f"E-mail enviado para {destinatario}")
    except Exception as e:
        print(f"Erro ao enviar e-mail para {destinatario}: {e}")

# Função principal para envio de e-mails em massa
def enviar_emails_em_massa(arquivo_csv, assunto, mensagem_padrao):
    contatos = ler_contatos(arquivo_csv)
    for contato in contatos:
        mensagem_personalizada = mensagem_padrao.format(nome=contato['nome'])
        enviar_email(contato['email'], assunto, mensagem_personalizada)

# Executando o envio de e-mails em massa
assunto = "Assunto do E-mail"
mensagem = "Olá, {nome}! Esta é uma mensagem personalizada."

arquivo_contatos = 'contatos.csv'  # Defina o caminho do arquivo CSV
enviar_emails_em_massa(arquivo_contatos, assunto, mensagem)
