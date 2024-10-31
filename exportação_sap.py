import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import csv
import win32com.client

# Configurações do servidor de e-mail
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL = 'seu_email@gmail.com'
PASSWORD = 'sua_senha_de_app'

# Função para conectar ao SAP e baixar o arquivo
def baixar_arquivo_sap(transacao , caminho_arquivo):
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        if not sap_gui:
            raise Exception("SAP GUI não encontrado.")
        
        application = sap_gui.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        # Acessa a transação desejada no SAP
        session.StartTransaction(Transaction=transacao)

        # Adicione os comandos necessários para realizar a extração no SAP
        # Exemplo: preencher filtros e executar o relatório para salvar o arquivo
        session.findById("...").Text = "..."  # Ajuste de acordo com os campos do SAP
        session.findById("...").press()

        # Salva o relatório ou arquivo no caminho especificado
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select()  # Exemplo de menu para exportação
        session.findById("...").Text = caminho_arquivo  # Defina o caminho de salvamento

        print(f"Arquivo salvo em: {caminho_arquivo}")

    except Exception as e:
        print(f"Erro ao extrair arquivo do SAP: {e}")

# Função para enviar e-mails com anexo
def enviar_email(destinatario, assunto, mensagem, caminho_arquivo):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL
        msg['To'] = destinatario
        msg['Subject'] = assunto

        msg.attach(MIMEText(mensagem, 'plain'))

        # Anexar o arquivo baixado do SAP
        with open(caminho_arquivo, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {caminho_arquivo}",
        )
        msg.attach(part)

        # Conectar ao servidor e enviar o e-mail
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL, PASSWORD)
            server.sendmail(EMAIL, destinatario, msg.as_string())
            print(f"E-mail enviado para {destinatario}")

    except Exception as e:
        print(f"Erro ao enviar e-mail para {destinatario}: {e}")

# Função principal para envio de e-mails em massa
def enviar_emails_em_massa(arquivo_csv, assunto, mensagem_padrao, caminho_arquivo_sap):
    # Baixar o arquivo do SAP
    baixar_arquivo_sap("ZRELATORIO", caminho_arquivo_sap)  # Defina a transação SAP e o caminho do arquivo

    # Ler contatos e enviar e-mails com o arquivo em anexo
    with open(arquivo_csv, mode='r') as file:
        reader = csv.reader(file)
        next(reader)  # Ignora o cabeçalho do CSV
        for row in reader:
            nome, email = row
            mensagem_personalizada = mensagem_padrao.format(nome=nome)
            enviar_email(email, assunto, mensagem_personalizada, caminho_arquivo_sap)

# Configurações do envio de e-mail
assunto = "Relatório Automático do SAP"
mensagem = "Olá, {nome}! Aqui está o relatório mais recente."
arquivo_contatos = 'contatos.csv'
caminho_arquivo_sap = 'C:\\caminho\\para\\relatorio.xlsx'  # Defina o caminho de salvamento do arquivo

# Executar o envio de e-mails em massa com o anexo do SAP
enviar_emails_em_massa(arquivo_contatos, assunto, mensagem, caminho_arquivo_sap)
