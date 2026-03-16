from __future__ import print_function
import os
import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from google.oauth2.credentials import Credentials
import base64

# Escopo do Gmail
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

# Função para autenticação no Gmail
def gmail_authenticate():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)

# Função para criar o e-mail com a imagem anexa
def create_email_with_attachment(to, subject, body_html, img_path):
    msg = MIMEMultipart()
    msg["To"] = to
    msg["Subject"] = subject
    msg.attach(MIMEText(body_html, "html", "utf-8"))
    
    # Anexando a imagem
    with open(img_path, "rb") as img_file:
        img = MIMEImage(img_file.read())
        img.add_header("Content-ID", "<banner>")
        msg.attach(img)

    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {"raw": raw_message}

# Função para enviar o e-mail
def send_email(service, to, subject, body_html, img_path):
    email_message = create_email_with_attachment(to, subject, body_html, img_path)
    service.users().messages().send(userId="me", body=email_message).execute()
    print(f"E-mail enviado para {to}")

# Autenticando na API do Gmail
service = gmail_authenticate()

# Lendo a planilha Excel
df = pd.read_excel("envioAuto emailPs.xlsx")

# Caminho da imagem
img_path = "banner.png"

# Loop para enviar os e-mails
for _, row in df.iterrows():
    nome = row["NOME"]
    email = row["EMAIL"]

    # Corpo do e-mail
    corpo_email = f"""\
    <!DOCTYPE html>
    <html lang="pt">
    <head>
      <meta charset="UTF-8">
      <title>Confirmação de Inscrição</title>
      <style>
        body {{ font-family: Arial, sans-serif; background-color: #f6f6f6; padding: 20px; }}
        .container {{ max-width: 600px; background: white; padding: 20px; border-radius: 10px; }}
        .content {{
            padding: 10px;
            border: 3px solid rgba(0, 0, 0, 0.5);  /* Borda preta com 70% de opacidade */
            border-radius: 10px;  /* Borda arredondada */
        }}
        .banner {{
            width: 100%;
            height: auto;
            padding-bottom: 20px;
        }}
        a {{ color: #007bff; text-decoration: none; }}
      </style>
    </head>
    <body>
      <div class="container">
        <div class="content">
          <img src="cid:banner" alt="Banner" class="banner">
          <p><b>Olá {nome}!</b></p>
          <p>Recebemos sua inscrição no nosso Processo Seletivo. Estamos muito felizes com o seu interesse em fazer parte da Empresa Júnior PUC-Rio! <b>Sua inscrição foi confirmada com sucesso. </b></p>
          <p> Nosso principal canal de comunicação será o <b>e-mail</b>, então fique de olho na sua caixa de entrada (e no spam, só por precaução). Em breve, entraremos em contato com mais informações sobre as próximas etapas.</p>
          <p>Se tiver dúvidas, envie um e-mail para <a href="mailto:processoseletivo@empresajunior.com.br">processoseletivo@empresajunior.com.br</a> ou nos chame no Instagram @empresajunior.</p>
          <p><b>Boa sorte!</b></p>
        </div>
      </div>
    </body>
    </html>
    """

    # Enviar e-mail com a imagem anexa
    send_email(service, email, "Confirmação de Inscrição", corpo_email, img_path)

print("Todos os e-mails foram enviados com sucesso!")