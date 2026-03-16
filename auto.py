from __future__ import print_function
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import base64

# Escopos necessários: envio de e-mail + leitura do Sheets
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/spreadsheets.readonly"
]

# ID da planilha e aba
SPREADSHEET_ID = "1tPeuNYoJdw5PzYG9-4QJrpgc7Xb2km7VFU2sF-M-fB8"
RANGE = "Inscritos PS 2026.1!B2:C"  # Coluna B = Nome do Candidato, C = Email

# Autenticação
def authenticate():
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
    gmail = build("gmail", "v1", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)
    return gmail, sheets

# Criar e-mail com imagem
def create_email(to, subject, body_html, img_path):
    msg = MIMEMultipart()
    msg["To"] = to
    msg["Subject"] = subject
    msg.attach(MIMEText(body_html, "html", "utf-8"))
    with open(img_path, "rb") as img_file:
        img = MIMEImage(img_file.read())
        img.add_header("Content-ID", "<banner>")
        msg.attach(img)
    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {"raw": raw_message}

# Enviar e-mail
def send_email(service, to, subject, body_html, img_path):
    email_message = create_email(to, subject, body_html, img_path)
    service.users().messages().send(userId="me", body=email_message).execute()
    print(f"E-mail enviado para {to}")

# Main
gmail_service, sheets_service = authenticate()

# Lendo dados do Google Sheets
result = sheets_service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID,
    range=RANGE
).execute()

rows = result.get("values", [])

if not rows:
    print("Nenhum dado encontrado na planilha.")
else:
    img_path = "banner.png"
    for row in rows:
        if len(row) < 2:
            continue  # pula linhas incompletas
        nome = row[0]
        email = row[1]

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
                border: 3px solid rgba(0, 0, 0, 0.5);
                border-radius: 10px;
            }}
            .banner {{ width: 100%; height: auto; padding-bottom: 20px; }}
            a {{ color: #007bff; text-decoration: none; }}
          </style>
        </head>
        <body>
          <div class="container">
            <div class="content">
              <img src="cid:banner" alt="Banner" class="banner">
              <p><b>Olá {nome}!</b></p>
              <p>Recebemos sua inscrição no nosso Processo Seletivo. Estamos muito felizes com o seu interesse em fazer parte da Empresa Júnior PUC-Rio! <b>Sua inscrição foi confirmada com sucesso.</b></p>
              <p>Nosso principal canal de comunicação será o <b>e-mail</b>, então fique de olho na sua caixa de entrada (e no spam, só por precaução). Em breve, entraremos em contato com mais informações sobre as próximas etapas.</p>
              <p>Se tiver dúvidas, envie um e-mail para <a href="mailto:processoseletivo@empresajunior.com.br">processoseletivo@empresajunior.com.br</a> ou nos chame no Instagram @empresajunior.</p>
              <p><b>Boa sorte!</b></p>
            </div>
          </div>
        </body>
        </html>
        """

        send_email(gmail_service, email, "Confirmação de Inscrição", corpo_email, img_path)

    print("Todos os e-mails foram enviados com sucesso!")