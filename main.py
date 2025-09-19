# main.py

import os.path
import pandas as pd
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import sys
from datetime import datetime
import config
import report_generator
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def autenticar():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", config.SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", config.SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds

def enviar_email(creds, para, assunto, corpo_html, nome_arquivo_anexo=None):
    try:
        service = build("gmail", "v1", credentials=creds)
        message = MIMEMultipart()
        message["to"] = para
        message["from"] = config.MEU_EMAIL_REMETENTE
        message["subject"] = assunto
        message.attach(MIMEText(corpo_html, "html"))
        if nome_arquivo_anexo:
            with open(nome_arquivo_anexo, "rb") as attachment:
                part = MIMEApplication(attachment.read(), Name=os.path.basename(nome_arquivo_anexo))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(nome_arquivo_anexo)}"'
            message.attach(part)
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        create_message = {"raw": encoded_message}
        send_message = (service.users().messages().send(userId="me", body=create_message).execute())
        print(f"E-mail enviado com sucesso para {para}. ID: {send_message['id']}")
    except HttpError as error:
        print(f"Ocorreu um erro ao enviar o e-mail: {error}")

def ler_dados_planilha(creds):
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        result = (sheet.values().get(spreadsheetId=config.SPREADSHEET_ID, range=config.RANGE_NAME).execute())
        values = result.get("values", [])
        if not values:
            print("Nenhum dado encontrado.")
            return None
        df = pd.DataFrame(values[1:], columns=values[0])
        df.columns = ["timestamp", "email_solicitante", "nome_solicitante", "cargo", "loja", "categoria", "codigo_produto", "produto", "tempo_ruptura", "tratativa", "usuario_tratativa", "data_tratativa"]
        print("Leitura de dados concluída com sucesso!")
        return df
    except HttpError as err:
        print(f"Ocorreu um erro na API do Sheets: {err}")
        return None

def main():
    print("="*60)
    if config.MODO_TESTE:
        print(f"ATENÇÃO: SCRIPT EM MODO DE TESTE. E-mails serão enviados para: {config.EMAIL_TESTE}")
    else:
        print("ATENÇÃO: SCRIPT EM MODO DE PRODUÇÃO. E-mails serão enviados para os destinatários reais.")
    print("="*60)
    confirmacao = input("Deseja continuar? (s/n): ").lower().strip()
    if confirmacao not in ['s', 'sim']:
        print("Operação cancelada pelo usuário.")
        sys.exit()
    
    data_hoje = datetime.now().strftime('%Y-%m-%d')
    pasta_principal = "Relatorios_Enviados"
    pasta_de_hoje = os.path.join(pasta_principal, data_hoje)
    pasta_gerentes = os.path.join(pasta_de_hoje, "gerentes")
    pasta_compradores = os.path.join(pasta_de_hoje, "compradores")
    
    os.makedirs(pasta_gerentes, exist_ok=True)
    os.makedirs(pasta_compradores, exist_ok=True)
    
    print(f"\nRelatórios de gerentes serão salvos em: '{pasta_gerentes}'")
    print(f"Relatórios de compradores serão salvos em: '{pasta_compradores}'")
    
    print("\nConfirmado. Iniciando processo...")
    creds = autenticar()
    df = ler_dados_planilha(creds)

    if df is not None:
        report_generator.gerar_relatorios_gerentes(creds, df, pasta_gerentes, enviar_email) 
        
        # *** ESTA É A LINHA QUE FOI CORRIGIDA ***
        report_generator.gerar_relatorios_compradores(creds, df, pasta_compradores, enviar_email)
        
        print("\nProcesso concluído!")

if __name__ == "__main__":
    main()