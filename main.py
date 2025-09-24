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
from datetime import timedelta

def autenticar():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", config.SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print("Não foi possível atualizar o token. Pode ser necessário re-autenticar devido à mudança de escopo.")
                print("Por favor, delete o arquivo 'token.json' e rode o script novamente.")
                sys.exit()
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
        if not values or len(values) < 2:
            print("Nenhum dado encontrado na planilha.")
            return None
        
        # *** INÍCIO DA CORREÇÃO ***
        header = values[0]
        data = values[1:]
        
        # Pega o número de colunas do cabeçalho (deve ser 14)
        num_cols = len(header)
        
        # Garante que todas as linhas de dados tenham o mesmo número de colunas que o cabeçalho
        # Adiciona células vazias ('') se uma linha for mais curta
        padded_data = []
        for row in data:
            while len(row) < num_cols:
                row.append('')
            padded_data.append(row)
            
        # Cria o DataFrame com os dados já corrigidos
        df = pd.DataFrame(padded_data, columns=header)
        df.rename(columns={
            'Carimbo de data/hora': 'timestamp',
            'Informe a loja da ruptura': 'loja',
            'Tratativa Comercial': 'tratativa',
            'Informe o código do produto em ruptura': 'codigo_produto',
            'Informe o produto em ruptura': 'produto',
            'Informe a categoria da ruptura': 'categoria',
            'Informe seu nome': 'nome_solicitante',
            'A quanto tempo esse produto está em ruptura?': 'tempo_ruptura',
        }, inplace=True)
        # *** FIM DA CORREÇÃO ***
        
        # Garante que a coluna 'Status Relatorio' exista
        if 'Status Relatorio' not in df.columns:
            df['Status Relatorio'] = ""

        # Guarda o índice original para usar na hora de atualizar a planilha
        df['original_index'] = df.index + 2 # +2 porque a planilha começa em 1 e tem cabeçalho
        
        print(f"Leitura de {len(df)} registros concluída com sucesso!")
        return df
    except HttpError as err:
        print(f"Ocorreu um erro na API do Sheets: {err}")
        return None
    
def marcar_como_enviado(creds, df_processado):
    print("\n--- ATUALIZANDO STATUS NA PLANILHA GOOGLE ---")
    if df_processado.empty:
        print("Nenhum registro para atualizar.")
        return

    try:
        # --- INÍCIO DA NOVA LÓGICA ---
        # Cria uma cópia para trabalhar e evitar avisos
        df_a_marcar = df_processado.copy()

        # Limpa a coluna 'tratativa' para identificar corretamente as não tratadas
        df_a_marcar['tratativa'] = df_a_marcar['tratativa'].replace('', pd.NA)
        df_a_marcar['tratativa'] = df_a_marcar['tratativa'].fillna("Sem Tratativa")

        # Agora, filtramos para manter APENAS as solicitações que foram tratadas
        df_tratadas = df_a_marcar[df_a_marcar['tratativa'] != "Sem Tratativa"]
        # --- FIM DA NOVA LÓGICA ---

        if df_tratadas.empty:
            print("Nenhuma solicitação tratada para marcar na planilha. As pendências continuarão aparecendo.")
            return

        print(f"Encontradas {len(df_tratadas)} solicitações tratadas para marcar como 'Enviado'.")
        
        service = build("sheets", "v4", credentials=creds)
        data_atualizacao = datetime.now().strftime('%Y-%m-%d %H:%M')
        status_texto = f"Enviado em {data_atualizacao}"
        
        # Prepara os dados para a atualização em lote, usando apenas as linhas tratadas
        data = []
        for index in df_tratadas['original_index']:
            data.append({
                'range': f'RUPTURAS LOJAS!M{index}', # Coluna M = Status Relatorio
                'values': [[status_texto]]
            })
            
        body = {
            'valueInputOption': 'USER_ENTERED',
            'data': data
        }
        
        result = service.spreadsheets().values().batchUpdate(
            spreadsheetId=config.SPREADSHEET_ID, body=body).execute()
        
        print(f"{result.get('totalUpdatedCells')} células atualizadas com sucesso na planilha.")
        
    except HttpError as err:
        print(f"Ocorreu um erro ao atualizar a planilha: {err}")

def main():
    print("="*60)
    if config.MODO_TESTE:
        print(f"ATENÇÃO: SCRIPT EM MODO DE TESTE. E-mails serão enviados para: {config.EMAIL_TESTE}")
    else:
        print("ATENÇÃO: SCRIPT EM MODO DE PRODUÇÃO.")
    print("="*60)
    
    # *** NOVO MENU DE ESCOLHA DE PERÍODO ***
    print("Selecione o período para análise das solicitações:")
    print("1 - Últimos 7 dias")
    print("2 - Definir um intervalo de datas personalizado")
    
    data_inicio, data_fim = None, None
    hoje = datetime.now()

    while True:
        escolha = input("Digite sua escolha (1 ou 2): ").strip()
        if escolha == '1':
            data_fim = hoje
            data_inicio = hoje - timedelta(days=7)
            break
        elif escolha == '2':
            try:
                str_inicio = input("Digite a data de início (DD/MM/AAAA): ")
                data_inicio = datetime.strptime(str_inicio, '%d/%m/%Y')
                str_fim = input("Digite a data de fim (DD/MM/AAAA): ")
                data_fim = datetime.strptime(str_fim, '%d/%m/%Y').replace(hour=23, minute=59, second=59) # Inclui o dia todo
                break
            except ValueError:
                print("Formato de data inválido. Por favor, use DD/MM/AAAA.")
        else:
            print("Escolha inválida. Tente novamente.")
            
    print(f"\nProcessando dados de {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}")

    confirmacao = input("Deseja continuar? (s/n): ").lower().strip()
    if confirmacao not in ['s', 'sim']:
        print("Operação cancelada pelo usuário.")
        sys.exit()
    
    data_hoje_str = datetime.now().strftime('%Y-%m-%d')
    pasta_principal = "Relatorios_Enviados"
    pasta_de_hoje = os.path.join(pasta_principal, data_hoje_str)
    pasta_gerentes = os.path.join(pasta_de_hoje, "gerentes")
    pasta_compradores = os.path.join(pasta_de_hoje, "compradores")
    
    os.makedirs(pasta_gerentes, exist_ok=True)
    os.makedirs(pasta_compradores, exist_ok=True)
    
    print("\nConfirmado. Iniciando processo...")
    creds = autenticar()
    df_full = ler_dados_planilha(creds)

    if df_full is not None:
        # *** NOVA LÓGICA DE FILTRAGEM ***
        # 1. Filtra apenas as linhas que ainda não foram enviadas
        df_para_processar = df_full[df_full['Status Relatorio'] == ''].copy()

        # 2. Converte a coluna de data e filtra pelo período escolhido
        df_para_processar['timestamp'] = pd.to_datetime(df_para_processar['timestamp'], dayfirst=True)
        df_periodo = df_para_processar[(df_para_processar['timestamp'] >= data_inicio) & (df_para_processar['timestamp'] <= data_fim)].copy()


        if df_periodo.empty:
            print("\nNenhuma nova solicitação encontrada para o período selecionado.")
        else:
            print(f"\n{len(df_periodo)} novas solicitações encontradas para processar.")
            
            # Geração dos relatórios com o DataFrame já filtrado
            report_generator.gerar_relatorios_gerentes(creds, df_periodo, pasta_gerentes, enviar_email)
            report_generator.gerar_relatorios_compradores(creds, df_periodo, pasta_compradores, enviar_email)
            caminho_pdf_gerencial = report_generator.gerar_relatorio_gerencial_pdf(df_periodo, pasta_de_hoje, data_inicio, data_fim)
            

            if caminho_pdf_gerencial:
                print("\n--- ENVIANDO RELATÓRIO GERENCIAL POR E-MAIL ---")
                assunto = f"Relatório Gerencial de Rupturas - {data_inicio.strftime('%d/%m')} a {data_fim.strftime('%d/%m')}"
                corpo_html = f"""
                <html>
                    <body>
                        <p>Prezado(a),</p>
                        <p>Segue em anexo o relatório gerencial consolidado de rupturas, referente ao período de {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}.</p>
                        <br>
                        <p>Atenciosamente,<br>Equipe Comercial</p>
                    </body>
                </html>
                """
                
                destinatarios = [config.EMAIL_TESTE] if config.MODO_TESTE else config.GERENCIAL_EMAILS
                
                for destinatario in destinatarios:
                    enviar_email(creds, destinatario, assunto, corpo_html, caminho_pdf_gerencial)

            # Marca as linhas processadas como enviadas na planilha
            marcar_como_enviado(creds, df_periodo)
            
        print("\nProcesso concluído!")

if __name__ == "__main__":
    main()