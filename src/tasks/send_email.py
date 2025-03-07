import smtplib
import pandas as pd
import os
from email.message import EmailMessage
from dotenv import load_dotenv

# Carregar variáveis do arquivo .env
load_dotenv()

# Configurações para Outlook (Gmail)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587  # Alterado para TLS
# EMAIL_USER = "eduardo.rochavargas2@gmail.com" #os.getenv("EMAIL_USER")
# EMAIL_PASS = "aplb ohnj ncpg epjb" #os.getenv("EMAIL_PASS")
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

# Lista de e-mails definidos manualmente
destinatarios = [
    "ghislaine.latorra.pb@compasso.com.br",
    "guilherme.sartori.pb@compasso.com.br",
    "marcos.eduardo.pb@compasso.com.br",
    # "luana.costa@compasso.com.br",
    # "rafael.vizzotto@compasso.com.br",
    # "aline.santos@compasso.com.br",
    "eduardo.rochavargas@outlook.com",
    "eduardo.vargas.pb@compasso.com.br"
]

# Criar e enviar o e-mail para vários destinatários
def send_email(output_sheet):
    if not destinatarios:
        print("Nenhum destinatário definido.")
        return
    
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()  # Habilita criptografia TLS
        server.login(EMAIL_USER, EMAIL_PASS)
        print("✅ Conexão e autenticação SMTP bem-sucedidas!")

        # Caminho atualizado da planilha de saída
        # PLANILHA_SAIDA = "C:\\develop\\customer_automation_process\\src\\utils\\data\\processed\\output_sheet_05-03-2025_10-42-33.xlsx"
        
        for destinatario in destinatarios:
            msg = EmailMessage()
            msg["From"] = EMAIL_USER
            msg["To"] = destinatario
            msg["Subject"] = f"RPA Cadastro Cliente - {pd.Timestamp.now().strftime('%d%m%Y')} – {pd.Timestamp.now().strftime('%H%M')}"
            
            msg.set_content(
                f"""
                O processo RPA Cadastro Cliente foi executado com sucesso na data {pd.Timestamp.now().strftime('%d/%m/%Y')} – {pd.Timestamp.now().strftime('%H:%M')}.
                """
            )
            
            # Adicionar anexo
            with open(output_sheet, "rb") as f:
                msg.add_attachment(f.read(), maintype="application", subtype="vnd.ms-excel", filename="output_sheet.xlsx")
            
            server.send_message(msg)
            print(f"E-mail enviado para {destinatario} com sucesso!")
        
        server.quit()
    except Exception as e:
        print(f"Erro ao enviar e-mails: {e}")

if __name__ == "__main__":
    send_email()
