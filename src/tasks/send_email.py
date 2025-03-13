import smtplib
import pandas as pd
import os
from email.message import EmailMessage
from dotenv import load_dotenv

from src.utils.error_handling import *

# Load environment variables from .env file
load_dotenv()

# SMTP Configuration for Gmail
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587  # TLS Port
EMAIL_USER = os.getenv("EMAIL_USER")  # Gmail account
EMAIL_PASS = os.getenv("EMAIL_PASS")  # Gmail app password

# List of recipients
recipients = [
    "ghislaine.latorra.pb@compasso.com.br",
    # "guilherme.sartori.pb@compasso.com.br",
    # "marcos.eduardo.pb@compasso.com.br",
    # "luana.costa@compasso.com.br",
    # "rafael.vizzotto@compasso.com.br",
    # "aline.santos@compasso.com.br",
    # "eduardo.rochavargas@outlook.com",
    # "eduardo.vargas.pb@compasso.com.br"
]

# Function to send an email with an attachment
def send_email(output_file, subject, body, logger_client,logger_dev):
    
    # Sends an email to multiple recipients with the given subject, body, and an attachment.
    if not recipients:
        log_msg_no_recipients = "Nenhum destinatário definido."
        logger_client.info(log_msg_no_recipients)
        logger_dev.warning("Tentativa de envio de e-mail sem destinatários.")
        return
    
    try:
        # Establish connection with SMTP server
        log_msg_smtp_connect = "Conectando ao servidor SMTP."
        logger_client.info(log_msg_smtp_connect)
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()  # Enable TLS encryption
        server.login(EMAIL_USER, EMAIL_PASS)
        log_msg_smtp_auth = "Autenticação no servidor SMTP bem-sucedida."
        logger_client.info(log_msg_smtp_auth)
        
        for recipient in recipients:
            log_msg_preparing_email = f"Preparando e-mail para {recipient}."
            logger_client.info(log_msg_preparing_email)
            msg = EmailMessage()
            msg["From"] = EMAIL_USER
            msg["To"] = recipient
            msg["Subject"] = subject
            
            msg.set_content(body)
            
            # Attach file
            file_name = os.path.basename(output_file)
            with open(output_file, "rb") as f:
                msg.add_attachment(f.read(), maintype="application", subtype="vnd.ms-excel", filename=file_name)
            
            # Send email
            server.send_message(msg)
            log_msg_email_sent = f"E-mail enviado com sucesso para {recipient}."
            logger_client.info(log_msg_email_sent)
            logger_dev.info(log_msg_email_sent)
        
        # Close SMTP connection
        log_msg_smtp_close = "Conexão com o servidor SMTP encerrada."
        logger_client.info(log_msg_smtp_close)
        server.quit()
    except Exception as e:
        # Log error and call handle_error
        log_msg_error = "Erro ao enviar e-mails. Consulte os logs do desenvolvedor para mais detalhes."
        logger_client.info(log_msg_error)
        logger_dev.error(f"Erro ao enviar e-mails: {e}")
        handle_error("Envio de e-mail", "Envio de e-mail", logger_client, logger_dev)


if __name__ == "__main__":
    send_email()
