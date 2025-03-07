from datetime import datetime
import os
import pandas as pd
from openpyxl import load_workbook
import re
import logging
import sys

import pyautogui

from config import ERROR_DIR
from tasks.send_email import send_email
from src.utils.email_templates import *

def validate_cep(output_sheet):
    
    try:
        # xlsx file reading
        df = pd.read_excel(output_sheet)
        count = 0

        # Loop to travel dataframe cells and fill empty cells
        for index, row in df.iterrows():
            if not re.match(r'^\d{8}$', str(row['CEP'])):
                logging.warning(f"CEP inválido ou ausente na linha {row.name + 2}")
                count += 1

        logging.info(f'Quantidade de CEPs inválidos ou não encontrados: {count}')
        
        if count >= 3:
            logging.warning('Finalizando execução do bot - Faltam dados para realizar cotações')
            sys.exit() # Finish the bot execution
            
            # Uncomment to mark this task as finished on BotMaestro
            # maestro.finish_task(
            #     task_id=execution.task_id,
            #     status=AutomationTaskFinishStatus.SUCCESS,
            #     message="Task Finished OK.",
            #     total_items=0,
            #     processed_items=0,
            #     failed_items=0
            # )
        
    except Exception as e:
        logging.warning(f"Erro ao validar CEPs na planilha de saída: {e}")   


def handle_error(process_name, task_name, logger_client, logger_dev):
 
    # Handle unexpected errors by capturing a screenshot, logging the error, and notifying via email.
    try:
        # Create error directory with timestamp
        timestamp = datetime.now().strftime('%d%m%Y-%H%M')
        error_folder = os.path.join(ERROR_DIR, f"Errors_{timestamp}")
        os.makedirs(error_folder, exist_ok=True)
        
        # Capture screenshot
        screenshot_filename = os.path.join(error_folder, f"Erro - RPA {task_name} - {timestamp}.png")
        screenshot = pyautogui.screenshot()
        screenshot.save(screenshot_filename)
        
        # Log error message
        log_msg = f"Erro inesperado na tarefa {task_name} do processo {process_name}. Screenshot salvo em {screenshot_filename}."
        logger_client.info(log_msg)
        logger_dev.error(log_msg)
        
        # Prepare email notification
        subject = email_error_subject()
        body = email_error_body()
        
        # Send email notification with screenshot attached
        send_email(screenshot_filename, subject, body)
        
    except Exception as e:
        logger_dev.error(f"Erro ao processar a contingência: {e}")