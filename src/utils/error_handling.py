import pandas as pd
from openpyxl import load_workbook
import re
import logging
import sys

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