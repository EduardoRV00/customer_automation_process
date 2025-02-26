import pandas as pd
from openpyxl import load_workbook
from src.services.api_client import *
from config import EXCEL_RAW_PATH, EXCEL_PROCESSED_PATH
from src.utils.setup_logs import *
from src.utils.manipulate_spreadsheet import save_status_to_output

def data_fill_processed(output_sheet):
    '''
    Fill data processed Excel file with API json.

    '''
    
    try:
        logging.info("Lendo arquivo Excel de entrada")
        entry_df = pd.read_excel(output_sheet)
        
        logging.info("Processando CNPJs e consultando API")
        wb = load_workbook(output_sheet)
        ws = wb.active
        
        # Mapping columns
        column_mapping = {
            "RAZÃO SOCIAL": 2,
            "NOME FANTASIA": 3,
            "ENDEREÇO": 4,
            "CEP": 5,
            "DESCRIÇÃO MATRIZ FILIAL": 6,
            "TELEFONE + DDD": 7,
            "E-MAIL": 8
        }
        
        for index, row in entry_df.iterrows():
            cnpj = str(row["CNPJ"])
            info = api_get(cnpj)
            
            if info is None:
                info = {
                    "razao_social": "Não disponível",
                    "nome_fantasia": "Não disponível",
                    "descricao_situacao_cadastral": "Não disponível",
                    "Endereco": "Não informado, S/N - Não informado",
                    "cep": "Não informado",
                    "identificador_matriz_filial": "Não informado",
                    "ddd_telefone_1": "Não informado",
                    "email": "N/A"
                }
                
            
            # Verify if the company it's active.
            status = "" if info["descricao_situacao_cadastral"] == "ATIVA" else "Empresa inativa"
            
            # Fill data to sheet
            ws.cell(row=index + 2, column=column_mapping["RAZÃO SOCIAL"]).value = info["razao_social"]
            ws.cell(row=index + 2, column=column_mapping["NOME FANTASIA"]).value = info["nome_fantasia"]
            ws.cell(row=index + 2, column=column_mapping["ENDEREÇO"]).value = info["Endereco"]
            ws.cell(row=index + 2, column=column_mapping["CEP"]).value = info["cep"]
            ws.cell(row=index + 2, column=column_mapping["DESCRIÇÃO MATRIZ FILIAL"]).value = info["identificador_matriz_filial"]
            ws.cell(row=index + 2, column=column_mapping["TELEFONE + DDD"]).value = info["ddd_telefone_1"]
            ws.cell(row=index + 2, column=column_mapping["E-MAIL"]).value = info["email"]
            
            
            if status:
                save_status_to_output(output_sheet, index, status)
            
            logging.info(f"Dados do CNPJ {cnpj} preenchidos com sucesso")
        
        wb.save(output_sheet)
        logging.info("Planilha preenchida com sucesso!")

    except Exception as e:
        logging.error(f"Erro ao preencher a planilha: {e}")
        