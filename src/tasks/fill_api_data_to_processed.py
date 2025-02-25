import pandas as pd
from openpyxl import load_workbook
from src.services.api_client import *
from config import EXCEL_RAW_PATH, EXCEL_PROCESSED_PATH
from src.utils.setup_logs import *
from src.tasks.processed_data import *

def data_fill_processed(output_sheet):
    try:
        logging.info("Lendo arquivo Excel de entrada")
        entry_df = pd.read_excel(EXCEL_RAW_PATH)
        results = []
        
        logging.info("Processando CNPJs e consultando API")
        for _, row in entry_df.iterrows():
            cnpj = str(row["CNPJ"])
            info = api_get(cnpj)

            if info is None:
                info = {
                    "cnpj": cnpj,
                    "razao_social": "Não disponível",
                    "nome_fantasia": "Não disponível",
                    "descricao_situacao_cadastral": "Não disponível",
                    "Endereco": "Não informado, S/N - Não informado",
                    "cep": "Não informado",
                    "identificador_matriz_filial": "Não informado",
                    "ddd_telefone_1": "Não informado",
                    "email": "N/A"
                }

            info["Status"] = "" if info["descricao_situacao_cadastral"] == "ATIVA" else "Empresa inativa"
            results.append(info)

        if not results:
            raise Exception("Erro: results está vazio.")

        logging.info("Convertendo results para DataFrame")
        df_results = pd.DataFrame(results)
        df_results = df_results.rename(columns={
            "cnpj": "CNPJ",
            "razao_social": "RAZÃO SOCIAL",
            "nome_fantasia": "NOME FANTASIA",
            "Endereco": "ENDEREÇO",
            "cep": "CEP",
            "identificador_matriz_filial": "DESCRIÇÃO MATRIZ FILIAL",
            "ddd_telefone_1": "TELEFONE + DDD",
            "email": "E-MAIL"
        })
        additional_columns = [
            "VALOR DO PEDIDO", "DIMENSÕES CAIXA", "PESO DO PRODUTO",
            "TIPO DE SERVIÇO JADLOG", "TIPO DE SERVIÇO CORREIOS",
            "VALOR COTAÇÃO JADLOG", "VALOR COTAÇÃO CORREIOS",
            "PRAZO DE ENTREGA CORREIOS",
            "Status"
        ]
        
        logging.info("Adicionando colunas adicionais se necessário")
        for col in additional_columns:
            if col not in df_results.columns:
                df_results[col] = ""

        logging.info("Abrindo planilha de saída")
        wb = load_workbook(output_sheet)
        ws = wb.active

        # Check if header it's available.
        if ws.max_column == 0:
            raise Exception("Erro: A planilha de saída não tem cabeçalhos definidos.")

        excel_headers = [ws.cell(row=1, column=col).value for col in range(1, ws.max_column + 1)]
         
        
        next_row = ws.max_row + 1
        logging.info("Preenchendo os respectivos dados no arquivo Excel")

        for index, df_row in df_results.iterrows():
            for col_index, header in enumerate(excel_headers, start=1):
                if header in df_results.columns:
                    value = df_row[header]
                else:
                    value = ""

                ws.cell(row=next_row + index, column=col_index, value=value)

        wb.save(output_sheet)
        logging.info("Planilha preenchida com sucesso!")

    except Exception as e:
        logging.error(f"Erro ao preencher a planilha: {e}")
