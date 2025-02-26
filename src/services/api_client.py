import requests 
from src.utils.setup_logs import *

def api_get(cnpj):
    '''
    
    Creates a function to consult API Brazil and retrieve company information based on CNPJ.
    Logs the request process and stores the retrieved data in a dictionary with key-value pairs.
    
    The function sends a request to Brasil API, extracts relevant company details,
    and returns them in a structured format. If an error occurs, it logs the issue and returns None.
    '''
    try:
        setup_logging()
        
        if len(cnpj) < 14:
            cnpj = cnpj.zfill(14)
            
        
        logging.info("Realizando request em Brasil API ")
        api_url = f'https://brasilapi.com.br/api/cnpj/v1/{cnpj}'
        response = requests.get(api_url)  
        response.raise_for_status()
        data = response.json()
        
        
        # Dict para formatação dos requests em chave : valor
        logging.info("Armazenando informações em formato de Dict chave : valor")
        info = {
            "razao_social": data.get("razao_social", "Não disponível"),
            "nome_fantasia": data.get("nome_fantasia", "Não disponível"),
            "descricao_situacao_cadastral": data.get("descricao_situacao_cadastral", "Nao disponivel"),
            "Endereco": f"{data.get('logradouro', 'Não informado')}, {data.get('numero', 'S/N')} - {data.get('municipio', 'Não informado')}",
            "cep": data.get("cep", "Não informado"),
            "identificador_matriz_filial": "Matriz" if data.get("identificador_matriz_filial") == 1 else "Filial",
            "ddd_telefone_1": data.get("ddd_telefone_1", "Não informado"),
            "email":"N/A" if data.get("email") == None else "email"
        }
        
        
        logging.info("Retorna o Dict com as respectivas informações")
        return  info
        
    except Exception as e:
        logging.error(f"Erro na consulta do CNPJ {cnpj}: {str(e)}")
        return None
