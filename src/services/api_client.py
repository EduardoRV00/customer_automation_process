import requests 
from src.utils.setup_logs import *

def api_get(cnpj):
    # Consulta de API Brasil. 
    try:
        setup_logging()
        
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
            "logradouro": data.get("logradouro", "Não informado"), 
            "numero_endereco": data.get('numero', 'S/N'),
            "municipio":  data.get('municipio', 'Não informado'),
            "cep": data.get('cep', 'Não informado'),
            "situacao_cadastral": data.get("situacao", "Não disponível")
        }
        
        
        logging.info("Retorna o Dict com as respectivas informações")
        return  info
        
    except Exception as e:
        logging.error(f"Erro na consulta do CNPJ {cnpj}: {str(e)}")
        return None
