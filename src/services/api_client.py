import requests 

def api_get(cnpj, logger_client, logger_dev):
    '''
    
    Creates a function to consult API Brazil and retrieve company information based on CNPJ.
    Logs the request process and stores the retrieved data in a dictionary with key-value pairs.
    
    The function sends a request to Brasil API, extracts relevant company details,
    and returns them in a structured format. If an error occurs, it logs the issue and returns None.
    '''
    try:
        
        if len(cnpj) < 14:
            cnpj = cnpj.zfill(14)
            
        msg_request = "Realizando request em Brasil API "
        logger_client.info(msg_request)
        logger_dev.info(msg_request)
        api_url = f'https://brasilapi.com.br/api/cnpj/v1/{cnpj}'
        response = requests.get(api_url)
        response.raise_for_status()
        data = response.json()
        

        if response.status_code == 200: 
            # Dict to format requests in key:value
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
            
            
            return  info
            
    except ValueError as e:
        logger_dev.error(f"Erro ao processar resposta JSON para o CNPJ {cnpj}: {str(e)}")
    except requests.exceptions.RequestException as e:
        logger_dev.warning(f"Erro ao consultar o CNPJ {cnpj}: {str(e)}")
    return None
