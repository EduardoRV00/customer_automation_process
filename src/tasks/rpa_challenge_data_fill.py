from botcity.web import WebBot, Browser, By
from src.utils.setup_logs import *
from src.services.api_client import *
from botcity.maestro import *
from config import *
from src.tasks.processed_data import *
import pandas as pd
import logging
from src.utils.error_handling import *


def open_rpa_challenge_website(bot, logger_client, logger_dev):
    '''
    Open Rpa Challenge WebSite
    '''

    bot.headless = False
    bot.browse(URL_RPA_CHALLENGE)
    msg_opening_website = "Abre site do Rpa Challenge."
    logger_client.info(msg_opening_website)
    logger_dev.info(msg_opening_website)
        

def fill_rpa_challenge(bot, output_sheet, logger_client, logger_dev):
    try:
    
        entry_df = pd.read_excel(output_sheet)
        
        msg_consulting_cnpj = "Processando CNPJs e consultando API" 
        logger_client.info(msg_consulting_cnpj)
        logger_dev.info(msg_consulting_cnpj)
        wb = load_workbook(output_sheet)
        ws = wb.active
        
        '''
        First Name: Razão Social;

        Last Name: Situação Cadastral;

        E-mail: E-mail;

        Endereço: Endereço;

        Role In Company: Descrição Matriz/Filial;

        Company Name: Nome Fantasia;

        Phone Number: Telefone 
        '''
        
        
        
        for index, row in entry_df.iterrows():
            cnpj = str(row["CNPJ"])
            info = api_get(cnpj, logger_client, logger_dev)
            
            if info is None:
                break
            
            first_name = info["razao_social"]
            last_name = info["descricao_situacao_cadastral"]
            email = info["email"]
            address = info["Endereco"]
            role_in_company = info["identificador_matriz_filial"]
            company_name = info["nome_fantasia"]
            phone_number = info["ddd_telefone_1"]
        
            bot.find_element(selector='[ng-reflect-name="labelFirstName"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelFirstName"]', by=By.CSS_SELECTOR).send_keys(first_name)
            msg_razao_social = "Preenchimento de primeiro nome, com razão social."
            logger_client.info(msg_razao_social)
            logger_dev.info(msg_razao_social)

            bot.find_element(selector='[ng-reflect-name="labelLastName"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelLastName"]', by=By.CSS_SELECTOR).send_keys(last_name)
            msg_situacao_cadastral = "Preenchimento de ultimo nome com situação cadastral."
            logger_client.info(msg_situacao_cadastral)
            logger_dev.info(msg_situacao_cadastral)

            bot.find_element(selector='[ng-reflect-name="labelCompanyName"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelCompanyName"]', by=By.CSS_SELECTOR).send_keys(company_name)
            msg_nome_fantasia = "Preenchimento de nome da empresa com nome fantasia."
            logger_client.info(msg_nome_fantasia)
            logger_dev.info(msg_nome_fantasia)

            bot.find_element(selector='[ng-reflect-name="labelRole"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelRole"]', by=By.CSS_SELECTOR).send_keys(role_in_company)
            msg_identify = "Preenchimento de cargo na empresa com identificador matriz/filial"
            logger_client.info(msg_identify)
            logger_dev.info(msg_identify)

            bot.find_element(selector='[ng-reflect-name="labelAddress"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelAddress"]', by=By.CSS_SELECTOR).send_keys(address)
            msg_address = "Preenchimento de endereço"
            logger_client.info(msg_address)
            logger_dev.info(msg_address)

            bot.find_element(selector='[ng-reflect-name="labelEmail"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelEmail"]', by=By.CSS_SELECTOR).send_keys(email)
            msg_email = "Preenchimento de e-mail"
            logger_client.info(msg_email)
            logger_dev.info(msg_email)

            bot.find_element(selector='[ng-reflect-name="labelPhone"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelPhone"]', by=By.CSS_SELECTOR).send_keys(phone_number)
            msg_cell = "Preenchimento de telefone celular"
            logger_client.info(msg_cell)
            logger_dev.info(msg_cell)


            submit_btn = bot.find_element(selector='//input[@value="Submit"]', by=By.XPATH)
            submit_btn.click()
            msg_submit = "Click de Submit"
            logger_client.info(msg_submit)
            logger_dev.info(msg_submit)
    
    except Exception as e:
        # Log the error and handle it
        logger_dev.error(f"Error no preenchimento de dados no RPA Challenge: {e}")
        handle_error("Preenchimento de dados(RPA Challenge)", "Error no preenchimento de dados no RPA Challenges", logger_client, logger_dev)
    
     
        
    



