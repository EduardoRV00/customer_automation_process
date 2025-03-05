from botcity.web import WebBot, Browser, By
from src.utils.setup_logs import *
from src.services.api_client import *
from botcity.maestro import *
from config import *
from src.tasks.processed_data import *
import pandas as pd
import logging


def open_rpa_challenge_website(bot):
    '''
    Open Rpa Challenge WebSite
    '''

    bot.headless = False
    bot.browse(URL_RPA_CHALLENGE)
    logging.info("Abre site do Rpa Challenge.") 



def fill_rpa_challenge(bot, output_sheet):
    try:
    
        entry_df = pd.read_excel(output_sheet)
            
        logging.info("Processando CNPJs e consultando API")
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
            info = api_get(cnpj)
            
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
            logging.info("Preenchimento de primeiro nome, com razão social.")

            bot.find_element(selector='[ng-reflect-name="labelLastName"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelLastName"]', by=By.CSS_SELECTOR).send_keys(last_name)
            logging.info("Preenchimento de ultimo nome com situação cadastral.")

            bot.find_element(selector='[ng-reflect-name="labelCompanyName"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelCompanyName"]', by=By.CSS_SELECTOR).send_keys(company_name)
            logging.info("Preenchimento de nome da empresa com nome fantasia.")

            bot.find_element(selector='[ng-reflect-name="labelRole"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelRole"]', by=By.CSS_SELECTOR).send_keys(role_in_company)
            logging.info("Preenchimento de cargo na empresa com identificador matriz/filial")

            bot.find_element(selector='[ng-reflect-name="labelAddress"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelAddress"]', by=By.CSS_SELECTOR).send_keys(address)
            logging.info("Preenchimento de endereço")

            bot.find_element(selector='[ng-reflect-name="labelEmail"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelEmail"]', by=By.CSS_SELECTOR).send_keys(email)
            logging.info("Preenchimento de e-mail")

            bot.find_element(selector='[ng-reflect-name="labelPhone"]', by=By.CSS_SELECTOR).click()
            bot.find_element(selector='[ng-reflect-name="labelPhone"]', by=By.CSS_SELECTOR).send_keys(phone_number)
            logging.info("Preenchimento de telefone celular")


            submit_btn = bot.find_element(selector='//input[@value="Submit"]', by=By.XPATH)
            submit_btn.click()
            logging.info("Click de Submit") 
    
    except Exception as e:
        logging.error(f"Error no preenchimento de dados no RPA Challenge")
    
     
        
    



