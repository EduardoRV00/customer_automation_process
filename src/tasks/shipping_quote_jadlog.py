from botcity.web import WebBot
from botcity.web import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging
import sys
import os

# Add the root folder to sys.path
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from config import *
from src.utils.setup_logs import *

# Placeholder data to execute jadlog_quote
cep_origem = 14092465
cep_destino = 14060510
modalidade = 'JadLog Econômico'
peso = 5
valor_mercadoria = 500
valor_coleta = 100
entrega = 'Retira C.O.'
largura = 50
altura = 30
comprimento = 40

def jadlog_quote():
    
    # Uncomment to test log
    setup_logging()
    
    # Open the jadlog quote simulation website
    bot = WebBot()
    bot.driver_path = CHROME_DRIVER
    bot.headless = False
    logging.info("Abre site dos correios no navegador.")
    bot.browse(URL_JADLOG)
    
    # Fill the form with xlsx data
    logging.info('Preencher dados na simulação Jadlog')
    field_origem = bot.find_element("//input[@id='origem']", By.XPATH)
    field_origem.send_keys(cep_origem)
    
    field_destino = bot.find_element("//input[@id='destino']", By.XPATH)
    field_destino.send_keys(cep_destino)
    
    modalidade_dropdown = bot.find_element("//select[@id='modalidade']", By.XPATH)
    select = Select(modalidade_dropdown)
    select.select_by_visible_text(modalidade)  # Substitua pelo texto visível da opção desejada
    
    field_peso = bot.find_element("//input[@id='peso']", By.XPATH)
    field_peso.clear()
    field_peso.send_keys(peso)
    
    field_valor_merc = bot.find_element("//input[@id='valor_mercadoria']", By.XPATH)
    field_valor_merc.clear()
    field_valor_merc.send_keys(valor_mercadoria)
    
    field_valor_coleta = bot.find_element("//input[@id='valor_coleta']", By.XPATH)
    field_valor_coleta.clear()
    field_valor_coleta.send_keys(valor_coleta)
    
    entrega_dropdown = bot.find_element("//select[@id='entrega']", By.XPATH)
    select = Select(entrega_dropdown)
    select.select_by_visible_text(entrega)
    
    field_largura = bot.find_element("//input[@id='valLargura']", By.XPATH)
    field_largura.clear()
    field_largura.send_keys(largura)
    
    field_altura = bot.find_element("//input[@id='valAltura']", By.XPATH)
    field_altura.clear()
    field_altura.send_keys(altura)
    
    field_comprimento = bot.find_element("//input[@id='valComprimento']", By.XPATH)
    field_comprimento.clear()
    field_comprimento.send_keys(comprimento)
    
    # Click the 'Simular' button
    logging.info('Realizar click em "Simular"')
    simular_btn = bot.find_element("//input[@value='Simular']", By.XPATH)
    simular_btn.click()

    # Get the quote value
    resultado = WebDriverWait(bot.driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="j_idt45_content"]/span'))
    )
    resultado_text = resultado.get_attribute('innerText')
    logging.info(f'Valor da simulação: {resultado_text}')

    bot.wait(3000)
    # return bot
       
# jadlog_quote()