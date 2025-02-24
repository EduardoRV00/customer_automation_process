from botcity.web import WebBot, By
import logging
from config import *

def open_correios_site():
  bot = WebBot()
  bot.headless = False
  bot.browse(URL_CORREIOS)
  bot.wait(3)
  logging.info("Abre site dos correios no navegador.")

# DADOS PROVISORIOS PARA TESTE
origin_zip_code = '38182-428'
destination_zip_code = '74013-020'
type_service = 'PAC'
dimension_height = 36
dimension_width = 28
dimension_length = 28
product_weight = 8

def fill_correios_form(bot):
  try:
    logging.info("Preenchendo formulário no site dos correios...")

    bot.find_element(selector="cepOrigem", by=By.NAME).send_keys(origin_zip_code)
    bot.find_element(selector="cepDestino", by=By.NAME).send_keys(destination_zip_code)
    bot.find_element(selector="servico", by=By.NAME).click()
    bot.find_element(selector="option[value='04510']", by=By.CSS_SELECTOR).click()
    bot.find_element(selector="option[value='outraEmbalagem1']", by=By.CSS_SELECTOR).click()
    bot.find_element(selector="Altura", by=By.NAME).send_keys(str(dimension_height))
    bot.find_element(selector="Largura", by=By.NAME).send_keys(str(dimension_width))
    bot.find_element(selector="Comprimento", by=By.NAME).send_keys(str(dimension_length))
    bot.find_element(selector="peso", by=By.NAME).click()
    bot.find_element(selector="option[value='8']", by=By.CSS_SELECTOR).click()
    bot.find_element(selector="Calcular", by=By.NAME).click()

    logging.info("Formulário de cotação preenchido.")

  except Exception as e:
    logging.error(f"Erro no preenchimento do formulário de cotação dos correios: {e}")
