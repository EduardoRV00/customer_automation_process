from botcity.web import WebBot, By
import re
import logging
from config import *
from openpyxl import load_workbook
from datetime import datetime, timedelta

data = [
    {
        "origin_zip_code": '38182-428',
        "destination_zip_code": '74013-020',
        "type_service": 'PAC',
        "dimension_height": 36,
        "dimension_width": 28,
        "dimension_length": 28,
        "product_weight": 8
    },
    {
        "origin_zip_code": '38182-428',
        "destination_zip_code": '80610-240',
        "type_service": 'SEDEX',
        "dimension_height": 36,
        "dimension_width": 28,
        "dimension_length": 28,
        "product_weight": 17
    },
    {
        "origin_zip_code": '38182-428',
        "destination_zip_code": '74020-170',
        "type_service": 'PAC',
        "dimension_height": 36,
        "dimension_width": 28,
        "dimension_length": 28,
        "product_weight": 10
    }
]


def open_correios_site(bot):
  """
  Abre site dos correios no navegador.
  """
  bot.headless = False
  bot.browse(URL_CORREIOS)
  # bot.wait(3000)
  logging.info("Abre site dos correios no navegador.")


def handle_delivery_time(delivery_text):
  """
  Extrai dias uteis (texto) e calcula prazo
  """
  delivery_working_days = re.search(r'(\d+)\s+dias\s+úteis', delivery_text)
  if not delivery_working_days:
     return "Prazo inválido"
  
  working_days = int(delivery_working_days.group(1))
  current_date = datetime.today()

  while working_days > 0:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5:  #segunda - sexta (dias uteis)
            working_days -= 1

  delivery_date = current_date.strftime("%d/%m/%Y")

  return delivery_date


def fill_correios_form(bot, data):
  try:
    # preenche campos de cep (oridem e estino)
    bot.find_element(selector="cepOrigem", by=By.NAME).send_keys(data['origin_zip_code'])
    bot.find_element(selector="cepDestino", by=By.NAME).send_keys(data['destination_zip_code'])

    #seleciona campo de serviço ("PAC ou "SEDEX")
    bot.find_element(selector="servico", by=By.NAME).click()
    if data["type_service"] == "PAC":
      bot.find_element(selector=f"option[value='04510']", by=By.CSS_SELECTOR).click()
    if data["type_service"] == "SEDEX":
      bot.find_element(selector=f"option[value='04014']", by=By.CSS_SELECTOR).click()

    #seleciona tipo de embalagem
    bot.find_element(selector="option[value='outraEmbalagem1']", by=By.CSS_SELECTOR).click()

    #preenche campos dimensões
    bot.find_element(selector="Altura", by=By.NAME).send_keys(str(data['dimension_height']))
    bot.find_element(selector="Largura", by=By.NAME).send_keys(str(data['dimension_width']))
    bot.find_element(selector="Comprimento", by=By.NAME).send_keys(str(data['dimension_length']))

    #seleciona peso
    bot.find_element(selector="peso", by=By.NAME).click()
    bot.find_element(selector=f"option[value='{data['product_weight']}']", by=By.CSS_SELECTOR).click()

    # logging.info("Clica em calcular frete.")
    bot.find_element(selector="Calcular", by=By.NAME).click()

    bot.wait(2000)

    open_tabs = bot.get_tabs()

    #ativa ultima aba aberta
    if len(open_tabs) > 1:
      bot.activate_tab(open_tabs[-1])

    #captura valor cotação
    value_element = bot.find_element(selector="//th[contains(text(),'Valor total')]/following-sibling::td", by=By.XPATH)
    value = value_element.text.strip() if value_element else "Valor não encontrado"
    print(f"Valor da cotação: {value}")


    #captura prazo entrega
    delivery_element = bot.find_element(selector="//th[contains(text(),'Prazo de entrega')]/following-sibling::td", by=By.XPATH)
    delivery_text = delivery_element.text.strip() if delivery_element else "Prazo não encontrado."
    delivery_date = handle_delivery_time(delivery_text)
    print(f"Prazo entrega: {delivery_date}")

    #fecha a segunda aba
    if len(open_tabs) > 1:
       bot.close_page()

    bot.activate_tab(open_tabs[0])
    bot.refresh()

    bot.wait(2000)
  except Exception as e:
    logging.error(f"Erro no preenchimento do formulário de cotação dos correios: {e}")


def process_shipping_quote_correios(bot, data_list):
   for data in data_list:
      fill_correios_form(bot, data)
