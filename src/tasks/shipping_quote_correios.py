from botcity.web import WebBot, By
import re
import logging
from config import *
from openpyxl import load_workbook
from datetime import datetime, timedelta
from src.utils.manipulate_spreadsheet import *



def open_correios_site(bot):
  """
  Abre site dos correios no navegador.
  """
  bot.headless = False
  bot.browse(URL_CORREIOS)
  logging.info("Abre site dos correios no navegador.")


def validate_data(service, height, width, length, weight):
  try:
    fields = {
          "Tipo de serviço": service,
          "Altura": height,
          "Largura": width,
          "Comprimento": length,
          "Peso": weight
      }
    
    missing_fields = [name for name, value in fields.items() if not value]
    if missing_fields:
       return False, missing_fields
    
    return True, None
  except Exception as e:
     logging.error(f"Erro ao validar dados: {e}")


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

  return current_date.strftime("%d/%m/%Y")


def handle_select_weight(bot, weight):
    try:       
        if weight:
            #converter para inteiro
            if isinstance(weight, str):
                if weight.isdigit():
                    weight = int(weight)
                else:
                    return

            if not (1 <= weight <= 30):
                logging.error(f"Peso {weight} inválido. Deve estar entre 1 e 30.")
                return

        else:
            logging.warning("Peso não fornecido. Ignorando linha.")
            return


        #seleciona o peso no site
        weight_str = str(weight)
        weight_select = bot.find_element(selector="peso", by=By.NAME).click()
        option = bot.find_element(selector=f"option[value='{weight_str}']:not([disabled])", by=By.CSS_SELECTOR)

        if option:
            option.click()
        else:
            logging.error(f"Peso {weight_str} não disponível ou desabilitado para seleção.")

    except Exception as e:
        logging.error(f"Erro ao selecionar o peso: {e}")


def fill_correios_form(bot, service, height, width, length, weight, row, output_sheet):
  try:
    #preenchimento do comprimento a prtir de 13
    if length < 13:
       wb = load_workbook(output_sheet)
       ws = wb.active
       ws.cell(row=row, column=17, value="Erro ao realizar a cotação Correios")
       wb.save(output_sheet)
       return 

    valid, missing_fields = validate_data(service, height, width, length, weight)
    if not valid:
       logging.info(f"Encontrado campos ausentes. Buscar próxima cotação.")
       return
       
    origin_zip = 38182428
    dest_zip = 80610240

    # preenche campos de cep (oridem e estino)
    bot.find_element(selector="cepOrigem", by=By.NAME).send_keys(origin_zip)
    bot.find_element(selector="cepDestino", by=By.NAME).send_keys(dest_zip)

    #seleciona campo de serviço ("PAC ou "SEDEX")
    bot.find_element(selector="servico", by=By.NAME).click()
    if service == "PAC":
      bot.find_element(selector="option[value='04510']", by=By.CSS_SELECTOR).click()
    if service == "SEDEX":
      bot.find_element(selector="option[value='04014']", by=By.CSS_SELECTOR).click()

    #seleciona tipo de embalagem
    bot.find_element(selector="option[value='outraEmbalagem1']", by=By.CSS_SELECTOR).click()

    #preenche campos dimensões
    bot.find_element(selector="Altura", by=By.NAME).send_keys(height)
    bot.find_element(selector="Largura", by=By.NAME).send_keys(width)
    bot.find_element(selector="Comprimento", by=By.NAME).send_keys(length)

    #seleciona peso
    handle_select_weight(bot, weight)

    # logging.info("Clica em calcular frete.")
    bot.find_element(selector="Calcular", by=By.NAME).click()

    open_tabs = bot.get_tabs()

    if len(open_tabs) > 1:
      bot.activate_tab(open_tabs[-1])

    #captura valor cotação
    quote_value = bot.find_element(selector="//th[contains(text(),'Valor total')]/following-sibling::td", by=By.XPATH).text
    quote_value = float(quote_value.replace('R$', '').replace('.', '').replace(',', '.'))
    quote_value = round(quote_value, 2)

    #captura prazo entrega
    delivery = bot.find_element(selector="//th[contains(text(),'Prazo de entrega')]/following-sibling::td", by=By.XPATH).text
    delivery_date = handle_delivery_time(delivery)

    save_quote_and_delivery(output_sheet, row, quote_value, delivery_date)

    close_tabs(bot)

  except Exception as e:
    logging.error(f"Erro ao preencher cotação dos correios: {e}")


def save_quote_and_delivery(output_sheet, row, quote_value, delivery_date):
  """
   Salva o valor da cotação e o prazo de entrega na planilha de saída
  """
  try:
    if row is None or row < 1:
      logging.error(f"Valor inválido de linha: {row}")
      return
        
    if quote_value is None or delivery_date is None:
      logging.error("Dados de cotação ou prazo ausentes. Não será salvo na planilha.")
      return

    wb = load_workbook(output_sheet)
    ws = wb.active
    ws.cell(row=row, column=15, value=quote_value)
    ws.cell(row=row, column=16, value=delivery_date)
    wb.save(output_sheet)
    logging.info(f"Valor da cotação e prazo de entrega salvos na linha {row}")
  except Exception as e:
     logging.error(f"Erro ao salvar cotação e prazo de entrega: {e}")
   

def close_tabs(bot):
   """
   Fecha a segunda aba e volta para a primeira aba.
   """
   open_tabs = bot.get_tabs()
   if len(open_tabs) > 1:
      bot.activate_tab(open_tabs[-1])
      bot.close_page()
      bot.activate_tab(open_tabs[0])
      bot.refresh()


def read_output_sheet(bot, output_sheet):
  try:
    wb = load_workbook(output_sheet)
    ws = wb.active

    data = []

    for row_index, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        dimensions = row[9]
        weight = row[10]
        service = row[12]

        if dimensions and "x" in dimensions:
           height, width, length = map(int, dimensions.split(" x "))
        else:
           logging.warning(f"Dimensões inválidas ou ausentes: {dimensions}. Pulando linha.")
           continue

        fill_correios_form(bot, service, height, width, length, weight, row_index, output_sheet)
             
  except Exception as e:
      logging.error(f"Erro ao ler a planilha de saída: {e}")

