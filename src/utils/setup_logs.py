import logging
import os
import pyautogui
from datetime import datetime
from config import *


def create_daily_log_client():
   try:
    current_date_cl = datetime.now().strftime('%d%m%Y')
    daily_log_dir_cl = os.path.join(LOG_DIR, current_date_cl)

    if not os.path.exists(daily_log_dir_cl):
            os.makedirs(daily_log_dir_cl)

    return daily_log_dir_cl
   except Exception as e:
      logging.error(f"Erro ao criar diretório diário de logs para cliente: {e}")
      raise

def create_daily_log_dev():
   try:
      current_date_dev = datetime.now().strftime('%d%m%Y')
      daily_log_dir_dev = os.path.join(LOG_DEV, current_date_dev)

      if not os.path.exists(daily_log_dir_dev):
         os.makedirs(daily_log_dir_dev)

      return daily_log_dir_dev
   except Exception as e:
      logging.error(f"Erro ao criar diretório para logs de desenvolvedor: {e}")


def generate_log_filename():
   try:
    daily_log_dir_cl = create_daily_log_client()
    daily_log_dir_dev = create_daily_log_dev()

    timestamp = datetime.now().strftime('%d%m%Y-%H%M%S')

    log_cli_filename = os.path.join(daily_log_dir_cl, f"{timestamp}.txt")
    log_dev_filename = os.path.join(daily_log_dir_dev, f"{timestamp}_dev.txt")

    return log_cli_filename, log_dev_filename
   except Exception as e:
      logging.error(f"Erro ao gerar os nomes dos arquivos de log: {e}")
      raise


def setup_logging():
  try:
    log_filename = generate_log_filename()
    logging.basicConfig(
            filename=log_filename,
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",  
        )
  except Exception as e:
     logging.error(f"Erro ao configurar o logging: {e}")
     raise


def get_screenshots():
   try:
      daily_log_dir_cl = create_daily_log_client()
      daily_log_dir_dev = create_daily_log_dev()

      screenshot = pyautogui.screenshot()
      timestamp = datetime.now().strftime('%d%m%Y-%H%M%S')

      screenshot_filename_client = os.path.join(daily_log_dir_cl, f"{timestamp} - RPA_Cadastrar_cliente_no sistema_challenge.jpg")
      screenshot_filename_dev = os.path.join(daily_log_dir_dev, f"{timestamp} - RPA_Cadastrar_cliente_no_sistema_challenge_dev.jpg")

      screenshot.save(screenshot_filename_client)
      screenshot.save(screenshot_filename_dev)

      logging.info(f"Screenshot capturada e salva em no log do cliente.")
      logging.info(f"Screenshot capturada e salva em no log do desenvolvedor.")

   except Exception as e:
    logging.error(f"Erro ao capturar os screenshots: {e}")
