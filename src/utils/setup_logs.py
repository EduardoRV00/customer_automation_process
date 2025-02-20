# setup_log.py
import logging
import os
import pyautogui
from datetime import datetime
from config import *


def create_daily_log():
   try:
    current_date = datetime.now().strftime('%d%m%Y')
    daily_log_dir = os.path.join(LOG_DIR, current_date)

    if not os.path.exists(daily_log_dir):
            os.makedirs(daily_log_dir)

    return daily_log_dir
   except Exception as e:
      logging.error(f"Erro ao criar diretório diário de logs: {e}")
      raise


def generate_log_filename():
   try:
    daily_log_dir = create_daily_log()
    log_filename = os.path.join(daily_log_dir, f"{datetime.now().strftime('%d%m%Y-%H%M%S')}")
    return log_filename
   except Exception as e:
      logging.error(f"Erro ao gerar o nome do arquivo de log: {e}")
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
      daily_log_dir = create_daily_log()
      screenshot = pyautogui.screenshot()
      timestamp = datetime.now().strftime('%d%m%Y-%H%M%S')
      screenshot_filename = os.path.join(daily_log_dir, f"{timestamp} - RPA_Cadastrar_cliente_no sistema_challenge.jpg")

      screenshot.save(screenshot_filename)
      logging.info(f"Screenshot capturada e salva em: {screenshot_filename}")

   except Exception as e:
    logging.error(f"Erro ao capturar os screenshots: {e}")
