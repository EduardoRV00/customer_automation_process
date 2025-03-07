import logging
import os
import pyautogui
from datetime import datetime
from config import *


def create_daily_logs():
   """
   Creates daily directories to store logs.
   """
   try:
    current_date = datetime.now().strftime('%d%m%Y')

   #daily log directories
    daily_log_dir_cl = os.path.join(LOG_DIR, current_date)
    daily_log_dir_dev = os.path.join(LOG_DEV, current_date)

    os.makedirs(daily_log_dir_cl, exist_ok=True)
    os.makedirs(daily_log_dir_dev, exist_ok=True)

    return daily_log_dir_cl, daily_log_dir_dev
   except Exception as e:
      logging.error(f"Erro ao criar diretório diário de logs para cliente: {e}")
      raise


def generate_log_filenames():
   """
   Generates the log file names for client and development.
   """
   try:
    daily_log_dir_cl, daily_log_dir_dev = create_daily_logs()

    timestamp = datetime.now().strftime('%d%m%Y-%H%M%S')

    log_cli_filename = os.path.join(daily_log_dir_cl, f"{timestamp}.txt")
    log_dev_filename = os.path.join(daily_log_dir_dev, f"{timestamp}_dev.txt")

    return log_cli_filename, log_dev_filename
   except Exception as e:
      logging.error(f"Erro ao gerar os nomes dos arquivos de log: {e}")
      raise


def setup_logging():
  """
  Configures loggers for client and development.
  """
  try:
    log_cli_filename, log_dev_filename = generate_log_filenames()

    #Log -> client
    logger_client = logging.getLogger("client")
    logger_client.setLevel(logging.INFO)
    if not logger_client.handlers:
      file_handler_cl = logging.FileHandler(log_cli_filename, mode='a')
      file_handler_cl.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
      logger_client.addHandler(file_handler_cl)
      logger_client.propagate = False


    #Log -> dev
    logger_dev = logging.getLogger("dev")
    logger_dev.setLevel(logging.DEBUG)
    if not logger_dev.handlers:
      file_handler_dev = logging.FileHandler(log_dev_filename, mode='a')
      file_handler_dev.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
      logger_dev.addHandler(file_handler_dev)
      logger_dev.propagate = False

    return logger_client, logger_dev
    
  except Exception as e:
     logging.error(f"Erro ao configurar os loggers: {e}")


def get_screenshots(logger_client, logger_dev):
   """
   Captures screenshot and saves it in the client and developer log directories.
   """
   try:
      daily_log_dir_cl, daily_log_dir_dev = create_daily_logs()

      #capture the screen
      screenshot = pyautogui.screenshot()
      timestamp = datetime.now().strftime('%d%m%Y-%H%M%S')

      screenshot_filename_client = os.path.join(daily_log_dir_cl, f"{timestamp}-RPA_Cadastrar_cliente_no sistema_challenge.jpg")
      screenshot_filename_dev = os.path.join(daily_log_dir_dev, f"{timestamp}_dev-RPA_Cadastrar_cliente_no_sistema_challenge.jpg")

      screenshot.save(screenshot_filename_client)
      screenshot.save(screenshot_filename_dev)
      
      logger_client.info(f"Imagem capturada e salva no log do cliente.")
      logger_dev.info(f"Imagem capturada e salva no log do desenvolvedor.")

   except Exception as e:
    logger_dev.error(f"Erro ao capturar os screenshots: {e}")
