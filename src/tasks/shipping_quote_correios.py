from botcity.web import WebBot
import logging
from config import *

def open_correios_site():
  bot = WebBot()
  bot.headless = False
  bot.browse(URL_CORREIOS)
  bot.wait(3)
  logging.info("Abre site dos correios no navegador.")
