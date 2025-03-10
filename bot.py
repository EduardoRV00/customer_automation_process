"""
WARNING:
Please make sure you install the bot dependencies with `pip install --upgrade -r requirements.txt`
in order to get all the dependencies on your Python environment.

"""

import logging
from botcity.web import WebBot, Browser, By
from botcity.maestro import *
from src.utils.setup_logs import *
from src.tasks.processed_data import *
from src.tasks.fill_api_data_to_processed import *
from src.tasks.shipping_quote_jadlog import *
from src.tasks.shipping_quote_correios import *
from src.utils.manipulate_spreadsheet import *
from src.tasks.fill_api_data_to_processed import *
from src.utils.error_handling import *
from openpyxl import load_workbook
from src.tasks.rpa_challenge_data_fill import *
from src.tasks.send_email import send_email
from src.utils.email_templates import *
BotMaestroSDK.RAISE_NOT_CONNECTED = False


def main():
    
    try:
        logger_client, logger_dev = setup_logging()
        logger_client.info("Início da execução do processo.")
        logger_dev.info("Processo iniciado pelo bot.")

        maestro = BotMaestroSDK.from_sys_args()
        execution = maestro.get_execution()

        print(f"Task ID is: {execution.task_id}")
        print(f"Task Parameters are: {execution.parameters}")

        bot = WebBot()

        # Configure whether or not to run on headless mode
        bot.headless = False

        # Webdriver setting path
        bot.driver_path = CHROME_DRIVER
    
        # Creates the output sheet and assigns the file path to the variable output_sheet
        output_sheet = create_output_sheet()

        process_spreadsheet(output_sheet, logger_client, logger_dev)
        
        # Fill output_sheet with API consultation data
        data_fill_processed(output_sheet, logger_client, logger_dev)
        
        fill_data_b4_rpachallenge(output_sheet)

        # Fill Rpa challenge text boxes.
        open_rpa_challenge_website(bot, logger_client, logger_dev)
        get_screenshots(logger_client, logger_dev)

        
        fill_rpa_challenge(bot, output_sheet, logger_client, logger_dev)

        # Performs quote on the correios website
        open_correios_site(bot, logger_client, logger_dev)
        msg_quote_correios = "Inicia busca de cotação dos Correios."
        logger_client.info(msg_quote_correios), logger_dev.info(msg_quote_correios)
        get_screenshots(logger_client, logger_dev)
        processed_output_sheet_quote_correios(bot, output_sheet, logger_client, logger_dev)
        msg_ends_quote_correios = "Finaliza busca de cotação dos Correios."
        logger_client.info(msg_ends_quote_correios), logger_dev.info(msg_ends_quote_correios)
        get_screenshots(logger_client, logger_dev)
        msg_close_correois_site = "Fecha site dos correios no navegador."
        logger_client.info(msg_close_correois_site), logger_dev.info(msg_close_correois_site)
        bot.stop_browser()
        
        
        # Opens jadlog site and performs quote
        open_jadlog_site(bot, logger_client, logger_dev)
        get_screenshots(logger_client, logger_dev)
        jadlog_quote(output_sheet, bot, logger_client, logger_dev)
        get_screenshots(logger_client, logger_dev)
        
        # Fills in empty output sheet cells after quotes
        fill_missing_values(output_sheet)

        msg_end_bot = 'Finalizando execução do bot... Processo Finalizado.'
        logger_client.info(msg_end_bot), logger_dev.info(msg_end_bot)
        # bot.wait(3000)

        # Leaving instances of the webdriver open
        bot.stop_browser()
        subject = email_success_subject()
        body = email_success_body()
        send_email(output_sheet, subject, body, logger_client, logger_dev)

        # maestro.finish_task(
        #     task_id=execution.task_id,
        #     status=AutomationTaskFinishStatus.SUCCESS,
        #     message="Task Finished OK.",
        #     total_items=0,
        #     processed_items=0,
        #     failed_items=0
        # )
        
    
    except Exception as e:
        # Log the error and call the error handling function
        logger_dev.error(f"Erro fatal durante a execução do bot: {e}")
        handle_error("Execução da automação completa", "Erro durante a execução do bot(Execução geral do bot)", logger_client, logger_dev)


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
