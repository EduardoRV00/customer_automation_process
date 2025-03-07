"""
WARNING:

Please make sure you install the bot dependencies with `pip install --upgrade -r requirements.txt`
in order to get all the dependencies on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the dependencies.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at
https://documentation.botcity.dev/tutorials/python-automations/web/
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
        # Runner passes the server url, the id of the task being executed,
        # the access token and the parameters that this task receives (when applicable).
        maestro = BotMaestroSDK.from_sys_args()
        ## Fetch the BotExecution with details from the task, including parameters
        execution = maestro.get_execution()

        print(f"Task ID is: {execution.task_id}")
        print(f"Task Parameters are: {execution.parameters}")

        bot = WebBot()

        # Configure whether or not to run on headless mode
        bot.headless = False

        # Uncomment to set the WebDriver path
        bot.driver_path = CHROME_DRIVER

        # Opens the BotCity website.
        # open_correios_site()
        # get_screenshots()

        # Implement here your logic...
    
        # Creates the output sheet and assigns the file path to the variable output_sheet
        output_sheet = create_output_sheet()

        process_spreadsheet(output_sheet, logger_client, logger_dev)
        
        # Fill output_sheet with API consultation data
        data_fill_processed(output_sheet)
        
        # Fills in empty output sheet cells
        fill_data_b4_rpachallenge(output_sheet)

        #Fill Rpa challenge text boxes.
        open_rpa_challenge_website(bot)
        get_screenshots(logger_client, logger_dev)
        
        fill_rpa_challenge(bot, output_sheet)

        #performs quote on the correios website
        open_correios_site(bot, logger_client, logger_dev)
        logger_client.info("Inicia busca de cotação dos Correios.")
        get_screenshots(logger_client, logger_dev)
        processed_output_sheet_quote_correios(bot, output_sheet, logger_client, logger_dev)
        logger_client.info("Finaliza busca de cotação dos Correios.")
        get_screenshots(logger_client, logger_dev)
        logger_client.info("Fecha site dos correios no navegador.")
        bot.stop_browser()

        
        # Check the output sheet information | Is currently running with placeholders
        # validar_informacoes(quote_data)
        
        # Performs quote on the jadlog website
        open_jadlog_site(bot)
        get_screenshots(logger_client, logger_dev)
        # Performs quote on the jadlog website
        jadlog_quote(output_sheet, bot)
        get_screenshots(logger_client, logger_dev)
        
        # Fills in empty output sheet cells after quotes
        fill_missing_values(output_sheet)

        logger_client.info('Finalizando execução do bot...')
        logger_client.info("Processo Finalizado.")
        # bot.wait(3000)

        # Finish and clean up the Web Browser
        # You MUST invoke the stop_browser to avoid
        # leaving instances of the webdriver open
        bot.stop_browser()
        subject = email_success_subject()
        body = email_success_body()
        send_email(output_sheet, subject, body, logger_client, logger_dev)

        # Uncomment to mark this task as finished on BotMaestro
        # maestro.finish_task(
        #     task_id=execution.task_id,
        #     status=AutomationTaskFinishStatus.SUCCESS,
        #     message="Task Finished OK.",
        #     total_items=0,
        #     processed_items=0,
        #     failed_items=0
        # )
        
    
    except Exception as e:
        logger_dev.error(f"Erro fatal durante a execução do bot: {e}")
        logging.info("Finalizando a execução do bot") 


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
