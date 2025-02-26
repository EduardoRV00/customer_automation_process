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
from src.tasks.shipping_quote_jadlog import *
from src.tasks.shipping_quote_correios import *
from src.utils.manipulate_spreadsheet import *
from openpyxl import load_workbook
BotMaestroSDK.RAISE_NOT_CONNECTED = False


def main():
    setup_logging()
    logging.info("Início da execução do processo.")
    logging.info("Processando dados...")
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    #bot = WebBot()

    # Configure whether or not to run on headless mode
    #bot.headless = False

    # Uncomment to change the default Browser to Firefox
    # bot.browser = Browser.FIREFOX

    # Uncomment to set the WebDriver path
    bot.driver_path = CHROME_DRIVER

    # Opens the BotCity website.
    # open_correios_site()
    # get_screenshots()

    # Implement here your logic...
    
    
    # Creates the output sheet and assigns the file path to the variable output_sheet
    output_sheet = create_output_sheet()

    process_spreadsheet(output_sheet)

    # ABRE SITE CORREIOS
    open_correios_site(bot)
    # PREENCHE FORMULARIO
    logging.info("Inicia preenchimento dos dados de cotação dos correios.")
    process_shipping_quote_correios(bot, data)
    logging.info("Finaliza preenchimento de cotação dos correios.")
    # process_shipping_quotes(bot, output_sheet)
    logging.info("Fecha site dos correios no navegador.")
    bot.stop_browser()

    
    # Check the output sheet information | Is currently running with placeholders
    validar_informacoes(quote_data)
    
    # Performs quote on the jadlog website
    jadlog_quote(output_sheet)
    bot.stop_browser()

    # Wait 3 seconds before closing
    logging.info('Finalizando execução do bot...')
    bot.wait(3000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    # bot.stop_browser()


    print(data_fill_processed())
    
    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK.",
    #     total_items=0,
    #     processed_items=0,
    #     failed_items=0
    # )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
