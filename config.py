import os

#Path to excel files
EXCEL_RAW_PATH = r'C:\customer_automation_process\src\utils\data\processed\Customer_Registration_and_New_Orders_Quote\Planilha_de_Entrada_Grupo3.xlsx'

EXCEL_PROCESSED_PATH = ''

#Path to chromeDriver.exe
CHROME_DRIVER = r'C:\customer_automation_process\src\utils\driver\chromedriver.exe'

#Path to log
# LOG_PATH = 'C:\customer_automation_process\log'

LOG_DIR = os.path.join(os.getcwd(), "Logs")

URL_CORREIOS = "https://www2.correios.com.br/sistemas/precosPrazos/"
URL_JADLOG = "https://www.jadlog.com.br/jadlog/simulacao"
