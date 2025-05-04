from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os, glob
import pandas as pd
import logging

logger = logging.getLogger(__name__)

logging.basicConfig(
    format='%(asctime)s %(levelname)-8s %(message)s',
    filename='descarga_reportes.log', encoding='utf-8',
    level=logging.INFO,
    datefmt='%Y-%m-%d %H:%M:%S')


# Getting All Files List
fileList = glob.glob('D:\\ETL\\Procesos\\FerreteriaSanJuan\\Reporte*.xlsx', recursive=True)
     
# Remove all files one by one
for file in fileList:
    try:
        os.remove(file)
    except OSError:
        logger.error("Error while deleting file")
  
logger.info("Removed all matched files!")

# Create Firefox options instance
options = webdriver.FirefoxOptions()

#Opcion in background
#options.add_argument('--headless')

# Set Firefox preferences for the download directory
options.set_preference("browser.download.folderList", 2)  # Use custom download path
options.set_preference("browser.download.dir", "D:\\ETL\\Procesos\\FerreteriaSanJuan")  # Set your custom download directory
options.set_preference("browser.download.useDownloadDir", True)
options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")  # MIME types to download without asking

# Set up the WebDriver (e.g., Chrome)
driver = webdriver.Firefox(options=options)  # Make sure chromedriver is installed and in PATH

# Open the login page
driver.get("https://app.miwally.com/")

# Allow the page to load completely
time.sleep(3)  # Adjust the sleep time as needed

# Locate the username/email field and enter the value
email_field = driver.find_element(By.ID, "wbof-textfield-username")  # Adjust the 'name' attribute if necessary
email_field.send_keys("juand.ratto@gmail.com")  # Replace with your email

# Locate the password field and enter the value
password_field = driver.find_element(By.ID, "wbof-textfield-password")  # Adjust the 'name' attribute if necessary
password_field.send_keys("R4tt0Wally$2023")  # Replace with your password

# Optionally, submit the form (either by pressing Enter or clicking the login button)
password_field.send_keys(Keys.RETURN)  # This submits the form by pressing Enter

# Alternatively, you can click the "Login" button:
login_button = driver.find_element(By.ID, "wbof-button-signin")
login_button.click()

# Allow time for the login to process
time.sleep(7)  # Adjust this sleep time based on how long the login takes

# Descargar el reporte de ventas

try:
    # Abrir la pagina de reporte de ventas
    driver.get("https://app.miwally.com/Report/Detail")
    menu_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "wbof-option-menu"))  # Use your own XPath or CSS selector
    )
    logger.info("El boton menú es clickable!")

    try:
        menu_button.click()
        logger.info("El boton menú fue clickeado!")
    except Exception as e:
        logger.error(e, stack_info=True, exc_info=True)

    try:
        day7_button = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div/div/div/main/div/div[2]/div/div[2]/div[1]/div[2]/div/div/div[2]/div/div/div/button[3]"))
        )
        logger.info("El boton descarga 7 dias es clickeable!")
    except Exception as e:
        logger.error(e, stack_info=True, exc_info=True)

    try:
        day7_button.click()
        logger.info("El boton descarga 7 dias fue clickeado!")
    except Exception as e:
        logger.error(e, stack_info=True, exc_info=True)

    #Boton download
    try:
        download_btn = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div/div/div[1]/main/div/div[2]/div/div[2]/div[1]/div[1]/div[3]/button")))
        logger.info("El boton descarga reporte de ventas es clickeable!")
    except Exception as e:
        logger.error(e, stack_info=True, exc_info=True)

    try:
        download_btn.click()
        logger.info("El boton descarga reporte de ventas fue clickeado!")
    except Exception as e:
        logger.error(e, stack_info=True, exc_info=True)
    
    # Download Ventas por Producto
    try:
        vtaXProd_btn = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div/div/div[3]/div/div/div[2]/form/div/div[1]/div[2]/div/div[3]/button")))
        logger.info("El boton descarga reporte ventas por producto es clickeable!")
    except Exception as e:
        logger.error(e, stack_info=True, exc_info=True)
    
    try:
        vtaXProd_btn.click()
        logger.info("Reporte ventas descargado.")
    except Exception as e:
        logger.error(e, stack_info=True, exc_info=True)    

except Exception as e:
    logger.info("Error: Tiemout! el boton menu no se volvio clickeable...")
    logger.error(e, stack_info=True, exc_info=True)
    driver.quit()
    exit() # Se detiene la ejecución


## Descargar el reporte de inventario
#directly navigate to the 'Report Inventory' page after login
driver.get("https://app.miwally.com/Inventory")
time.sleep(5)

# Clic en el 'menu'
menu_button = driver.find_element(By.ID, "wbof-option-menu")
menu_button.click()

time.sleep(3)

# Download the report
download_btn = driver.find_element(By.XPATH, "/html/body/div[5]/div/div/div/main/div/div[2]/div/div/div[2]/div/div/div[2]/button")
download_btn.click()

time.sleep(10)
logger.info("Reporte Inventario descargado.")

## Descargar el reporte de productos
# directly navigate to the 'Report Inventory' page after login
driver.get("https://app.miwally.com/Product")
time.sleep(10)

# Clic en el 'menu'
menu_button = driver.find_element(By.ID, "wbof-option-menu")
menu_button.click()

time.sleep(3)

# Download the report
download_btn = driver.find_element(By.XPATH, "//*[@id='app']/div/main/div/div[2]/div/div/div[1]/div[2]/div/button")
download_btn.click()

time.sleep(10)
logger.info("Reporte Producto descargado.")

# Close the browser (optional)
driver.quit()

#Convertir los xlsx a csv delimitados por tab.
# Define the path where your Excel files are located
input_path = 'D:\\ETL\\Procesos\\FerreteriaSanJuan'  # Replace with your directory path
output_path = 'D:\\ETL\\Procesos\\FerreteriaSanJuan'  # Replace with your output directory path

# Ensure the output directory exists
os.makedirs(output_path, exist_ok=True)

# Iterate over all Excel files in the directory
for file in glob.glob(os.path.join(input_path, 'Reporte*.xlsx')):
    # Read the Excel file
    df = pd.read_excel(file)
    
    # Get the base filename without the extension
    base_name = os.path.basename(file).replace('.xlsx', '.txt')
    
    # Define the output file path
    output_file = os.path.join(output_path, base_name)
    
    # Save the DataFrame as a tab-delimited text file
    df.to_csv(output_file, sep='\t', index=False,  encoding='latin1')
    
    print(f"Converted: {file} -> {output_file}")

logger.info("All Excel files have been processed!")

