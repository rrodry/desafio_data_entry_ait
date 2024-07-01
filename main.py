from selenium import webdriver
from datetime import datetime
from openpyxl import Workbook, load_workbook
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.firefox.options import Options
from Functions.functions import file_size, file_change_name, modify_excel, uploadFiles
from openpyxl.utils import get_column_letter
import requests

import time
import os

#Data
user = "desafiodataentry"
password = "desafiodataentrypass"
overlayDownload = True
columnData = ["CODIGO", "DESCR", "MARCA", "PRECIO"]
date = datetime.now().strftime("%d-%m-%Y")  # Date today

#Firefox config

# Select directoy folder name
current_directory = os.path.join(os.getcwd(), 'downloads')

# Create if directory is not found
if not (os.path.exists(current_directory)):
    os.makedirs(current_directory)

options = Options()
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.manager.showWhenStarting", False)
options.set_preference("browser.download.dir", current_directory)
options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-gzip")

#Driver
driver = webdriver.Firefox(options=options)
urlOriginal = "https://desafiodataentryfront.vercel.app/"

#Direction URL
driver.get(urlOriginal)
original_window = driver.current_window_handle
buttons = driver.find_elements(By.TAG_NAME, "button")
buttonsId = []
files = os.listdir(current_directory)  # Get directory

#select all buttons
for i in range(0, len(buttons)):
    buttonsId.append(buttons[i].get_attribute("id"))

suppliers = [supplier.text for supplier in driver.find_elements(By.TAG_NAME, "h3")]

for index in range(0, len(buttonsId)):
    driver.find_element(By.ID, buttonsId[index]).click()
    if urlOriginal != driver.current_url:
        if index <= 1:
            files = os.listdir(current_directory)  # Get directory
            driver.find_element(By.NAME, "username").send_keys(user)
            driver.find_element(By.NAME, "password").send_keys(password)
            driver.find_element(By.TAG_NAME, "button").click()
            time.sleep(2)
            # Select Checkbox and click
            divCh = driver.find_elements(By.CSS_SELECTOR, "#brands-checkboxes .flex")  # find div all supplier
            for div in divCh:
                div.find_element(By.TAG_NAME, "input").click()  # browse all suppliers and select

        time.sleep(2)
        driver.find_element(By.ID, "download-button").click()  # once all are select download all

    #Here i saw that after press the button Download, overlay show up and when go the downdload is ready.
    overlay = True  # flag overlay
    timeout = 120  # maximum waiting time in seconds
    start_time = time.time()

    # While 120s not response or download is ready
    filesQuantity = len(files)
    downloading = True

    while time.time() - start_time < timeout and overlay:
        files = os.listdir(current_directory)

        try:
            # Check if overlay is present
            overlayIn = driver.find_element(By.ID, 'loading-overlay')
        except NoSuchElementException:
            print("La descarga ha finalizado.")
            downloading = False
            overlay = False
            file_change_name(files, current_directory)
            break
        time.sleep(0.5)

    driver.get(urlOriginal)
    time.sleep(2)

driver.quit()
time.sleep(2)

try:
    modify_excel()
except NoSuchElementException:
    print("Erro on function Modify")
try:
    uploadFiles()
except NoSuchElementException:
    print("Erro on function uploadFiles")

