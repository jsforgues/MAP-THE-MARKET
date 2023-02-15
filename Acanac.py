from __future__ import print_function
import pytest
import time
import datetime
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support import expected_conditions as EC


import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account

# https://developers.google.com/sheets/api/quickstart/python

################################
SERVICE_ACCOUNT_FILE = 'potent-result-377317-9d73e13040c7.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# The ID spreadsheet.
SAMPLE_SPREADSHEET_ID = '1xYbOPRAPCiiD2ZOqwyGCeQWoIZVcSiID_A-pdLUkY7Q'


service = build('sheets', 'v4', credentials=creds)
# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range="testa!B2:B3843").execute()
values = result.get('values', [])

####################################


PATH = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)

driver.get("https://acanac.com/fr/internet-quebec//")

for row in values:

    driver.find_element(By.ID, "street-address").click()
    driver.find_element(By.ID, "street-address").send_keys(row[0][0])
    time.sleep(0.1)
    driver.find_element(By.ID, "street-address").send_keys(row[0][1])
    time.sleep(0.1)
    driver.find_element(By.ID, "street-address").send_keys(row[0][2])
    time.sleep(0.1)
    driver.find_element(By.ID, "street-address").send_keys(row[0][3])
    time.sleep(0.1)
    driver.find_element(By.ID, "street-address").send_keys(row[0][4])
    time.sleep(0.1)
    driver.find_element(By.ID, "street-address").send_keys(row[0][5])
    time.sleep(0.5)
    driver.find_element(By.ID, "street-address").send_keys(Keys.ENTER)
    time.sleep(0.5)
    driver.find_element(By.ID, "street-address").send_keys(Keys.ENTER)
    element = WebDriverWait(driver, 10000).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Effacer")))
    
    # Récupérer la date du jour
    today = datetime.datetime.today().strftime('%Y-%m-%d')

    # Écrire la date dans le Google Sheet
    range_name = 'TestA!L' + str(values.index(row) + 2)
    value_input_option = 'USER_ENTERED'
    value_range_body = {
        'values': [[today]]
    }
    request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name, valueInputOption=value_input_option, body=value_range_body)
    response = request.execute()






    element = driver.find_element(By.ID, "APIRESPONSE")
    
    if element.get_attribute('style') == 'display: none;':
        print('N')
        # Écrire done dans le google sheet
        range_name = 'TestA!C' + str(values.index(row) + 2)
        value_input_option = 'USER_ENTERED'
        value_range_body = {
            'values': [['N']]
        }
        request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name, valueInputOption=value_input_option, body=value_range_body)
        response = request.execute()
        time.sleep(5)

    else:
        
        range_name = 'TestA!C' + str(values.index(row) + 2)
        value_input_option = 'USER_ENTERED'
        value_range_body = {
            'values': [['Y']]
        }
        request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name, valueInputOption=value_input_option, body=value_range_body)
        response = request.execute()


       # trouve le div qui contient les balises li des packages
        div_element = driver.find_element(By.ID, "APIRESPONSE")

        # trouve tous les éléments li dans le div
        list_items = div_element.find_elements(By.TAG_NAME, "li")

        # compte le nombre d'éléments li
        num_list_items = len(list_items)

        # affiche le nombre d'éléments li
        print(num_list_items)

        range_name = 'TestA!D' + str(values.index(row) + 2)
        value_input_option = 'USER_ENTERED'
        value_range_body = {
            'values': [[num_list_items]]
        }
        request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name, valueInputOption=value_input_option, body=value_range_body)
        response = request.execute()


        for li in list_items:
            
            h5_element = li.find_element(By.TAG_NAME, "h5")
            if "6" in h5_element.text and "60" not in h5_element.text:
                print("6")
                div_element = li.find_element(By.CLASS_NAME, "cnt-btm")
                print(div_element.text)
                range_name = 'TestA!E' + str(values.index(row) + 2)
                value_input_option = 'USER_ENTERED'
                value_range_body = {
                    'values': [[div_element.text]]
                }
                request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name, valueInputOption=value_input_option, body=value_range_body)
                response = request.execute()


            
            h5_element = li.find_element(By.TAG_NAME, "h5")
            if "15" in h5_element.text:
                print("15")
                div_element = li.find_element(By.CLASS_NAME, "cnt-btm")
                print(div_element.text)
                range_name = 'TestA!F' + str(values.index(row) + 2)
                value_input_option = 'USER_ENTERED'
                value_range_body = {
                    'values': [[div_element.text]]
                }
                request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name, valueInputOption=value_input_option, body=value_range_body)
                response = request.execute()



                
            elif "25" in h5_element.text:
                print("25")
                div_element = li.find_element(By.CLASS_NAME, "cnt-btm")
                print(div_element.text)
                range_name = 'TestA!G' + str(values.index(row) + 2)
                value_input_option = 'USER_ENTERED'
                value_range_body = {
                    'values': [[div_element.text]]
                }
                request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name, valueInputOption=value_input_option, body=value_range_body)
                response = request.execute()

                

            elif "50" in h5_element.text:
                print("50")
                div_element = li.find_element(By.CLASS_NAME, "cnt-btm")
                div_text = div_element.text
                last_line = div_text.split("\n")[-1]
                print(last_line)
                range_name = 'TestA!H' + str(values.index(row) + 2)
                value_input_option = 'USER_ENTERED'
                value_range_body = {
                    'values': [[last_line]]
                }
                request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name, valueInputOption=value_input_option, body=value_range_body)
                response = request.execute()


            elif "60" in h5_element.text:
                print("60")
                div_element = li.find_element(By.CLASS_NAME, "cnt-btm")
                print(div_element.text)
                range_name = 'TestA!I' + str(values.index(row) + 2)
                value_input_option = 'USER_ENTERED'
                value_range_body = {
                    'values': [[div_element.text]]
                }
                request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name, valueInputOption=value_input_option, body=value_range_body)
                response = request.execute()

            elif "100" in h5_element.text:
                print("100")
                div_element = li.find_element(By.CLASS_NAME, "cnt-btm")
                print(div_element.text)
                range_name = 'TestA!J' + str(values.index(row) + 2)
                value_input_option = 'USER_ENTERED'
                value_range_body = {
                    'values': [[div_element.text]]
                }
                request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name, valueInputOption=value_input_option, body=value_range_body)
                response = request.execute()

            elif "400" in h5_element.text:
                print("400")
                div_element = li.find_element(By.CLASS_NAME, "cnt-btm")
                print(div_element.text)
                range_name = 'TestA!K' + str(values.index(row) + 2)
                value_input_option = 'USER_ENTERED'
                value_range_body = {
                    'values': [[div_element.text]]
                }
                request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name, valueInputOption=value_input_option, body=value_range_body)
                response = request.execute()


    driver.find_element(By.LINK_TEXT, "Effacer").click()

driver.quit()
