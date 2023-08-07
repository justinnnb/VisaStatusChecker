from multiprocessing.connection import wait
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager


import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from datetime import datetime

from selenium.webdriver.chrome.service import Service
service = Service()

with open('../../Json/keys.json') as f:
    config = json.load(f)

sheets_key = config['sheets_key']

# load usernames

class Database:
    def __init__(self):
        
        SERVICE_ACCOUNT_FILE = "../../Json/keys.json"
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

        self.credentials = None
        self.credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        self.clients_list_sheet = sheets_key
        service = build("sheets", "v4", credentials=self.credentials)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=self.clients_list_sheet, range="Immi Credentials").execute()
        values = result.get('values', [])
        self.database = pd.DataFrame(values)
        self.database.columns = ["Date Lodged",	"Name",	"Expiry", "Username", "Password", "Medical Exam Date", "Status", "Current Status Date", "Previous Status Date"]

    def update_sheet(self, database):
        gc = gspread.authorize(self.credentials)
        gs = gc.open_by_key(self.clients_list_sheet)
        payment_plan_worksheet = gs.worksheet('Immi Credentials')
        set_with_dataframe(worksheet=payment_plan_worksheet, dataframe=database, include_index=False,
        include_column_header=False, resize=True)

data = Database()

options = Options()
# options.headless = True

driver = webdriver.Chrome('/Users/justinbilao/Downloads/chromedriver-mac-arm64/chromedriver', options=options)
 
def main():

    for x in range(1, len(data.database["Username"])):

        # Skip if Application Status is already Finalised
        if data.database.at[x,"Status"] != "Finalised":
            # Login
            try:
                driver.get("https://online.immi.gov.au/ola/app")
                driver.find_element("name", "username").clear()
                driver.find_element("name", "username").send_keys(data.database.at[x,"Username"])
                driver.find_element("name", "password").send_keys(data.database.at[x,"Password"])
                driver.find_element("name", "login").send_keys(Keys.ENTER)
                
                driver.find_element("name", "continue").send_keys(Keys.ENTER)
            
                def update_status():
                    status_wait = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/section/div/div/div[3]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/p/strong")))
                    global status
                    status = driver.find_element("xpath", "/html/body/form/section/div/div/div[3]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/p/strong").text
                    data.database.at[x, "Status"] = status

                update_status()

                # Copy Last Update Date and Paste to Excel File 

                def last_update_date():
                    global last_update_date_cell
                    last_update_date_cell = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/section/div/div/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/time"))).text
                    # last_update_date_cell = driver.find_element("xpath", "/html/body/form/section/div/div/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/time").text

                    if last_update_date_cell == "":
                        last_update_date() # recurse until Update Date is visible/detected

                    print("New Update:", last_update_date_cell)

                    # if there is a new update, the current update will be moved to the next cell. the new update will be at the current cell.
                    if data.database.at[x, "Current Status Date"] != last_update_date_cell:
                        data.database.at[x, "Previous Status Date"] = str(data.database.at[x, "Current Status Date"])
                        print("New Update:", last_update_date_cell)

                    data.database.at[x, "Current Status Date"] = str(last_update_date_cell)

                last_update_date()

                # Print status to the terminal for each user
                user = str(data.database.at[x, "Username"])
                print("%s is %s" % (user, status))

                # Logout
                driver.find_element("xpath", "/html/body/form/header/div/div/ol/li[3]/button").click()
                driver.find_element("xpath", "/html/body/header/div/ul/li/div/a").click()

            except Exception as e:
                data.database.at[x, "Status"] = "Error"
                
    data.update_sheet(data.database)
    driver.quit()

main()