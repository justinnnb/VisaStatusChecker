from multiprocessing.connection import wait
import pandas as pd

import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from datetime import datetime
from dotenv import load_dotenv
import os

from firebase_admin import firestore
import firebase_admin
from firebase_admin import credentials

from cryptography.fernet import Fernet




# Load environment variables
load_dotenv()

# Replace the config loading with environment variables
SERVICE_ACCOUNT_FILE = os.getenv('SERVICE_ACCOUNT_FILE')
sheets_key = os.getenv('GOOGLE_SHEETS_KEY')
# Add debug print to verify the values
# to connect to firestore db


db_key_path = os.getenv('FIRESTORE_DB_KEY')
cred = credentials.Certificate(db_key_path)
firebase_admin.initialize_app(cred)
db = firestore.client()
async def get_firestore_data(collection_name):
    docs = db.collection(collection_name).stream()
    data = []
    for doc in docs:
        doc_dict = doc.to_dict()
        doc_dict['id'] = doc.id
        data.append(doc_dict)
    return data

async def get_encrypted_password(username):
    user_data = await get_firestore_data('users').where('username', '==', username).get()
    for user in user_data:
        if user['username'] == username:
            return decrypt_password(user['password'])
    return None


# db applications is like this
# applications = {
#     application: {
#         'user_id': user_id,
#         'referenceNo': reference_number,
#         'visaType': visa_type,
#         'lastUpdated': last_update,
#         'submitDate': date_submitted,
#         'updates': {
#             'status': status,
#             'lastUpdated': last_update,
#         }
#     }
# }

# Add a check for the FERNET_KEY environment variable
key = os.getenv('FERNET_KEY').encode()
fernet = Fernet(key)

def encrypt_password(password):
    return fernet.encrypt(password.encode()).decode()

def decrypt_password(encrypted_password):
    try:
        return fernet.decrypt(encrypted_password).decode()
    except Exception as e:
        print(f"Decryption error: {e}", encrypted_password)
        return None

def parse_date(date_str):
    if not date_str:
        return None
    try:
        # Assuming the date format is 'DD MMM YYYY', e.g., '17 Aug 2023'
        return datetime.strptime(date_str, '%d %b %Y')
    except ValueError as e:
        print(f"Date parsing error: {e}")
        return None
    
async def get_application_data(db, reference_number):
    application = db.collection('applications').where('referenceNo', '==', reference_number).get()
    if application:
        print(application[0].to_dict())
        return application[0].to_dict()
    return False

async def update_application(db, api_response):
    # Parse dates
    last_update = parse_date(api_response['last_updated'])
    reference_no = api_response['reference_number']
        
    user_application_data = await get_application_data(db, reference_no)

    if user_application_data:
        # if lastUpdated in user_application_data is different from last_update, then update the updates
        if user_application_data.get('lastUpdated') != last_update:
            application_data = {        
                'lastUpdated': last_update,
            }

            update_data = {
                'status': api_response['status'],
                'lastUpdated': last_update
            }
            # Let Firestore auto-generate the document ID for updates
            update_ref = db.collection('applications').where('referenceNo', '==', reference_no).collection('updates').get()
            # if date not in update_ref, then add it. if present, check if status is different. if different, update. if same, do nothing.
            # convert update_ref to a list of objects
            update_ref = [obj.to_dict() for obj in update_ref]
            if not update_ref:
                db.collection('applications').where('referenceNo', '==', reference_no).collection('updates').add(update_data)
            else:
                #if latest date status is different from the new status, then update
                # list contains dicts with date and status
                latest_date = max(update_ref, key=lambda x: x['lastUpdated'])
                if latest_date['status'] != api_response['status']:
                    db.collection('applications').where('referenceNo', '==', reference_no).collection('updates').add(update_data)

    else:
        # if the application is not in the database, then add it
        application_data = {
            'referenceNo': api_response['reference_number'],
            'lastUpdated': last_update,
            'submitDate': api_response['date_submitted'],
        }
        updates = {
            'status': api_response['status'],
            'lastUpdated': last_update
        }

        print(application_data, " is the application data")
        # Use reference number as document ID
        application_ref = db.collection('applications').document(reference_no).set(application_data)
        updates_ref = db.collection('applications').document(reference_no).collection('updates').add(updates)

        return application_ref, updates_ref

class Database:
    def __init__(self):
        
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

        self.credentials = None
        self.credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        self.clients_list_sheet = sheets_key
        service = build("sheets", "v4", credentials=self.credentials)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=self.clients_list_sheet, range="Immi Credentials").execute()
        
        # Add error checking for the API response
        if not result or 'values' not in result:
            raise Exception("Failed to fetch data from Google Sheets or sheet is empty")
            
        values = result.get('values', [])
        if not values:
            raise Exception("No data found in the specified range")
            
        self.database = pd.DataFrame(values)
        self.database.columns = ["Date Lodged",	"Name",	"Expiry", "Username", "Password", "Medical Exam Date", "Status", "Current Status Date", "Previous Status Date", "User ID"]

    def update_sheet(self, database):
        gc = gspread.authorize(self.credentials)
        gs = gc.open_by_key(self.clients_list_sheet)
        payment_plan_worksheet = gs.worksheet('Immi Credentials')
        set_with_dataframe(worksheet=payment_plan_worksheet, dataframe=database, include_index=False,
        include_column_header=False, resize=True)

data = Database()


# service = Service(ChromeDriverManager().install())
options = Options()
# options.add_argument("--headless=new")

driver = webdriver.Chrome(options=options)
# driver = webdriver.Chrome(service=service,options=options)

async def main():
    users_ref = db.collection('users')

    for x in range(1, len(data.database["Username"])):

        # Skip if Application Status is already Finalised
        if data.database.at[x,"Status"] != "Finalised":
            # Login
            if data.database.at[x,"Username"] == "etchen.sordilla":




                
                continue

            if data.database.at[x,"Username"] == "cabeljennifervisa@gmail.com":
            #add new user
                    driver.get("https://online.immi.gov.au/ola/app")

                    user = db.collection('users').where('username', '==', data.database.at[x,"Username"]).get()
                    username = user[0].to_dict()['username']
                    passkey = user[0].to_dict()['password']
                    password = decrypt_password(passkey)
                    driver.find_element("name", "username").clear()
                    driver.find_element("name", "username").send_keys(username)
                    driver.find_element("name", "password").send_keys(password)
                    driver.find_element("name", "login").send_keys(Keys.ENTER)
                    break

            try:
                driver.get("https://online.immi.gov.au/ola/app")
                driver.find_element("name", "username").clear()
                driver.find_element("name", "username").send_keys(data.database.at[x,"Username"])
                driver.find_element("name", "password").send_keys(data.database.at[x,"Password"])
                driver.find_element("name", "login").send_keys(Keys.ENTER)

                try:
                    no_button = driver.find_element(By.XPATH, "//button[contains(.,'No')]")
                    no_button.click()
                except NoSuchElementException:
                    pass

                driver.find_element("name", "continue").send_keys(Keys.ENTER)

                def update_status():
                    status_wait = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/section/div/div/div[3]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/p/strong")))
                    global status
                    status = driver.find_element("xpath", "/html/body/form/section/div/div/div[3]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/p/strong").text
                    
                    data.database.at[x, "Status"] = status


                # Copy Last Update Date and Paste to Excel File 

                def last_update_date():
                    global last_update_date_cell
                    last_update_date_cell = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/section/div/div/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/time"))).text
                    # last_update_date_cell = driver.find_element("xpath", "/html/body/form/section/div/div/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/time").text

                    if last_update_date_cell == "":
                        last_update_date() # recurse until Update Date is visible/detected


                    # if there is a new update, the current update will be moved to the next cell. the new update will be at the current cell.
                    if data.database.at[x, "Current Status Date"] != last_update_date_cell:
                        data.database.at[x, "Previous Status Date"] = str(data.database.at[x, "Current Status Date"])
                        # print("New Update:", last_update_date_cell)

                    print("New Update:", last_update_date_cell)
                    data.database.at[x, "Current Status Date"] = str(last_update_date_cell)

                # update_status()
                # last_update_date()

                # Add this section to handle tabs and their content
                tabs = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[role='tab']"))
                )
                #
                no_of_applications = len(tabs)
                applications = []
                # Click through each tab and get its content
                for x in range(0, no_of_applications):
                    try:
                        # Click the tab to make its panel visible
                        #results tab ID is MyAppsResultTab_x
                        #results content tab ID is MyAppsResultTab_x-content
                        tab = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, f"MyAppsResultTab_{x}"))
                        )
                        if tab.get_attribute('aria-expanded') == "false":
                            tab.click()
                        tab_content = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, f"MyAppsResultTab_{x}-content"))
                        )
                        if tab_content:
                            #get the status
                            status = tab.find_element(By.XPATH, ".//div/p/strong").text
                            # /html/body/form/section/div/div/div[3]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/p/strong
                            #get text reference
                            reference_number = tab_content.find_element(By.XPATH, ".//div/div/div[1]/div/div[1]/div/div/div[1]/div/div/div[1]/div/div/div/div")
                            # /div/div/div[1]/div/div[1]/div/div/div[1]/div/div/div[1]/div/div/div/div
                            #get visa type
                            visa_type = tab_content.find_element(By.XPATH, ".//div/div/div[1]/div/div[1]/div/div/div[1]/div/div/div[2]/div/div/div/div")
                            # /div/div/div[1]/div/div[1]/div/div/div[1]/div/div/div[1]/div/div/div/div/div[1]/div/div/div/div
                            # get last update
                            last_update = tab_content.find_element(By.XPATH, ".//div/div/div[1]/div/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/div/time")
                            # /div/div/div[1]/div/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/div/time
                            # get date submitted
                            date_submitted = tab_content.find_element(By.XPATH, ".//div/div/div[1]/div/div[1]/div/div/div[2]/div/div/div[2]/div/div/div/div/time")
                            # /div/div/div[1]/div/div[1]/div/div/div[2]/div/div/div[2]/div/div/div/div/time

                        updates = {
                            "status": status,
                            "last_update": last_update.text
                        }

                        application = {
                            "reference_number": reference_number.text,
                            "last_updated": last_update.text,
                            "status": status,
                            "visa_type": visa_type.text,
                            "date_submitted": date_submitted.text,
                            "updates": updates
                        }
                        
                        update_ref = await update_application(db, application)
                        print(f"Update reference: {update_ref}")
                        
                    except Exception as e:
                        print(f"Error accessing tab: {str(e)}")

                # Print status to the terminal for each user5
                user = str(data.database.at[x, "Username"])
                print("%s is %s" % (user, status))

                # Logout
                driver.find_element("xpath", "/html/body/form/header/div/div/ol/li[3]/button").click()
                driver.find_element("xpath", "/html/body/header/div/ul/li/div/a").click()

                break

            except Exception as e:
                data.database.at[x, "Status"] = e

    data.update_sheet(data.database)
    driver.quit()

import asyncio
asyncio.run(main())