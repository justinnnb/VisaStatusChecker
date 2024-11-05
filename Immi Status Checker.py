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
from datetime import datetime, timezone
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
#         'reference_number': reference_number,
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
        return datetime.strptime(date_str, '%d %b %Y').replace(tzinfo=timezone.utc)
    except ValueError as e:
        print(f"Date parsing error: {e}")
        return None
    
async def get_application_data(db, reference_number):
    application = db.collection('applications').where('reference_number', '==', reference_number).get()
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
            update_ref = db.collection('applications').where('reference_number', '==', reference_no).collection('updates').get()
            # if date not in update_ref, then add it. if present, check if status is different. if different, update. if same, do nothing.
            # convert update_ref to a list of objects
            update_ref = [obj.to_dict() for obj in update_ref]
            if not update_ref:
                db.collection('applications').where('reference_number', '==', reference_no).collection('updates').add(update_data)
            else:
                #if latest date status is different from the new status, then update
                # list contains dicts with date and status
                latest_date = max(update_ref, key=lambda x: x['lastUpdated'].replace(tzinfo=timezone.utc))
                if latest_date['status'] != api_response['status']:
                    db.collection('applications').where('reference_number', '==', reference_no).collection('updates').add(update_data)

    else:
        # if the application is not in the database, then add it
        application_data = {
            'reference_number': api_response['reference_number'],
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
options.add_argument("--headless=new")

driver = webdriver.Chrome(options=options)
# driver = webdriver.Chrome(service=service,options=options)

async def main():
    users_ref = db.collection('users')

    for x in range(1, len(data.database["Username"])):
        # get applications where latest update is not finalised

        # Skip if Application Status is already Finalised
        if data.database.at[x,"Status"] != "Finalised":
            # Login
            # if data.database.at[x,"Username"] == "etchen.sordilla":
            #     continue

            #add new user
            print(data.database.at[x,"Username"], " is the username")

            try:

                driver.get("https://online.immi.gov.au/ola/app")

                user = db.collection('users').where('username', '==', data.database.at[x,"Username"]).get()
                user_doc_id = user[0].id
                username = user[0].to_dict()['username']
                passkey = user[0].to_dict()['password']

                password = decrypt_password(passkey)
                driver.find_element("name", "username").clear()
                driver.find_element("name", "username").send_keys(username)
                driver.find_element("name", "password").send_keys(password)
                driver.find_element("name", "login").send_keys(Keys.ENTER)

                # try:
                #     no_button = driver.find_element(By.XPATH, "//button[contains(.,'No')]")
                #     no_button.click()
                # except NoSuchElementException:
                #     pass

                driver.find_element("name", "continue").send_keys(Keys.ENTER)

                # Add this section to handle tabs and their content
                tabs = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[role='tab']"))
                )
                #
                no_of_applications = len(tabs)
                applications = []
                # Click through each tab and get its content
                for y in range(0, no_of_applications):
                    try:
                        # Click the tab to make its panel visible
                        #results tab ID is MyAppsResultTab_x
                        #results content tab ID is MyAppsResultTab_x-content

                        tab = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, f"MyAppsResultTab_{y}"))
                        )
                        if tab.get_attribute('aria-expanded') == "false":
                            tab.click()
                        tab_content = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, f"MyAppsResultTab_{y}-content"))
                        )
                        if tab_content:
                            #get the status
                            status = WebDriverWait(tab, 10).until(
                                EC.text_to_be_present_in_element((By.XPATH, ".//div/p/strong"), "")
                            )
                            status = tab.find_element(By.XPATH, ".//div/p/strong").text
                            #get text reference
                            reference_number = WebDriverWait(tab_content, 10).until(
                                EC.text_to_be_present_in_element((By.XPATH, ".//div/div/div[1]/div/div[1]/div/div/div[1]/div/div/div[1]/div/div/div/div"), "")
                            )
                            reference_number = tab_content.find_element(By.XPATH, ".//div/div/div[1]/div/div[1]/div/div/div[1]/div/div/div[1]/div/div/div/div").text
                            #get visa type
                            visa_type = WebDriverWait(tab_content, 10).until(
                                EC.text_to_be_present_in_element((By.XPATH, ".//div/div/div[1]/div/div[1]/div/div/div[1]/div/div/div[2]/div/div/div/div"), "")
                            )
                            visa_type = tab_content.find_element(By.XPATH, ".//div/div/div[1]/div/div[1]/div/div/div[1]/div/div/div[2]/div/div/div/div").text
                            # get last update

                            last_update_check = False
                            while not last_update_check:
                                last_update_check = WebDriverWait(tab_content, 10).until(
                                    EC.text_to_be_present_in_element((By.XPATH, ".//div/div/div[1]/div/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/div/time"), "")
                                )
                            last_update = tab_content.find_element(By.XPATH, ".//div/div/div[1]/div/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/div/time").text
                                                        
                            #convert from selenium element to string
                            if last_update == "":
                                print("No last update found")
                            # last_update = datetime.strptime(last_update, '%d %b %Y')
                            # get date submitted
                            date_submitted = WebDriverWait(tab_content, 10).until(
                                EC.text_to_be_present_in_element((By.XPATH, ".//div/div/div[1]/div/div[1]/div/div/div[2]/div/div/div[2]/div/div/div/div/time"), "")
                            )
                            date_submitted = tab_content.find_element(By.XPATH, ".//div/div/div[1]/div/div[1]/div/div/div[2]/div/div/div[2]/div/div/div/div/time").text
                            date_submitted = datetime.strptime(date_submitted, '%d %b %Y')

                        updates = {
                            "status": status,
                            "update_date": datetime.strptime(last_update, '%d %b %Y')
                        }

                        application = {
                            "reference_number": reference_number,
                            "last_updated": datetime.strptime(last_update, '%d %b %Y'),
                            "status": status,
                            "visa_type": visa_type,
                            "date_submitted": date_submitted,
                            "user_id": user_doc_id,
                        }

                        #add application to firestore. find reference number in applications. if not found, add. if found, update.
                        application_ref = db.collection('applications').where('reference_number', '==', reference_number).get()

                        if application_ref:
                            #update application
                            updates_query = db.collection('applications').document(application_ref[0].id).collection('updates')\
                                .order_by('update_date', direction=firestore.Query.DESCENDING).limit(1).get()
                            updates_ref = [doc.to_dict() for doc in updates_query]

                            #check last_updated in updates_ref. if different from last_updated in application, then add.
                            #check if latest last updated is different from application last updated. if different, add. 
                            if updates_ref:
                                latest_update = updates_ref[0]
                                latest_update_date = latest_update.get('update_date')
                                latest_update_status = latest_update.get('status')
                            else:
                                latest_update_date = None
                                latest_update_status = None

                            print(latest_update_date, updates['update_date'], "compare")
                            if latest_update_date == None:
                                db.collection('applications').document(application_ref[0].id).collection('updates').add(updates)

                            # Convert both dates to UTC and remove microseconds for comparison
                            elif (latest_update_date.astimezone(timezone.utc).replace(microsecond=0) < 
                                  updates['update_date'].astimezone(timezone.utc).replace(microsecond=0)):
                                db.collection('applications').document(application_ref[0].id).collection('updates').add(updates)

                            # Also normalize dates for status comparison
                            elif (latest_update_status != updates['status'] and 
                                  latest_update_date.astimezone(timezone.utc).replace(microsecond=0) == 
                                  updates['update_date'].astimezone(timezone.utc).replace(microsecond=0)):
                                db.collection('applications').document(application_ref[0].id).collection('updates').add(updates)

                            else:
                                print(f"No updates needed for application {reference_number}")
                        else:
                            # add the application to firestore
                            new_application_ref = db.collection('applications').add(application)
                            application_ref = db.collection('applications').where('reference_number', '==', reference_number).get()
                            new_id = application_ref[0].id
                            print(new_id, " is the new application id")
                            db.collection('applications').document(new_id).collection('updates').add(updates)
                            print("Application not found. Created.", x)


                        # query again to check if application status and last update date is equal to latest update
                        application_ref = db.collection('applications').where('reference_number', '==', reference_number).get()
                        if application_ref:
                            application_data = application_ref[0].to_dict()
                            if application_data['status'] == updates['status'] and application_data['last_updated'] == updates['update_date']:
                                print(f"Application {reference_number} is finalised")
                        
                    except Exception as e:
                        print(f"Error accessing tab: {str(e)}")

                # Print status to the terminal for each user
                user = str(data.database.at[x, "Username"])
                print("%s is %s" % (user, status))

                # Logout
                driver.find_element("xpath", "/html/body/form/header/div/div/ol/li[3]/button").click()
                driver.find_element("xpath", "/html/body/header/div/ul/li/div/a").click()

            except Exception as e:
                data.database.at[x, "Status"] = e

    data.update_sheet(data.database)
    driver.quit()

import asyncio
asyncio.run(main())