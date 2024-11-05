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


from browser_actions import login_to_immi
from utils.encryption import decrypt_password

import db.actions as db_actions

# Load environment variables
load_dotenv()

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



try:
    db = db_actions.get_db()
except Exception as e:
    print(f"Error getting database: {e}")
    quit()

# service = Service(ChromeDriverManager().install())
options = Options()
# options.add_argument("--headless=new")

driver = webdriver.Chrome(options=options)
# driver = webdriver.Chrome(service=service,options=options)

async def main():
    not_finalised_applications = db.collection('applications').where('status', 'not-in', ['Finalised', 'Incomplete']).get()

    # Convert to dict
    not_finalised_applications = [doc.to_dict() for doc in not_finalised_applications]
    #remove my health declarations
    not_finalised_applications = [application for application in not_finalised_applications if application['visa_type'] != "My Health Declarations"]
    
    for application in not_finalised_applications:
        user_doc_id = application['user_id']

        if user_doc_id:
            try:

                user = db.collection('users').document(application['user_id']).get()
                user = user.to_dict()

                # breakpoint()

                if user:
                    await login_to_immi(user, driver)

                # Add this section to handle tabs and their content
                tabs = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[role='tab']"))
                )

                # Click through each tab and get its content
                no_of_applications = len(tabs)
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
                        try: 
                            application_ref_check = db.collection('applications').where('reference_number', '==', reference_number).get()
                            application_update_ref_check = db.collection('applications').document(application_ref_check[0].id).collection('updates')\
                                .order_by('update_date', direction=firestore.Query.DESCENDING).limit(1).get()
                        
                            if application_update_ref_check:
                                application_ref_check = application_ref_check[0].to_dict()
                                application_data_check = application_update_ref_check[0].to_dict()

                                if application_ref_check['status'] == application_data_check['status'] and \
                                    application_ref_check['last_updated'] == application_data_check['update_date']:
                                    print("Checked application status and last update date is equal to latest update for ", reference_number)
                                else:
                                    print("Checked application status and last update date is not equal to latest update for ", reference_number)

                        except Exception as e:  
                            print(f"Error checking application status: {str(e)}")
                        
                    except Exception as e:
                        print(f"Error accessing tab: {str(e)}")


                # Logout
                driver.find_element("xpath", "/html/body/form/header/div/div/ol/li[3]/button").click()
                driver.find_element("xpath", "/html/body/header/div/ul/li/div/a").click()

            except Exception as e:
                print(f"Error checking application status: {str(e)}")

    driver.quit()

import asyncio
asyncio.run(main())