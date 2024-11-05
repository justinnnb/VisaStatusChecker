from utils.encryption import decrypt_password
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
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
#get firestore actions


async def login_to_immi(db, driver, user_id):
    try:
        print(f"Logging in to Immi for user {user_id}")
        
        # Await the database query if it supports async operations
        user = db.collection('users').where('id', '==', user_id).get()
        
        print(f"User found: {user}")
        if len(user) > 0:                               
            username = user[0].to_dict()['username']
            passkey = user[0].to_dict()['password']
            password = decrypt_password(passkey)
        else:
            print(f"User not found: {user_id}")
            return False

        driver.get("https://online.immi.gov.au/ola/app")
        driver.find_element("name", "username").clear()
        driver.find_element("name", "username").send_keys(username)
        driver.find_element("name", "password").send_keys(password)
        driver.find_element("name", "login").send_keys(Keys.ENTER)

        try:
            no_button = driver.find_element(By.XPATH, "//button[contains(.,'No')]")
            no_button.click()
        except NoSuchElementException:
            pass

        driver.find_element("name", "continue").send_keys(Keys.ENTER)
        print("Logged in to Immi")
    except Exception as e:
        print(f"Error logging in to Immi: {e}")
        return False    

    return True