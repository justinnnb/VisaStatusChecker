import os
import firebase_admin
from firebase_admin import credentials, firestore

db_key_path = os.getenv('FIRESTORE_DB_KEY')
cred = credentials.Certificate(db_key_path)
firebase_admin.initialize_app(cred)
db = firestore.client()

def get_db():
    return db
