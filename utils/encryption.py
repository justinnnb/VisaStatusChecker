import os
from dotenv import load_dotenv
from cryptography.fernet import Fernet

load_dotenv()

key = os.getenv('FERNET_KEY').encode()
fernet = Fernet(key)

def encrypt_password(password):
    return fernet.encrypt(password.encode()).decode()

def decrypt_password(encrypted_password):
    try:
        return fernet.decrypt(encrypted_password).decode()
    except Exception as e:
        print(f"Decryption error: {e}")
        return None
