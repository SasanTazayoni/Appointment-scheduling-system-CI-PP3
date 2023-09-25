import gspread
from google.oauth2.service_account import Credentials
import colorama
from colorama import Fore, Back
colorama.init(autoreset=True)

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('python_scheduling_system')

def login_prompt():
    """
    Prompt the user for a login and a password
    """
    print(Fore.GREEN + 'Appointment Booking System')

    print('Please enter the correct login details\n')

    while True:
        login = input("Username: \n")
        
        if not login.strip():
            print(Fore.RED + "Username cannot be empty.\n")
            continue
        
        while True:
            password = input("Password: \n")

            if not password.strip():
                print(Fore.RED + "Password cannot be empty.\n")
                continue
            
            if login == "admin" and password == "password":
                print(Fore.GREEN + "Login successful!")
                return
            else:
                print(Fore.RED + "Login failed. Please try again.\n")
                break

login_prompt()