import gspread
from google.oauth2.service_account import Credentials
import colorama
from colorama import Fore, Back
colorama.init(autoreset=True)
import datetime

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('appointment_scheduling_system')

def login_prompt():
    """
    Prompt the user for a login and a password
    """
    print(Fore.GREEN + 'Appointment Booking System\n')

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
                print(Fore.GREEN + "Login successful!\n")
                update_cell_dates()
                return
            else:
                print(Fore.RED + "Login failed. Please try again.\n")
                break
                
def update_cell_dates():
    
    print(Fore.GREEN + 'Updating worksheet...')
    worksheet_names = ["week1", "week2", "week3", "week4", "week5", "week6", "week7", "week8", "week9", "week10", "week11", "week12"]
    current_datetime = datetime.datetime.now()
    
    for i, worksheet_name in enumerate(worksheet_names):
        worksheet = SHEET.worksheet(worksheet_name)

        days_since_monday = (current_datetime.weekday() - 0) % 7
        start_date = current_datetime - datetime.timedelta(days=days_since_monday) + datetime.timedelta(weeks=i)

        dates = [start_date + datetime.timedelta(days=j) for j in range(5)]

        formatted_dates = [f"{date.strftime('%A')} ({date.strftime('%d-%m-%Y')})" for date in dates]

        cell_values = worksheet.range("A2:A6")

        for j, date in enumerate(formatted_dates):
            cell_values[j].value = date

        worksheet.update_cells(cell_values)

login_prompt()

