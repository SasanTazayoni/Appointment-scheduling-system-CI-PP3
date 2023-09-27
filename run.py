import gspread
from google.oauth2.service_account import Credentials
from config import USERNAME, PASSWORD
import colorama
from colorama import Fore, Back
colorama.init(autoreset=True)
import datetime
import time
import os

# Define the scope for Google Sheets API
SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

# Initialize Google Sheets API credentials
CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('appointment_scheduling_system')

# Define maximum login attempts and lockout duration
MAX_LOGIN_ATTEMPTS = 3
LOCKOUT_DURATION = 10

def handle_lockout():
    """
    Handle the lockout behavior.
    """
    locked_out = True
    
    # Disable keyboard input
    os.system('stty -echo')
    
    try:
        for remaining_lockout_time in range(LOCKOUT_DURATION, 0, -1):
            # Print lockout message
            print(Fore.RED + f"Maximum login attempts reached. You are now locked out for {remaining_lockout_time} seconds.", end="\r")
            time.sleep(1)
    finally:
        # Enable keyboard input after lockout
        os.system('stty echo')
    
    login_prompt()

def login_prompt():
    """
    Prompt the user for a login and a password
    """
    print()
    print(Fore.BLUE + 'Appointment Booking System\n')

    login_attempts = 0
    locked_out = False

    if locked_out:
        locked_out = handle_lockout()

    print(Fore.BLUE + 'Please enter the correct login details\n')

    while login_attempts < MAX_LOGIN_ATTEMPTS and not locked_out:
        login = input("Username: \n")
        
        if not login.strip():
            print(Fore.RED + "Username cannot be empty.\n")
            continue
        
        while True:
            password = input("Password: \n")

            if not password.strip():
                print(Fore.RED + "Password cannot be empty.\n")
                continue
            
            if login == USERNAME and password == PASSWORD:
                print(Fore.GREEN + "Login successful!\n")
                login_attempts = 0
                update_cell_dates()
                return
            else:
                login_attempts += 1
                if login_attempts == MAX_LOGIN_ATTEMPTS:
                    print(Fore.RED + "Login failed.\n")
                    locked_out = handle_lockout()
                else:
                    print(Fore.RED + "Login failed. Please try again.\n")
                break
                
def update_cell_dates():
    """
    Update worksheet based on login date and time.
    """
    current_datetime = datetime.datetime.now()
    monday_date = get_monday_date(current_datetime)

    # Get Monday's date for the first week on the spreadsheet
    w1_worksheet_cell_value = SHEET.worksheet("week1").acell("A2").value

    difference_in_weeks = get_week_difference(monday_date, w1_worksheet_cell_value)

    # Update worksheets based on the difference in weeks
    if difference_in_weeks == 0:
        print(Fore.BLUE + "Worksheets are up to date")
    elif difference_in_weeks > 12:
        print(Fore.BLUE + "It seems that you have been away for a long time.\n")
        print(Fore.YELLOW + "Renewing worksheets...")

        for worksheet in SHEET.worksheets():
            refresh_cells(worksheet)
            set_cell_dates(worksheet, current_datetime)

        print(Fore.GREEN + "All worksheets have been updated.")
    elif difference_in_weeks > 0:
        print(Fore.YELLOW + "Updating worksheets...")

        for i in range(1, 13):
            worksheet_title = f"week{i}"
            worksheet = SHEET.worksheet(worksheet_title)
            week_number = int(worksheet_title.lstrip("week"))
            target_week_number = week_number + difference_in_weeks

            # Check if the target week number is within the valid range (1 to 12)
            if 1 <= target_week_number <= 12:
                target_worksheet_title = f"week{target_week_number}"
                target_worksheet = SHEET.worksheet(target_worksheet_title)
                target_data = target_worksheet.get('B2:Q6')
                worksheet.update('B2:Q6', target_data)
                set_cell_dates(worksheet, current_datetime)

                print(f"{worksheet_title} updated from {target_worksheet_title}")
            else:
                # If the target week number is not within the valid range, update cells B2:Q6 with 'OPEN'
                cell_list = worksheet.range('B2:Q6')
                for cell in cell_list:
                    cell.value = 'OPEN'
                worksheet.update_cells(cell_list)

        print(Fore.GREEN + "Worksheets have been updated.")

def get_monday_date(current_datetime):
    """
    Get Monday's date of the current week.
    """
    days_since_monday = current_datetime.weekday()
    monday_date = current_datetime - datetime.timedelta(days=days_since_monday)
    return monday_date

def get_week_difference(monday_date, w1_worksheet_cell_value):
    """
    Determine the difference in weeks between the current week and the spreadsheet's first week.
    """
    # Extract the date part from the cell value
    date_part = w1_worksheet_cell_value.split("(")[1].split(")")[0].strip()
    # Convert it to a datetime object
    cell_date = datetime.datetime.strptime(date_part, "%d-%m-%Y")
    # Calculate the difference in weeks
    difference_in_weeks = (monday_date - cell_date).days // 7
    return difference_in_weeks

def set_cell_dates(worksheet, current_datetime):
    """
    Set cell dates based on the current date and the worksheet name.
    """
    worksheet_names = [
        "week1", "week2", "week3", "week4",
        "week5", "week6", "week7", "week8",
        "week9", "week10", "week11", "week12"
    ]
    
    # Get the index of the worksheet in the list
    worksheet_index = worksheet_names.index(worksheet.title)

    # Calculate the start date for the worksheet
    days_since_monday = (current_datetime.weekday() - 0) % 7
    start_date = current_datetime - datetime.timedelta(days=days_since_monday) - datetime.timedelta(weeks=0) + datetime.timedelta(weeks=worksheet_index)

    # Generate dates for the worksheet
    dates = [start_date + datetime.timedelta(days=j) for j in range(5)]

    # Format dates as strings
    formatted_dates = [f"{date.strftime('%A')} ({date.strftime('%d-%m-%Y')})" for date in dates]

    # Update the A2:A6 cells with the formatted dates
    cell_values = worksheet.range("A2:A6")
    for j, date in enumerate(formatted_dates):
        cell_values[j].value = date

    # Update the cells in the worksheet
    worksheet.update_cells(cell_values)

def refresh_cells(worksheet):
    """
    Clear all appointment bookings from old dates.
    """
    # Get a list of cells in the specified range
    cell_list = worksheet.range('B2:Q6')
    # Set the value of each cell to 'OPEN'
    for cell in cell_list:
        cell.value = 'OPEN'
    # Update the cells in the worksheet
    worksheet.update_cells(cell_list)

login_prompt()

def fix_cell_dates():
    """
    This function is entirely for testing purposes. 
    The dates of the spreadsheet can be set manually.
    """
    
    print(Fore.YELLOW + 'Setting dates...')
    worksheet_names = [ "week1", "week2", "week3", "week4", "week5", "week6", "week7", "week8", "week9", "week10", "week11", "week12" ]
    current_datetime = datetime.datetime.now()
    
    for i, worksheet_name in enumerate(worksheet_names):
        worksheet = SHEET.worksheet(worksheet_name)

        days_since_monday = (current_datetime.weekday() - 0) % 7
        start_date = current_datetime - datetime.timedelta(days=days_since_monday) - datetime.timedelta(weeks=2) + datetime.timedelta(weeks=i)

        dates = [start_date + datetime.timedelta(days=j) for j in range(5)]

        formatted_dates = [f"{date.strftime('%A')} ({date.strftime('%d-%m-%Y')})" for date in dates]

        cell_values = worksheet.range("A2:A6")

        for j, date in enumerate(formatted_dates):
            cell_values[j].value = date

        worksheet.update_cells(cell_values)

# fix_cell_dates()