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

    print(Fore.BLUE + 'Appointment Booking System\n')
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
    """
    Update worksheet based on login date and time.
    """
    # Get Monday's date of the current week
    current_datetime = datetime.datetime.now()
    days_since_monday = current_datetime.weekday()
    monday_date = current_datetime - datetime.timedelta(days=days_since_monday)

    # Get Monday's date for the first week on the spreadsheet
    w1_worksheet_cell_value = SHEET.worksheet("week1").acell("A2").value

    # Determine the difference in weeks between the 2
    date_part = w1_worksheet_cell_value.split("(")[1].split(")")[0].strip()
    cell_date = datetime.datetime.strptime(date_part, "%d-%m-%Y")
    difference_in_weeks = (monday_date - cell_date).days // 7

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

def set_cell_dates(worksheet, login_date):
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
    cell_list = worksheet.range('B2:Q6')
    for cell in cell_list:
        cell.value = 'OPEN'
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