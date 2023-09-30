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
            print(Fore.RED + f"Maximum login attempts reached. You are now locked out for {remaining_lockout_time} second(s).", end="\r")
            time.sleep(1)
    finally:
        # Enable keyboard input after lockout
        os.system('stty echo')
    
    login_prompt()

def login_prompt():
    """
    Ask whether the user would like to log in and present the name of the application
    """
    print()
    print(Fore.BLUE + 'Appointment Booking System\n')

    while True:
        choice = input('Would you like to log in? (yes/no): \n').strip().lower()

        if choice == 'yes':
            login()
            break
        elif choice == 'no':
            print()
            print(Fore.YELLOW + "Exiting the program.")
            exit()
        elif not choice:
            print(Fore.RED + "The input cannot be empty. Please enter 'yes' or 'no'.\n")
        else:
            print(Fore.RED + "Invalid choice. Please enter 'yes' or 'no'.\n")

def login():
    """
    Ask the user for a login and a password
    """

    login_attempts = 0
    locked_out = False

    if locked_out:
        locked_out = handle_lockout()

    print()
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
                print()
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

    # Prompt the user to select a week
    pick_week()

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

def pick_week():
    """
    Ask the user to pick a week from the spreadsheet or log out.
    """
    # Generate week titles from Week1 to Week12
    week_titles = [f"week{i}" for i in range(1, 13)]

    while True:
        print(Fore.BLUE + "Please select a number between 1 to 12 where week1 represents the current week or enter '0' to log out:\n")
    
        # Display week options
        for i, title in enumerate(week_titles, start=1):
            print(f"{Fore.BLUE}[{i}] {Fore.WHITE}{title}", end=f"{Fore.WHITE}, ")

        # Display the log out option 
        print(f"{Fore.BLUE}[0] {Fore.WHITE}Log out")
        print()
        
        try:
            choice = input("Enter the number of the week you want to select (or '0' to log out): ")
            if choice == '0':
                login_prompt()  # Log out and trigger the login function
            else:
                choice = int(choice)
                if 1 <= choice <= len(week_titles):
                    selected_week = week_titles[choice - 1]
                    get_dates_from_worksheet(selected_week)  # Call the function with the selected week to find the days of the week
                    break
                else:
                    print(Fore.RED + "Appointments in the past and appointments beyond 12 weeks in the future are not accessible. Please enter a number between 0 and 12.")
        except ValueError:
            print(Fore.RED + "Invalid input. Please enter a number between 0 and 12.")

def get_dates_from_worksheet(selected_week):
    """
    Get the dates from cells A2 to A6 in the selected worksheet.
    """
    print(Fore.YELLOW + 'Retrieving dates...')
    worksheet = SHEET.worksheet(selected_week)
    date_values = worksheet.range("A2:A6")
    dates = [date.value for date in date_values]
    pick_day(dates, selected_week)

def pick_day(dates, selected_week):
    """
    Ask the user to pick a day of the week from the selected week in the spreadsheet.
    """
    # Print the dates and options
    start_date = dates[0]
    print(Fore.GREEN + f"Retrieved dates for week beginning on '{start_date}':")
    for i, date in enumerate(dates, start=1):
        print(f"{Fore.BLUE}[{i}] {Fore.WHITE}{date}")

    # Ask the user to pick a date or enter '0' to repick week
    while True:
        try:
            choice = input("Enter the number of the date you want to select (or '0' to repick week):\n")
            choice = int(choice)

            # Check if the user wants to return to week selection
            if choice == 0:
                pick_week()  # Return to week selection
                break
            # Check if the user selected a valid date
            elif 1 <= choice <= len(dates):
                selected_date = dates[choice - 1]
                print(Fore.GREEN + f"You selected: {selected_date}")
                display_appointment_slots(selected_date, selected_week)
                break
            else:
                print(Fore.RED + "Invalid choice. Please enter a valid number.")
        except ValueError:
            print(Fore.RED + "Invalid input. Please enter a number or '0' to repick day.")

def retrieve_appointment_slots(selected_date, selected_week):
    """
    Retrieve appointment slots for the selected day in the selected week.
    """
    # Extract the day name from the selected_date
    selected_day = selected_date.split(" ")[0]
    
    # Map day names to column indices (Monday: 2, Tuesday: 3, Wednesday: 4, Thursday: 5, Friday: 6)
    day_to_column = {
        "Monday": 2,
        "Tuesday": 3,
        "Wednesday": 4,
        "Thursday": 5,
        "Friday": 6
    }

    # Get the column index for the selected day
    column_index = day_to_column.get(selected_day)

    # Get the selected worksheet
    worksheet = SHEET.worksheet(selected_week)

    # Get appointment slots for the selected day
    slots = worksheet.range(f"B{column_index}:Q{column_index}")

    # Get the time slots from row 2 (A2:Q2)
    time_slots_range = worksheet.range(f"B1:Q1")
    time_slots = [cell.value for cell in time_slots_range]

    # Create a dictionary to store time slots and their statuses
    time_slot_status = {}

    # Iterate through the appointment slots and check availability
    for i, slot in enumerate(slots):
        time_slot = time_slots[i]
        if slot.value == "OPEN":
            time_slot_status[time_slot] = "OPEN"
        elif slot.value == "BOOKED":
            time_slot_status[time_slot] = "BOOKED"
        else:
            time_slot_status[time_slot] = "BLOCKED"

    return time_slot_status

def display_appointment_slots(selected_date, selected_week):
    """
    Display appointment slots for the selected day in the selected week.
    """
    print(Fore.YELLOW + f"Retrieving appointment slots...")
    
    all_slots = retrieve_appointment_slots(selected_date, selected_week)

    # Present visual display for time slots
    print()
    print(Fore.GREEN + f"Retrieved appointment slots for {Fore.BLUE}{selected_date}:\n")

    color_codes = {
        "BLOCKED": Fore.RED,
        "BOOKED": Fore.BLUE,
        "OPEN": Fore.GREEN
    }
    
    # Iterate through the dictionary and format the slots
    formatted_slots = " ".join([f"{time_slot} {color_codes[status]}{status} {Fore.RESET}" for time_slot, status in all_slots.items()])
    print(formatted_slots)

    print()
    select_appointment_slot(all_slots, selected_date, selected_week)

def select_appointment_slot(all_slots, selected_date, selected_week):
    """
    Prompt the user to select a single time slot or a time range.
    """

    while True:
        try:
            choice = input("Enter the desired appointment time or range (e.g. '09:00' or '10:30-11:30'),'cancel' to return to date selection or 'logout' to log out: \n")

            if choice.lower() == 'cancel':
                print(Fore.YELLOW + 'Returning to previous menu...')
                get_dates_from_worksheet(selected_week)  # Return to date selection
                return
            elif choice.lower() == 'logout':
                print(Fore.YELLOW + 'Logging out...')
                login_prompt()  # Log out and trigger the login function
                return

            # Split the user input by '-' to check for a range
            if '-' in choice:
                start_time, end_time = choice.split('-')

                # Check if both start_time and end_time are valid
                if start_time in all_slots and end_time in all_slots:
                    selected_time_range = f"{start_time}-{end_time}"
                    print(Fore.GREEN + f"You selected: {selected_time_range}")
                    break
                else:
                    print(Fore.RED + "Invalid time range. Please enter valid time slots.")
            else:
                # Check if the single choice is valid
                if choice in all_slots:
                    print(Fore.GREEN + f"You selected: {choice}")
                    retrieve_appointment_details(selected_week, selected_date, choice)
                    break
                else:
                    print(Fore.RED + "Invalid time slot. Please enter a valid time slot.")
        except ValueError:
            print(Fore.RED + "Invalid input. Please enter a valid appointment time or range.")

def retrieve_appointment_details(selected_week, selected_date, selected_time):
    """
    Retrieve the appointment details from the selected cell in the worksheet.
    """
    # Access the worksheet for the selected week
    worksheet = SHEET.worksheet(selected_week)

    # Find the cell corresponding to the selected_date in column A (A2:A6)
    date_cells = worksheet.range("A2:A6")
    date_cell = next((cell for cell in date_cells if cell.value == selected_date), None)

    # Find the cell corresponding to the selected_time in row 1 (B1:Q1)
    time_cells = worksheet.range("B1:Q1")
    time_cell = next((cell for cell in time_cells if cell.value == selected_time), None)

    # Access the cell using row and column values and print its value
    appointment_details = worksheet.cell(date_cell.row, time_cell.col).value

    handle_appointment_action(appointment_details)

def handle_appointment_action(appointment_details):
    """
    Handle user actions based on appointment details.
    """
    while True:
        if appointment_details == "OPEN":
            print(Fore.BLUE + f"This is an {Fore.GREEN}OPEN {Fore.BLUE}slot.")
            action = input("Enter '1' to book, '2' to block the slot or '3' cancel the action \n")
            if action == "1":
                # Implement booking logic here
                print(Fore.GREEN + f"The slot is now {Fore.BLUE}BOOKED.")
                break
            elif action == "2":
                # Implement blocking logic here
                print(Fore.GREEN + f"The slot is now {Fore.RED}BLOCKED.")
                break
            elif action == "3":
                print(Fore.YELLOW + "Action cancelled.")
                break
            else:
                print(Fore.RED + "Invalid input. Please enter a valid value.")

        elif appointment_details == "BOOKED":
            print(Fore.BLUE + "This is a BOOKED slot.")
            action = input("Enter '1' to cancel the booking or '2' to cancel the action \n")
            if action == "1":
                # Implement appointment cancellation logic here
                print(Fore.GREEN + "The appointment was cancelled and is now OPEN.")
                break
            elif action == "2":
                print(Fore.YELLOW + "Action cancelled.")
                break
            else:
                print(Fore.RED + "Invalid input. Please enter a valid value.")

        elif appointment_details == "BLOCKED":
            print(Fore.BLUE + f"This is a {Fore.RED}BLOCKED {Fore.BLUE}slot.")
            action = input("Enter '1' to unblock the slot or '2' to cancel the action \n")
            if action == "1":
                # Implement slot unblocking logic here
                print(Fore.GREEN + "The slot was unblocked and is now OPEN.")
                break
            elif action == "2":
                print(Fore.YELLOW + "Action cancelled.")
                break
            else:
                print(Fore.RED + "Invalid input. Please enter a valid value.")

# login_prompt()
selected_week = "week4"
selected_date = "Friday (20-10-2023)"
selected_time = "11:30"
retrieve_appointment_details(selected_week, selected_date, selected_time)