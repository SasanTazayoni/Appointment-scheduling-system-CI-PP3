import gspread
from google.oauth2.service_account import Credentials
import getpass
import colorama
from colorama import Fore, Back
import datetime
import time
import os
colorama.init(autoreset=True)
if os.path.exists("env.py"):
    import env

USERNAME = os.environ.get('USERNAME')
PASSWORD = os.environ.get('PASSWORD')

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
    When triggered, this function prevents the user from taking any action.
    This is a security measure.
    """
    locked_out = True

    # Disable keyboard input
    os.system('stty -echo')

    try:
        for remaining_lockout_time in range(LOCKOUT_DURATION, 0, -1):
            # Print lockout message
            print(
                f"{Fore.RED}Maximum login attempts reached. "
                "You are now locked out for "
                f"{remaining_lockout_time} second(s).", end="\r"
            )
            time.sleep(1)
    finally:
        # Enable keyboard input after lockout
        os.system('stty echo')

    login_prompt()


def login_prompt():
    """
    Ask whether the user would like to log in and present the name of the
    application. Incorrect login details causes the login to fail.
    Failing a login 3 times causes the user to be locked out temporarily as a
    security measure.
    """
    print()
    print(
        Fore.BLUE +
        "Appointment Booking System - designed to manage weekly bookings\n"
    )

    # Login
    while True:
        choice = input(
            f"Would you like to log in? {Fore.BLUE}(y/n){Fore.RESET}: \n"
        ).strip().lower()

        if choice == 'y':
            login()
            break
        elif choice == 'n':
            print()
            print(Fore.YELLOW + "Exiting the program...")
            exit()
        elif not choice:
            print(
                f"{Fore.RED}The input cannot be empty. "
                "Please enter 'y' or 'n'.\n"
            )
        else:
            print(Fore.RED + "Invalid choice. Please enter 'y' or 'n'.\n")


def login():
    """
    Ask the user for a login and a password.
    """

    # Initialize login attempts and locked out status
    login_attempts = 0
    locked_out = False

    # Check if the user is locked out and handle the lockout if necessary
    if locked_out:
        locked_out = handle_lockout()

    print()
    print(Fore.BLUE + 'Please enter the correct login details\n')

    while login_attempts < MAX_LOGIN_ATTEMPTS and not locked_out:
        login = getpass.getpass("Username: ")

        # Check if the entered username is empty
        if not login.strip():
            print(Fore.RED + "Username cannot be empty.")
            continue

        # Loop for login attempts
        while True:
            password = getpass.getpass("Password: ")

            # Check if the entered password is empty
            if not password.strip():
                print(Fore.RED + "Password cannot be empty.")
                continue

            # Check if the entered login and password match the expected values
            if login == USERNAME and password == PASSWORD:
                print()
                print(Fore.GREEN + "Login successful!")
                login_attempts = 0
                update_cell_dates()
                return
            else:
                login_attempts += 1
                # Check if maximum login attempts reached
                if login_attempts == MAX_LOGIN_ATTEMPTS:
                    print(Fore.RED + "Login failed.\n")
                    locked_out = handle_lockout()
                else:
                    print(Fore.RED + "Login failed. Please try again.")
                break


def update_cell_dates():
    """
    Update worksheet based on login date and time.
    """
    current_datetime = datetime.datetime.now()
    monday_date = get_monday_date(current_datetime)

    # Get Monday's date for the first week on the spreadsheet
    w1_worksheet_cell_value = SHEET.worksheet("week1").acell("A2").value

    # Calculate the difference in weeks between the current date and 'week1'
    difference_in_weeks = get_week_difference(
        monday_date, w1_worksheet_cell_value
    )
    print(Fore.YELLOW + "Checking worksheets...")

    # Check if the worksheets need updating based on the difference in weeks
    if difference_in_weeks == 0:
        print(Fore.BLUE + "Worksheets are up to date")
    elif difference_in_weeks > 12:
        print(Fore.BLUE + "It seems that you have been away for a long time.")
        print(Fore.YELLOW + "Renewing worksheets...")

        # Refresh and update all worksheets if the difference is > 12 weeks
        for worksheet in SHEET.worksheets():
            refresh_cells(worksheet)
            set_cell_dates(worksheet, current_datetime)

        print(Fore.GREEN + "All worksheets have been updated.")
    elif difference_in_weeks > 0:
        print(Fore.YELLOW + "Updating worksheets...")

        # Loop through each week, update data from a target week if valid
        for i in range(1, 13):
            worksheet_title = f"week{i}"
            worksheet = SHEET.worksheet(worksheet_title)
            week_number = int(worksheet_title.lstrip("week"))
            target_week_number = week_number + difference_in_weeks

            # Check if the target week number is within the valid range (1-12)
            if 1 <= target_week_number <= 12:
                target_worksheet_title = f"week{target_week_number}"
                target_worksheet = SHEET.worksheet(target_worksheet_title)
                target_data = target_worksheet.get('B2:Q6')
                worksheet.update('B2:Q6', target_data)
                set_cell_dates(worksheet, current_datetime)
            else:
                # Update cells B2:Q6 with 'OPEN'
                cell_list = worksheet.range('B2:Q6')
                for cell in cell_list:
                    cell.value = 'OPEN'
                worksheet.update_cells(cell_list)

        print(Fore.GREEN + "Worksheets have been updated.")

    # Prompt the user to select a week
    select_week()


def get_monday_date(current_datetime):
    """
    Get Monday's date of the current week.
    """
    days_since_monday = current_datetime.weekday()
    monday_date = current_datetime - datetime.timedelta(days=days_since_monday)
    return monday_date


def get_week_difference(monday_date, w1_worksheet_cell_value):
    """
    Determine the difference in weeks between the current week and the
    spreadsheet's first week.
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
    start_date = (current_datetime -
                  datetime.timedelta(days=days_since_monday) -
                  datetime.timedelta(weeks=0) +
                  datetime.timedelta(weeks=worksheet_index))
    # Generate dates for the worksheet
    dates = [start_date + datetime.timedelta(days=j) for j in range(5)]

    # Format dates as strings
    formatted_dates = [
        f"{date.strftime('%A')} ({date.strftime('%d-%m-%Y')})"
        for date in dates
        ]

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


def select_week():
    """
    Ask the user to select a week from the spreadsheet or exit.
    """
    # Generate week titles from week1 to week12
    week_titles = [f"week{i}" for i in range(1, 13)]

    print()
    print(
        f"{Fore.BLUE}Select a number between '0-12'. "
        f"{Fore.BLUE}'week1' represents the current week.\n"
    )

    while True:
        # Display week options
        for i, title in enumerate(week_titles, start=1):
            print(
                f"{Fore.BLUE}[{i}] {Fore.WHITE}{title}", end=f"{Fore.WHITE} "
            )

        # Display the exit option
        print(f"{Fore.BLUE}[0] {Fore.WHITE}Exit")
        print()

        choice = input("Please Enter your choice here: \n")

        if choice == '0':
            print(Fore.YELLOW + "Exiting the program...")
            exit()  # Exit the program

        try:
            choice = int(choice)
            if 1 <= choice <= len(week_titles):
                selected_week = week_titles[choice - 1]
                # Find the days of the week with selected week
                get_dates_from_worksheet(selected_week)
                break
            else:
                print(
                    f"{Fore.RED}Appointments > 12 weeks are not accessible. "
                    "\nPlease enter a number between '0' and '12'."
                )
        except ValueError:
            print(Fore.RED + "Invalid input. Please enter a valid number.")


def get_dates_from_worksheet(selected_week):
    """
    Get the dates from cells A2 to A6 in the selected worksheet.
    """
    print(Fore.YELLOW + 'Retrieving dates...')
    worksheet = SHEET.worksheet(selected_week)
    date_values = worksheet.range("A2:A6")
    dates = [date.value for date in date_values]
    select_day(dates, selected_week)


def select_day(dates, selected_week):
    """
    Ask the user to select a day of the week from the selected week in the
    spreadsheet.
    """
    # Print the dates and options
    start_date = dates[0]
    print(f"{Fore.GREEN}Retrieved dates for week beginning on '{start_date}':")
    for i, date in enumerate(dates, start=1):
        print(f"{Fore.BLUE}[{i}] {Fore.WHITE}{date}")

    # Ask the user to select a date or enter '0' to reselect week
    while True:
        try:
            choice = input(
                f"Enter a number from {Fore.BLUE}'1-5'{Fore.RESET} to select "
                f"a date or {Fore.BLUE}'0'{Fore.RESET} to reselect the week:\n"
                )
            choice = int(choice)

            # Check if the user wants to return to week selection
            if choice == 0:
                select_week()  # Return to week selection
                break
            # Check if the user selected a valid date
            elif 1 <= choice <= len(dates):
                selected_date = dates[choice - 1]
                display_appointment_slots(selected_date, selected_week)
                break
            else:
                print(
                    f"{Fore.RED}Invalid number. Please enter a number "
                    "within the specified range."
                    )
        except ValueError:
            print(Fore.RED + "Invalid input. Please enter a number.")


def retrieve_appointment_slots(selected_date, selected_week):
    """
    Retrieve appointment slots for the selected day in the selected week.
    """
    # Extract the day name from the selected_date
    selected_day = selected_date.split(" ")[0]

    # Map day names to row indices
    day_to_row = {
        "Monday": 2,
        "Tuesday": 3,
        "Wednesday": 4,
        "Thursday": 5,
        "Friday": 6
    }

    # Get the row index for the selected day
    row_index = day_to_row.get(selected_day)

    # Get the selected worksheet
    worksheet = SHEET.worksheet(selected_week)

    # Get appointment slots for the selected day
    slots = worksheet.range(f"B{row_index}:Q{row_index}")

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

    # Retrieve appointment slots for the selected date and week
    all_slots = retrieve_appointment_slots(selected_date, selected_week)

    # Present visual display for time slots
    print(
        f"{Fore.GREEN}Retrieved appointment slots for "
        f"{Fore.BLUE}{selected_date}:\n"
        )

    color_codes = {
        "BLOCKED": Fore.RED,
        "BOOKED": Fore.BLUE,
        "OPEN": Fore.GREEN
    }

    # Iterate through the dictionary and format the slots
    formatted_slots = " ".join(
        [f"{time_slot} {color_codes[status]}{status}{Fore.RESET}"
            for time_slot, status in all_slots.items()]
    )
    print(formatted_slots)
    print()

    select_appointment_slots(all_slots, selected_date, selected_week)


def select_appointment_slots(all_slots, selected_date, selected_week):
    """
    Prompt the user to select a single time slot or a time range.
    """
    print(
        f"{Fore.BLUE}Booked slots cannot be blocked and blocked slots "
        "cannot be booked."
    )

    while True:
        try:
            choice = input(
                "Enter the desired appointment time or range (e.g. "
                f"{Fore.BLUE}'09:00'{Fore.RESET} or "
                f"{Fore.BLUE}'10:30-11:30'{Fore.RESET}), "
                f"{Fore.BLUE}'cancel'{Fore.RESET} to return to date selection "
                f"or {Fore.BLUE}'exit'{Fore.RESET} to exit: \n"
            )

            if choice.lower() == 'cancel':
                print(Fore.YELLOW + 'Returning to previous menu...')
                # Return to date selection
                get_dates_from_worksheet(selected_week)
                return
            elif choice.lower() == 'exit':
                print(Fore.YELLOW + "Exiting the program...")
                exit()  # Exit the program
                return

            # Split the user input by '-' to check for a range
            if '-' in choice:
                start_time, end_time = choice.split('-')

                # Check if both start_time and end_time are valid
                if start_time in all_slots and end_time in all_slots:
                    # Generate time slots between start_time and end_time
                    start_time_parts = start_time.split(':')
                    end_time_parts = end_time.split(':')
                    start_hour, start_minute = (
                        int(start_time_parts[0]),
                        int(start_time_parts[1])
                    )
                    end_hour, end_minute = (
                        int(end_time_parts[0]),
                        int(end_time_parts[1])
                    )

                    # Create a list of time slots
                    selected_time_range = []
                    while (start_hour < end_hour or
                            (start_hour == end_hour and start_minute <= end_minute)):
                        selected_time_range.append(
                            f"{start_hour:02}:{start_minute:02}"
                        )
                        start_minute += 30
                        if start_minute == 60:
                            start_hour += 1
                            start_minute = 0
                    access_appointment_slots(
                        selected_date, selected_week, selected_time_range
                        )
                    break
                else:
                    print(
                        f"{Fore.RED}Invalid time range input. Please ensure "
                        "you enter times in the correct format and between "
                        "09:00 and 16:30"
                    )
            else:
                # Check if the single choice is valid
                if choice in all_slots:
                    selected_time = [choice]
                    access_appointment_slots(selected_date, selected_week, selected_time)
                    break
                else:
                    print(
                        f"{Fore.RED}Invalid time input. Please ensure you "
                        "enter times in the correct format and between 09:00 "
                        "and 16:30."
                    )
        except ValueError:
            print(
                f"{Fore.RED}Invalid input. Please enter a valid appointment "
                "time or range."
            )


def access_appointment_slots(selected_date, selected_week, selected_time):
    """
    This function accesses the values of the appointment slots so that they
    are ready for editing.
    """
    # Access the worksheet for the selected week
    worksheet = SHEET.worksheet(selected_week)

    # Find the cell corresponding to the selected_date in column A (A2:A6)
    date_cells = worksheet.range("A2:A6")
    date_cell = next((cell for cell in date_cells if cell.value == selected_date), None)

    # Access the range B1:Q1 containing time slots
    time_cells = worksheet.range("B1:Q1")

    # Initialise a list to store the cell objects matching the selected times
    selected_time_cells = []

    # Iterate through selected times and match cells to time_cells
    for selected_time_slot in selected_time:
        time_cell = next((cell for cell in time_cells if cell.value == selected_time_slot), None)
        if time_cell:
            selected_time_cells.append(time_cell)

    # Initialise a list to store appointment details
    appointment_details_list = []

    # Access coordinates of the cell
    for time_cell in selected_time_cells:
        appointment_details = worksheet.cell(date_cell.row, time_cell.col).value
        appointment_details_list.append(appointment_details)

    # Check the size of appointment_details_list
    if len(appointment_details_list) == 1:
        # If there's only one item in the list, pass it to handle_slot_action
        slot_update = handle_slot_action(appointment_details_list[0])
        # Update single appointment slot
        update_appointment_slot(
            selected_date, selected_week, slot_update, worksheet, date_cell,
            selected_time_cells
        )
    else:
        # If list items > 1, trigger a separate function
        multislot_update = handle_multislot_action(appointment_details_list)
        # Update 2 or more appointment slots
        update_multi_appointment_slots(
            selected_date, selected_week, multislot_update, worksheet,
            date_cell, selected_time_cells
        )


def handle_slot_action(appointment_details):
    """
    Handle actions for a single appointment slot change.
    """
    while True:
        if appointment_details == "OPEN":
            slot_update = handle_open_slot()
        elif appointment_details == "BOOKED":
            slot_update = handle_booked_slot()
        elif appointment_details == "BLOCKED":
            slot_update = handle_blocked_slot()

        return slot_update


def handle_open_slot():
    """
    Handle actions for an OPEN appointment slot.
    """
    print(Fore.BLUE + f"This is an {Fore.GREEN}OPEN {Fore.BLUE}slot.")

    while True:
        action = input(
            f"Enter {Fore.BLUE}'1'{Fore.RESET} to book the slot, "
            f"{Fore.BLUE}'2'{Fore.RESET} to block the slot or "
            f"{Fore.BLUE}'3'{Fore.RESET} to return to the previous menu: \n"
        )
        if action == "1":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                return "BOOKED"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "2":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                return "BLOCKED"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "3":
            print(Fore.YELLOW + "Returning to the previous menu...")
            return ''
        else:
            print(Fore.RED + "Invalid input. Please enter a valid value.")


def handle_booked_slot():
    """
    Handle actions for a BOOKED appointment slot.
    """
    print(Fore.BLUE + "This is a BOOKED appointment slot.")

    while True:
        action = input(
            f"Enter {Fore.BLUE}'1'{Fore.RESET} to cancel the slot or "
            f"{Fore.BLUE}'2'{Fore.RESET} to return to the previous menu: \n"
        )
        if action == "1":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                print(Fore.GREEN + "The appointment was cancelled.")
                return "OPEN"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "2":
            print(Fore.YELLOW + "Returning to the previous menu...")
            return ''
        else:
            print(Fore.RED + "Invalid input. Please enter a valid value.")


def handle_blocked_slot():
    """
    Handle actions for a BLOCKED appointment slot.
    """
    print(Fore.BLUE + f"This is an {Fore.RED}BLOCKED {Fore.BLUE}slot.")

    while True:
        action = input(
            f"Enter {Fore.BLUE}'1'{Fore.RESET} to unblock the slot or "
            f"{Fore.BLUE}'2'{Fore.RESET} to return to the previous menu: \n"
        )
        if action == "1":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                print(Fore.GREEN + "The slot was unblocked.")
                return "OPEN"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "2":
            print(Fore.YELLOW + "Returning to the previous menu...")
            return ''
        else:
            print(Fore.RED + "Invalid input. Please enter a valid value.")


def update_appointment_slot(
            selected_date, selected_week, slot_update, worksheet,
            date_cell, selected_time_cells):
    """
    Update the slot and then ask if the user would like to schedule more
    appointments or exit.
    """
    # Extract time cell from list
    time_cell = selected_time_cells[0]

    # Update the slot based on user input
    if slot_update == "":
        # Trigger the function again if slot_update is an empty string
        display_appointment_slots(selected_date, selected_week)
    elif slot_update == "BLOCKED":
        worksheet.update_cell(date_cell.row, time_cell.col, 'BLOCKED')
        print(Fore.GREEN + f"The slot is now {Fore.RED}BLOCKED.")
    elif slot_update == "OPEN":
        worksheet.update_cell(date_cell.row, time_cell.col, 'OPEN')
        print(Fore.GREEN + "The slot is now OPEN.")
    elif slot_update == "BOOKED":
        worksheet.update_cell(date_cell.row, time_cell.col, 'BOOKED')
        print(Fore.GREEN + f"The appointment slot is now {Fore.BLUE}BOOKED.")

        # Check the cell before and after to see if they are "OPEN" and update
        prev_time_col = time_cell.col - 1
        next_time_col = time_cell.col + 1

        prev_slot = worksheet.cell(date_cell.row, prev_time_col).value
        next_slot = worksheet.cell(date_cell.row, next_time_col).value

        # If the slot before is "OPEN," update it to "BLOCKED"
        if prev_slot == "OPEN":
            worksheet.update_cell(date_cell.row, prev_time_col, 'BLOCKED')
        # If the slot after is "OPEN," update it to "BLOCKED"
        if next_slot == "OPEN":
            worksheet.update_cell(date_cell.row, next_time_col, 'BLOCKED')

    # Check if the user wants to schedule more appointments
    if prompt_scheduling():
        # Trigger the function to display appointment slots again
        display_appointment_slots(selected_date, selected_week)


def handle_multislot_action(appointment_details_list):
    """
    Handle action for multiple appointment slot changes.
    """
    # Create sets to store unique slot states
    slot_states = set(appointment_details_list)

    # Determine the combinations of slot states
    combinations = ", ".join(sorted(slot_states))

    # Select the appropriate action based on the slots selected by the user
    if combinations == "BLOCKED, BOOKED, OPEN":
        multislot_update = handle_mixture_of_blocked_booked_open()
    elif combinations == "BLOCKED, BOOKED":
        multislot_update = handle_mixture_of_blocked_booked()
    elif combinations == "BLOCKED, OPEN":
        multislot_update = handle_mixture_of_blocked_open()
    elif combinations == "BOOKED, OPEN":
        multislot_update = handle_mixture_of_booked_open()
    elif "BLOCKED" in slot_states:
        multislot_update = handle_multiple_blocked()
    elif "BOOKED" in slot_states:
        multislot_update = handle_multiple_booked()
    elif "OPEN" in slot_states:
        multislot_update = handle_multiple_open()

    return multislot_update


def handle_mixture_of_blocked_booked_open():
    """
    Provide the user with the option to open multiple slots
    (which cancels all appointments and unblocks all blocked slots) or to
    cancel the action.
    """
    print(
        f"{Fore.BLUE}You have selected a mixture of {Fore.RED}BLOCKED, "
        f"{Fore.GREEN}OPEN {Fore.BLUE}and BOOKED slots."
    )

    while True:
        action = input(
            f"Enter {Fore.BLUE}'1'{Fore.RESET} to open all the slots "
            "(all appointments will be cancelled and all blocked slots will "
            f"be unblocked) or {Fore.BLUE}'2'{Fore.RESET} to return to the "
            "previous menu: \n"
        )
        if action == "1":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                print(
                    f"{Fore.GREEN}All appointments were cancelled and slots "
                    "were unblocked."
                )
                return "OPEN"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "2":
            print(Fore.YELLOW + "Returning to the previous menu...")
            return ''
        else:
            print(Fore.RED + "Invalid input. Please enter a valid value.")


def handle_mixture_of_blocked_booked():
    """
    Provide the user with the option to open multiple slots
    (which cancels all appointments and unblocks all blocked slots) or to
    cancel the action.
    """
    print(
        f"{Fore.BLUE}You have selected a mixture of "
        f"{Fore.RED}BLOCKED {Fore.BLUE}and BOOKED slots."
    )

    while True:
        action = input(
            f"Enter {Fore.BLUE}'1'{Fore.RESET} to open all the slots (all "
            "appointments will be cancelled and all blocked slots will be "
            f"unblocked) or {Fore.BLUE}'2'{Fore.RESET} to return to the "
            "previous menu: \n"
        )
        if action == "1":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                print(
                    f"{Fore.GREEN}All appointments were cancelled and slots "
                    "were unblocked."
                )
                return "OPEN"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "2":
            print(Fore.YELLOW + "Returning to the previous menu...")
            return ''
        else:
            print(Fore.RED + "Invalid input. Please enter a valid value.")


def handle_mixture_of_blocked_open():
    """
    Provide the user with the option to unblock or block multiple slots or to
    cancel the action.
    """
    print(
        f"{Fore.BLUE}You have selected a mixture of "
        f"{Fore.GREEN}OPEN {Fore.BLUE}and {Fore.RED}BLOCKED {Fore.BLUE}slots."
    )

    while True:
        action = input(
            f"Enter {Fore.BLUE}'1'{Fore.RESET} to unblock and open all the "
            f"slots, {Fore.BLUE}'2'{Fore.RESET} to block all the slots or "
            f"{Fore.BLUE}'3'{Fore.RESET} to return to the previous menu: \n"
        )
        if action == "1":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                return "OPEN"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "2":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                return "BLOCKED"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "3":
            print(Fore.YELLOW + "Returning to the previous menu...")
            return ''
        else:
            print(Fore.RED + "Invalid input. Please enter a valid value.")


def handle_mixture_of_booked_open():
    """
    Provide the user with the option to cancel or book multiple slots or to
    cancel the action.
    """
    print(
        f"{Fore.BLUE}You have selected a mixture of "
        f"{Fore.GREEN}OPEN {Fore.BLUE}and BOOKED slots."
    )

    while True:
        action = input(
            f"Enter {Fore.BLUE}'1'{Fore.RESET} to book all the slots, "
            f"{Fore.BLUE}'2'{Fore.RESET} cancel all appointments and open the "
            f"slots or {Fore.BLUE}'3'{Fore.RESET} to return to the previous "
            "menu: \n"
        )
        if action == "1":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                return "BOOKED"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "2":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                return "OPEN"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "3":
            print(Fore.YELLOW + "Returning to the previous menu...")
            return ''
        else:
            print(Fore.RED + "Invalid input. Please enter a valid value.")


def handle_multiple_blocked():
    """
    Provide the user with the option to unblock multiple slots or to cancel
    the action.
    """
    print(
        f"{Fore.BLUE}You have selected multiple "
        f"{Fore.RED}BLOCKED {Fore.BLUE}slots."
    )

    while True:
        action = input(
            f"Enter {Fore.BLUE}'1'{Fore.RESET} to unblock the slots or "
            f"{Fore.BLUE}'2'{Fore.RESET} to return to the previous menu: \n"
        )
        if action == "1":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                print(Fore.GREEN + "The slots were unblocked.")
                return "OPEN"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "2":
            print(Fore.YELLOW + "Returning to the previous menu...")
            return ''
        else:
            print(Fore.RED + "Invalid input. Please enter a valid value.")


def handle_multiple_booked():
    """
    Provide the user with the option to cancel multiple slots or to cancel
    the action.
    """
    print(Fore.BLUE + "You have selected multiple BOOKED slots.")

    while True:
        action = input(
            f"Enter {Fore.BLUE}'1'{Fore.RESET} to cancel the slots or "
            f"{Fore.BLUE}'2'{Fore.RESET} to return to the previous menu: \n"
        )
        if action == "1":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                print(Fore.GREEN + "The bookings were cancelled.")
                return "OPEN"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "2":
            print(Fore.YELLOW + "Returning to the previous menu...")
            return ''
        else:
            print(Fore.RED + "Invalid input. Please enter a valid value.")


def handle_multiple_open():
    """
    Provide the user with the option to either block or book multiple slots
    with an option to cancel the action.
    """
    print(
        f"{Fore.BLUE}You have selected multiple "
        f"{Fore.GREEN}OPEN {Fore.BLUE}slots."
    )

    while True:
        action = input(
            f"Enter {Fore.BLUE}'1'{Fore.RESET} to book all the slots, "
            f"{Fore.BLUE}'2'{Fore.RESET} to block all the slots or "
            f"{Fore.BLUE}'3'{Fore.RESET} to return to the previous menu: \n"
        )
        if action == "1":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                return "BOOKED"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "2":
            # Ask for confirmation
            confirmed = get_confirmation()
            if confirmed:
                print(Fore.YELLOW + "Processing request...")
                return "BLOCKED"
            else:
                print(Fore.YELLOW + "Aborting...")
                continue
        elif action == "3":
            print(Fore.YELLOW + "Returning to the previous menu...")
            return ''
        else:
            print(Fore.RED + "Invalid input. Please enter a valid value.")


def update_multi_appointment_slots(
            selected_date, selected_week, multislot_update, worksheet,
            date_cell, selected_time_cells):
    """
    Update multiple slots and then ask if the user would like to schedule more
    appointments or exit.
    """
    for time_cell in selected_time_cells:
        # Update the slot based on user input for each selected time cell
        if multislot_update == "":
            # Trigger the function again if multislot_update is an empty string
            display_appointment_slots(selected_date, selected_week)
        elif multislot_update == "BLOCKED":
            worksheet.update_cell(date_cell.row, time_cell.col, 'BLOCKED')
        elif multislot_update == "OPEN":
            worksheet.update_cell(date_cell.row, time_cell.col, 'OPEN')
        elif multislot_update == "BOOKED":
            worksheet.update_cell(date_cell.row, time_cell.col, 'BOOKED')
            prev_time_col = time_cell.col - 1
            next_time_col = time_cell.col + 1

            # Check and update slots before and after the booked slot
            prev_slot = worksheet.cell(date_cell.row, prev_time_col).value
            next_slot = worksheet.cell(date_cell.row, next_time_col).value

            # If the slot before is "OPEN," update it to "BLOCKED"
            if prev_slot == "OPEN":
                worksheet.update_cell(date_cell.row, prev_time_col, 'BLOCKED')
            # If the slot after is "OPEN," update it to "BLOCKED"
            if next_slot == "OPEN":
                worksheet.update_cell(date_cell.row, next_time_col, 'BLOCKED')

    # Print the update result message based on the multislot_update action
    if multislot_update == "BLOCKED":
        print(Fore.GREEN + f"The slots are now {Fore.RED}BLOCKED.")
    elif multislot_update == "OPEN":
        print(Fore.GREEN + "The slots are now OPEN.")
    elif multislot_update == "BOOKED":
        print(Fore.GREEN + f"The appointment slots are now {Fore.BLUE}BOOKED.")

    # Check if the user wants to schedule more appointments
    if prompt_scheduling():
        # Trigger the function to display appointment slots again
        display_appointment_slots(selected_date, selected_week)


def get_confirmation():
    """
    Ask the user for confirmation when changing a slot or multiple slots.
    """
    while True:
        confirm = input(
            "Do you wish to confirm this change? "
            f"{Fore.BLUE}(y/n){Fore.RESET}: \n"
        ).lower()
        if confirm == "y":
            return True
        elif confirm == "n":
            return False
        else:
            print(Fore.RED + "Invalid input. Please enter 'y' or 'n'.")


def prompt_scheduling():
    """
    Ask the user if they want to schedule more appointments.
    """
    while True:
        choice = input(
            "Do you want to schedule more appointments? "
            f"{Fore.BLUE}(y/n){Fore.RESET}: \n"
        ).strip().lower()
        if choice == 'y':
            return True
        elif choice == 'n':
            print(Fore.YELLOW + "Exiting the program...")
            exit()  # Exit the program
        else:
            print(Fore.RED + "Invalid choice. Please enter 'y' or 'n'.")


login_prompt()
