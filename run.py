import gspread
from google.oauth2.service_account import Credentials

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('python_scheduling_system')

week1 = SHEET.worksheet('current').get_all_values()
week2 = SHEET.worksheet('1week').get_all_values()
week3 = SHEET.worksheet('2weeks').get_all_values()
week4 = SHEET.worksheet('3weeks').get_all_values()
# print(week1)
# print(week2)
# print(week3)
# print(week4)

