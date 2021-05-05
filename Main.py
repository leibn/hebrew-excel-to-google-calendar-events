import random

from Google import NewCalendar
from SendToGoogleCalendar import send_to_google_calendar
from SpreedSheatMaker import make_work_book

print("enter name to this secondary calendar(no space's in the name, enter '0' for default ):  ")
print("can be change later in google account")
calendar_name = input()
if calendar_name == "0":
    calendar_name = "B-days" + str(random.randint(1, 100))

print("assuming you already move yor excel data file to the directory folder")
print("enter the name that the excel have(no space's in the name, enter '0' for default):  ")
excel_path = input()
if excel_path == "0":
    excel_path = "row_data.xlsx"

print("to how many years foreword you want to set thous event?")
next_n_years = int(input())
if next_n_years < 1:
    next_n_years = int(1)

work_book_name = make_work_book(excel_path, next_n_years)
print(work_book_name)

calendar_id = NewCalendar.make_new_calendar(summary=calendar_name, time_zone='Asia/Jerusalem')
send_to_google_calendar(work_book_name, calendar_id)

print("finished in great successes")
