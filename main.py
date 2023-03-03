# importing libraries
from imap_tools import MailBox, A
import datetime
import pytz
import pyautogui as p
from xlwt import Workbook
import sys
from boxing import boxing

# Setting variables
seen = 0
flags = 0
total_mail_count = 0
ans = 0
flag_types = ['Total', 'Seen', 'Flagged', 'Answered']
mails_dict = {}
# MAIL_SERVER = 'imap.gmail.com'
MAIL_SERVER = 'outlook.office365.com'
timezone = pytz.timezone("Asia/Kolkata")

# print('''*\t\t\t\tMade by AVANA\t\t\t*
# *\t\t\t\t   Email Book Keeper \t\t*''')
box_text = boxing("  MADE BY AVANA\nEmail Book-Keeper", style='double')
print(box_text, file=sys.stdout)


def ask_for_date(var):
    """Function asking for date"""
    while True:
        try:
            DD = int(input(f'Enter {var} D: '))
            MM = int(input(f'Enter {var} M: '))
            YYYY = int(input(f'Enter {var} YYYY: '))
            date = datetime.datetime(YYYY, MM, DD)

            return date
        except ValueError:
            print('Enter date in correct format')
            pass


# taking user's credentials
while True:
    username = input('Enter email: ')
    password = p.password(text='Enter Password:', mask='*')  # Masking users password

    # Checking credentials
    print('Verifying credentials. Please wait...')
    try:
        login = MailBox(MAIL_SERVER).login(username=username, password=password)
        login.logout()
        print('Verified.')
        break
    except:
        print('Enter valid credentials')

# Ask for start and end date
print('Please Input Dates in Ascending Order.')
start_date = timezone.localize(ask_for_date('Start'))
end_date = timezone.localize(ask_for_date('End'))

#for instant use only date and remove _gte
# Logins to given credentials and read
print('Fetching inbox details...')
with MailBox(MAIL_SERVER).login(username, password) as mailbox:
    # Bye default inbox is fetched and in reversed order
    for msg in mailbox.fetch(A(date_gte=start_date.date()), mark_seen=False, reverse=True):
        if start_date < msg.date < end_date:
            sender = msg.from_
            if sender in mails_dict.keys():
                # increments email count from one sender
                mails_dict[sender]['Total'] += 1
            else:
                mails_dict[sender] = {'Total': 1, 'Seen': 0, 'Flagged': 0, 'Answered': 0}
            total_mail_count = total_mail_count + 1

            mails_dict[sender]['Seen'] = msg.flags.count('\\Seen')
            mails_dict[sender]['Answered'] = msg.flags.count('\\Answered')
            mails_dict[sender]['Flagged'] = msg.flags.count('\\Flagged')

        elif msg.date < start_date:
            break

# Opens excel workbook
wb = Workbook()
SHEET_NAME = 'Email Records'
BOOK_NAME = 'Email_Records.xls'
email_record = wb.add_sheet(SHEET_NAME)
# Naming columns
email_record.write(0, 0, 'Emails')
for i in range(len(flag_types)):
    email_record.write(0, i + 1, str(flag_types[i]))
# Entering data obtained into excel
emails = list(mails_dict.keys())
for i in range(len(emails)):
    email_record.write(i + 1, 0, emails[i])
    email_values = list(mails_dict[emails[i]].values())
    for j in range(len(email_values)):
        email_record.write(i + 1, j + 1, email_values[j])
# Saves excel sheet
wb.save(BOOK_NAME)
print('Saved details in {}'.format(BOOK_NAME))
