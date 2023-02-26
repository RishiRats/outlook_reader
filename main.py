# importing libraries
from imap_tools import MailBox
import datetime
import pyautogui as p
from xlwt import Workbook

# Setting variables
seen = 0
flags = 0
total_mail_count = 0
ans = 0
mails_dict = {}
mail_type = {'Total': 0, 'Seen': 0, 'Flagged': 0, 'Answered': 0}
MAIL_SERVER = 'outlook.office365.com'


def ask_for_date(var):
    """Function asking for date"""
    while True:
        try:
            DD = int(input(f'Enter {var} DD: '))
            MM = int(input(f'Enter {var} MM: '))
            YYYY = int(input(f'Enter {var} YYYY: '))
            date = datetime.datetime(YYYY, MM, DD, 0, 0, 0)
            return date
        except ValueError:
            print('Enter date in correct format')
            pass


# taking user's credentials
while True:
    username = 'testuser_elpida@outlook.com'  # input('Username:')
    password = p.password(text='Enter Password:', mask='*')     # Masking users password

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
start_data = ask_for_date('Start')
end_date = ask_for_date('End')

# Logins to given credentials and read
print('Fetching inbox details...')
with MailBox(MAIL_SERVER).login(username, password) as mailbox:
    # Bye default inbox isfetched
    for msg in mailbox.fetch(mark_seen=False):
        sender = msg.from_
        if sender in mails_dict.keys():
            # increments email count from one sender
            mails_dict[sender]['Total'] += 1
        else:
            mails_dict[sender] = mail_type
        total_mail_count = total_mail_count + 1

        # Shows emails only for mentioned time period
        if start_data < msg.date.replace(tzinfo=None) < end_date:
            try:
                # Checks if mail is seen, answered or flagged
                if msg.flags[0] == "\Seen":
                    mails_dict[sender]['Seen'] += 1
                elif msg.flags[0] == '\Answered' or (msg.flags[1]) == '\Answered':
                    mails_dict[sender]['Answered'] += 1
                elif msg.flags[0] == '\Flagged' or msg.flags[1] == '\Flagged' or msg.flags[2] == '\Flagged':
                    mails_dict[sender]['Flagged'] += 1
            except:
                pass

# print('{}\nSeen: {}, Unseen: {}, Answered: {}, Flagged: {}, Total Received Mails: {}'.format(mails_dict, seen,
#                                                                                              total_mail_count - seen, ans,
#                                                                                              flags, total_mail_count))

# Opens excel workbook
wb = Workbook()
SHEET_NAME = 'Email Records'
BOOK_NAME = 'Email_Records.xls'
email_record = wb.add_sheet(SHEET_NAME)
# Naming columns
email_record.write(0, 0, 'Emails')
column_name = list(mail_type.keys())
for i in range(len(column_name)):
    email_record.write(0, i+1, str(column_name[i]))
# Entering data obtained into excel
emails = list(mails_dict.keys())
for i in range(len(emails)):
    email_record.write(i+1, 0, emails[i])
    email_values = list(mails_dict[emails[i]].values())
    for j in range(len(email_values)):
        email_record.write(i+1, j+1, email_values[j])
# Saves excel sheet
wb.save(BOOK_NAME)
print('Saved details in {}'.format(BOOK_NAME))
