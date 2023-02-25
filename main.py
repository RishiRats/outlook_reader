# importing libraries
from imap_tools import MailBox, AND
import datetime
from pytz import timezone

# taking user's credentials
username = input('Username: ')
password = input('password: ')

# Setting variables
seen = int(0)
flags = int(0)
totalmails = int(0)
noofmails = {}

# Setting Period
print('***************************************')
try:
    DD = int(input('Input start DD: '))
    MM = int(input('Input start MM: '))
    YYYY = int(input('Input start YYYY: '))
    startdate = datetime.datetime(YYYY, MM, DD, 0, 0, 0)

    nDD = int(input('\nInput end DD: '))
    nMM = int(input('Input end MM: '))
    nYYYY = int(input('Input end YYYY: '))
    enddate = datetime.datetime(nYYYY, nMM, nDD, 0, 0, 0)
except:
    print('Please follow the given format')
print('***************************************')

# mainstuff
with MailBox('outlook.office365.com').login( username, password) as mailbox:
    for msg in mailbox.fetch(mark_seen = False):
        test=msg.flags
        sender = (msg.from_)
        if sender in noofmails.keys():
           nomails = int((noofmails.get(sender)))
           nomails = nomails + 1
           noofmails.update({ sender : nomails })
        else:
            noofmails[sender] = int(1)

        totalmails = totalmails + 1
        if startdate<(msg.date).replace(tzinfo=None)<enddate:
            try:
                if (msg.flags[0]) == ('\Seen'):
                    seen = seen + 1
            except:
                print(' ')
            try:
                if (msg.flags[0]) == ('\Flagged'):
                    flags = flags + 1
            except:
                    print(' ')
            try:
                if (msg.flags[1]) == ('\Flagged'):
                    flags = flags + 1
            except:
                print(' ')
            try:
                if (msg.flags[2]) == ('\Flagged'):
                    flags = flags + 1
            except:
                print(' ')
            try:
                if (msg.flags[0]) == ('\Answered'):
                    ans = ans + 1
            except:
                print(' ')
            try:
                if (msg.flags[1]) == ('\Answered'):
                    ans = ans + 1
            except:
                print(' ')
        else:
            print('Please choose correct period')
print(noofmails)
print('Seen: {}, Unseen: {}, Answered: {}, Flagged: {}, Total Received Mails: {}'.format(seen, totalmails - seen, ans, flags, totalmails))
