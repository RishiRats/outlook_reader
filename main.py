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

# Setting Period
print('***************************************')
DD = int(input('Input start DD: '))
MM = int(input('Input start MM: '))
YYYY = int(input('Input start YYYY: '))
startdate = datetime.datetime(YYYY, MM, DD, 0, 0, 0)

nDD = int(input('\nInput end DD: '))
nMM = int(input('Input end MM: '))
nYYYY = int(input('Input end YYYY: '))
enddate = datetime.datetime(nYYYY, nMM, nDD, 0, 0, 0)
print('***************************************')

# mainstuff
with MailBox('outlook.office365.com').login( username, password) as mailbox:
    for msg in mailbox.fetch():
        if startdate<(msg.date).replace(tzinfo=None)<enddate:
            if (msg.flags[0]) == ('\Seen'):
                seen = seen + 1
                try:
                    if (msg.flags[1]) == ('\Flagged'):
                        flags = flags + 1
                except:
                    print(' ')
        else:
            print('Please choose correct period')
print('seen: {}, flagged: {}'.format(seen, flags))
