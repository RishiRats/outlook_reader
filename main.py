# importing libraries
from imap_tools import MailBox, AND

# taking user's credentials

#username = input('Username: ')
#password = input('\nPassword: ')
username = 'testuser_elpida@outlook.com'
password = '_elpidatestuser@testing'

seen = int(0)
flags = int(0)

# mainstuff
with MailBox('outlook.office365.com').login( username, password) as mailbox:
    for msg in mailbox.fetch():
        if (msg.flags[0]) == ('\Seen'):
            seen = seen + 1
        else:
            flags = flags + 1
print('seen: {}, flags: {}'.format(seen, flags))
