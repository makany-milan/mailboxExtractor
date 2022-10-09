from pathlib import Path

# List of IMAP servers and ports bellow.
# https://www.arclab.com/en/kb/email/list-of-smtp-and-imap-servers-mailserver-list.html
SERVER = 'imap.gmail.com'
PORT = 993

# PERSONAL INFORMATION
EMAIL_ADDRESS = 'makanym@gmail.com'
PASSWORD = 'mrvg nace swkk lbzr'

MAILBOXES = ['\"[Gmail]/Starred\"', '\"[Gmail]/Sent Mail\"']
# FOLDER FOR ALL EMAILS: '\"[Gmail]/All Mail\"'
# MAILBOX = 'Inbox'

EXPORT_LOCATION = Path(r'C:\Users\Milan\Downloads\mail\export')

FILTER_SENDER = False
SENDER_ORGANISATION = ''
