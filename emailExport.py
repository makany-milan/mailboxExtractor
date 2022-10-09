# Code to extract data from a mailbox, including sender, recipient
# date, subject, content and attachments.
# The folder has to be specified and the senders can be filtered.

# Current version: 2.0
# Hopefully corrected most issues which were in the previous version,
# including decoding errors and messages & dates being shown incorrectly.

import imaplib
import mimetypes
from operator import length_hint

import pandas as pd
from html2text import html2text
import re

from sys import exit
import os
from tqdm import tqdm
from pathlib import Path

import email
from email.policy import default
from email.header import decode_header, make_header

# Change these settings in the email_settings.py file.
from emailSettings import SERVER, PORT, EMAIL_ADDRESS, PASSWORD, MAILBOXES, \
    EXPORT_LOCATION
DUPLICATE_DATA = []


def login(user, password, server, port):
    mailb = imaplib.IMAP4_SSL(server, port)
    mailb.login(user, password)
    return mailb


def checkDuplicate(new_entry):
    if new_entry not in DUPLICATE_DATA:
        DUPLICATE_DATA.append(new_entry)
        return False
    else:
        return True


def exportData(df: pd.DataFrame):
    options = {}
    options['strings_to_formulas'] = False
    options['strings_to_urls'] = False

    export_loc = (EXPORT_LOCATION / 'emailData.xlsx')
    
    with pd.ExcelWriter(export_loc, options=options, engine='xlsxwriter') as w:
        df.to_excel(w, sheet_name='Data', index=False)


def getEmail(id):
    status, maildata = mailbox.uid('fetch', bytes(id), '(RFC822)')
    if status == 'OK':
        return maildata
    else:
        return 'Error'


def clearString(s: str):
    ret = re.sub('\r|\t', '', s)
    return ret


def extractText(ePart, charSet):
    if charSet is not None:
        try:
            message = str(ePart.get_payload(decode=True), str(charSet),
                          'ignore')
            message = clearString(message)
        except Exception as e:
            print(e)
    if charSet is None:
        message = ePart.get_payload(decode=True)
        message = message.decode(errors='replace')
        message = clearString(message)

    return message


def extractHTML(ePart, loc):
    html = ePart.get_payload(decode=True)
    exportHTML = (FOLDER_LOCATION / 'raw' / f'{str(loc)}.html').resolve()

    with open(exportHTML, 'wb') as f:
        f.write(html)

    return str(exportHTML), html


def extractPDF(ePart, attachmentID, loc):
    export = (FOLDER_LOCATION / 'attachments' / f'{str(loc)}_{str(attachmentID)}.pdf').resolve()
    content = ePart.get_payload(decode=True)
    with open(export, 'wb') as f:
        f.write(content)
    return export


def extractWord(ePart, attachmentID, loc):
    export = (FOLDER_LOCATION / 'attachments' / f'{str(loc)}_{str(attachmentID)}.docx').resolve()
    content = ePart.get_payload(decode=True)
    with open(export, 'wb') as f:
        f.write(content)
    return export


def exportUnknown(ePart, attachmentID, loc, ftype):
    export = (FOLDER_LOCATION / 'attachments' / f'{str(loc)}_{str(attachmentID)}{ftype}').resolve()
    content = ePart.get_payload(decode=True)
    with open(export, 'wb') as f:
        f.write(content)
    return export


def extractParts(message: email.message.Message, loc):
    retText = ''
    textExtracted = []
    htmlLoc = ''
    attachments = ''

    attachmentID = 1
    if message.is_multipart():
        attachments = []
        for ePart in message.walk():
            contentType = ePart.get_content_type()
            charSet = ePart.get_content_charset()
            if contentType == 'text/plain':
                text = extractText(ePart, charSet)
                textExtracted.append(text)
            elif contentType == 'text/html':
                htmlLoc, htmlpayload = extractHTML(ePart, loc)
            elif contentType == 'application/pdf':
                pdfLoc = extractPDF(ePart, attachmentID, loc)
                attachmentID += 1
                attachments.append(str(pdfLoc))
            elif contentType == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                wordLoc = extractWord(ePart, attachmentID, loc)
                attachments.append(str(wordLoc))
            else:
                ftype = mimetypes.guess_extension(ePart.get_content_type())
                if not ftype:
                    continue
                try:
                    fname = exportUnknown(ePart, attachmentID, loc, ftype)
                    attachmentID += 1
                    attachments.append(str(fname))
                except:
                    continue

        try:
            retText = max(textExtracted, key=len)
        except:
            retText = ''

        attachments = ','.join(attachments)

    else:
        retText = extractText(message, None)

    if retText == '':
        try:
            retText = html2text(htmlpayload.decode(errors='replace'))
        except UnboundLocalError:
            # print(f' Email {loc} has no content. Check whether it is empty.')
            pass
        except Exception as e:
            print(e)

    htmltags = ['<head>', '<body>', '<tr>', '<title>', '<html>',
    '<h1>', '<p>', '<li>', '<div>', '<table>', '<td>', '<br']
    for item in htmltags:
        if item in retText:
            try:
                retText = html2text(retText)
            except Exception as e:
                print(e)

            break
    
    retText = clearEmailJunk(retText)
    
    return retText, htmlLoc, attachments


def clearEmailJunk(message):
    #clean = re.sub('\[*\]', '', message)
    # I have not completely figured out the regular expressions
    # to get rid of all the junk in the start of the email.
    clean = message.replace('|', '').replace('#', '').replace('[', '')
    clean = clean.replace(']', '').replace('---', '')
    clean = clean.lstrip().lstrip('!')
    return clean


def extractHeaders(message):
    subject_whitespaces = message['Subject']
    subject = str(make_header(decode_header(subject_whitespaces)))
    date = message['Date']
    try:
        sender = message['From'].split('<')[1].replace('>', '')
    except IndexError:
        sender = message['From']
    try:
        to_s = message['To'].split(',')
        to_ex = []
        for t in to_s:
            to_ex.append(t.split('<')[1].replace('>', ''))
        to = ','.join(to_ex)
    except:
        to = message['To']
    try:
        td = date.split(',  ')[1].split(' ')
        dateformatted = ' '.join(td[0:3])
        timeformatted = td[3]
    except:
        try:
            td = date.split(', ')[1].split(' ')
            dateformatted = ' '.join(td[0:3])
            timeformatted = td[3]
        except:
            td = date.split(' ')
            dateformatted = ' '.join(td[1:3])
            timeformatted = td[3]
    # Remove the parentheses from the timezone.
    try:
        timezone = td[5].replace('(', '').replace(')', '')
    except:
        timezone = ''

    return sender, to, subject, dateformatted, timeformatted, timezone


def extract_main_body(text):
    if len(text) == 0:
        return text, 0
    main_body = text
    potential_delimiters = ['\nFrom:', '\n>', '\nOn', '\n\n20']
    for delimiter in potential_delimiters:
        parts = main_body.split(delimiter)
        if len(parts) > 1:
            main_body = parts[0]

    main_body = main_body.strip()

    potential_greetings = ['hi', 'dear', 'hello', 'szia', 'hope you are']
    potential_farewells =  ['best', 'regards', 'kind', 'looking forward']
    # The length cutoff determines whats the longest line the code accepts as relevant text
    length_cutoff_greetings = 20
    length_cutoff_farewells = 35

    lines = main_body.splitlines(True)
    total_lines = len(lines)

    end_reached = False

    for inx, line in enumerate(lines):
        if end_reached:
            lines.pop(inx)
            continue
        if inx < 2:
            for s in potential_greetings:
                if s in line.lower():
                    if len(line) > length_cutoff_greetings:
                        lines.pop(inx)
        if total_lines - inx < 4:
            for s in potential_farewells:
                if s in line.lower():
                    if len(line) > length_cutoff_greetings:
                        lines.pop(inx)
                        end_reached = True
    
    main_text = ''.join(lines)
    
    return main_text, len(main_text.split(None))


def fetchEmailData(folder, emails):
    # Allows the function to use mailbox and progessbar.
    global mailbox
    global pbar
    # The data will be stored in a pandas dataframe object.
    retData = pd.DataFrame(columns=['Mailbox', 'ID', 'From', 'To', 'Subject', 'Date',
                                    'Time', 'Timezone', 'Message', 'Main Body', 'Main Body Length',
                                    'HTML Location', 'Attachment Location'])
    # Unique ID number of each email in the dataframe.
    loc = 1
    for e in emails:
        # Updates the progressbar.
        pbar.update(1)
        maildata = getEmail(e)
        if maildata == 'Error':
            print('Failed fetching email: ' + str(e))
            continue
        else:
            pass

        mBinary = maildata[0][1]
        message = email.message_from_bytes(mBinary, policy=default)

        # Headers
        sender, to, subject, dateformatted, \
            timeformatted, timezone = extractHeaders(message)
        # Body
        text, htmlLoc, attachments = extractParts(message, loc)

        main_body, main_body_len = extract_main_body(text)

        # Check for duplicates
        try:
            entry = sender + to + subject + dateformatted + timeformatted
            dupe = checkDuplicate(entry)
        except:
            dupe = False
        if dupe is True:
            try:
                os.remove(htmlLoc)
            except:
                pass
        elif dupe is False:
            retData.loc[loc] = [folder, str(loc), sender, to, subject, dateformatted,
                                timeformatted, timezone, text, main_body, main_body_len,
                                htmlLoc, attachments]
            loc += 1

    return retData


if __name__ == '__main__':
    # Login to the account.
    try:
        mailbox = login(EMAIL_ADDRESS, PASSWORD, SERVER, PORT)
        print(f'Successful Login: {EMAIL_ADDRESS}')
    except Exception as e:
        print('Login Failed...')
        print(e)
        exit()

    # Create export folders
    try:
        os.mkdir(EXPORT_LOCATION)
    except FileExistsError:
        for x in range(100):
            new_export = Path(EXPORT_LOCATION.parent / f'export{x}')
            if os.path.exists(new_export):
                continue
            else:
                EXPORT_LOCATION = new_export
                os.mkdir(EXPORT_LOCATION)
                break
    
    master_data = pd.DataFrame(columns=['Mailbox', 'ID', 'From', 'To', 'Subject', 'Date',
                                    'Time', 'Timezone', 'Message',
                                    'HTML Location', 'Attachment Location'])
    '''''
    # This sections lists all the available folders.
    for i in mailbox.list()[1]:
        l = i.decode().split(' "/" ')
        print(l[0] + " = " + l[1])
    '''''
    for folder in MAILBOXES:
        # Select the mailbox.
        try:
            mailbox.select(folder)
            result, data = mailbox.uid('search', None, 'All')
            if result == 'OK':
                emails = data[0].split()
                print(f'{str(len(emails))} emails found in {folder}')
        except Exception as e:
            print('Accessing Mailbox Failed...')
            print(e)
            exit()
        
        if len(emails) > 0:
            # MAKE MAILBOX FOLDER
            clean_folder = folder.replace('[', '').replace(']', '').replace('/', '-').replace('\\', '').replace(' ', '-')
            clean_folder = clean_folder.replace('\"', '')

            FOLDER_LOCATION = (EXPORT_LOCATION / f'{clean_folder}')
            os.mkdir(FOLDER_LOCATION)

            raw_data_folder = (EXPORT_LOCATION / f'{clean_folder}/raw').resolve()
            attachments_folder = (EXPORT_LOCATION / f'{clean_folder}/attachments').resolve()

            if not os.path.exists(raw_data_folder):
                os.mkdir(raw_data_folder)

            if not os.path.exists(attachments_folder):
                os.mkdir(attachments_folder)

            # Fetch the email data
            # TQDM Provides a progress bar to easier track the process.
            pbar = tqdm(total=len(emails), desc='Fetcing Emails')
            # Threading
            # from multiprocessing.pool import ThreadPool
            #p = ThreadPool(40)    
            # data = p.map(fetchEmailData, emails)
            #p.close()
            data = fetchEmailData(clean_folder, emails)
            pbar.close()
            print(f'{len(emails)} emails successfully downloaded.')

            master_data = pd.concat([master_data, data])
        else:
            print(f'No emails found in this mailbox.')

        mailbox.unselect()

    exportData(master_data)