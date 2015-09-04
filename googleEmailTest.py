__author__ = 'brian'

import httplib2
import os
import re

# ----------------------------------------------------------------------------------------------------------------
# HACK follows:
# If you don't set the 3rd party path ahead of Mac OS X's path, it will get the following error:
#      parts = urllib.parse.urlparse(uri)
#      AttributeError: 'Module_six_moves_urllib_parse' object has no attribute 'urlparse'
#
# Details: see comment at bottom of page here: http://stackoverflow.com/questions/29190604/attribute-error-trying-to-run-gmail-api-quickstart-in-python
#
# Force the loading of the 3rd party libs first...
import sys
sys.path.insert(1, '/Library/Python/2.7/site-packages')
# ----------------------------------------------------------------------------------------------------------------

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import gspread

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

from gmaillib import CreateMessage
from gmaillib import CreateDraft
from gmaillib import SendMessage

# Gmail scopes: https://developers.google.com/gmail/api/auth/scopes
# Hint: Delete the ~/.credentials/*.json file if you change the scope, then the user will be asked again and the
# credentials file re-created with the new scope(s)
#SCOPES = 'https://www.googleapis.com/auth/drive.metadata.readonly'
SCOPES = 'https://www.googleapis.com/auth/drive.file \
         https://spreadsheets.google.com/feeds \
         https://docs.google.com/feeds \
         https://www.googleapis.com/auth/gmail.send \
         https://www.googleapis.com/auth/gmail.compose'

# bsdrummond@gmail.com
#CLIENT_SECRET_FILE = 'client_secret_981612606649-03ptve9ed27jj4h6k1ume8beqg6lm1bs.apps.googleusercontent.com.json'

# bdrummond@linkedin.com
# CLIENT_SECRET_FILE = 'client_secret_302402880436-pqb9hvba1459g8mghnddqoklj5uq48pn.apps.googleusercontent.com.json'
CLIENT_SECRET_FILE = 'client_secret_981612606649-quhvrk6rfnofct1h098h12nceoq10pt2.apps.googleusercontent.com.json'

#APPLICATION_NAME = 'Drive API Quickstart'
APPLICATION_NAME = 'Google Gmail API Tester'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'googleGMailAPITest.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatability with Python 2.6
            credentials = tools.run(flow, store)
        print 'Storing credentials to ' + credential_path
    return credentials

def main():
    """Shows basic usage of the Google Drive API.

    Creates a Google Drive API service object and outputs the names and IDs
    for up to 10 files.
    """

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)

    # Part 2: create a new spreadsheet
    gc = gspread.authorize(credentials)


    # TEST:
    # Open & access a spreadsheet/worksheet containing email addresses
    sheetName = "Test Data Sheet"
    worksheetName = "email test"
    try:
        ss = gc.open(sheetName)
        ws = ss.worksheet(worksheetName)
    except:
        print "Could not Open" + sheetName

    rowheader_of_projects_table = 3
    records = ws.get_all_records(empty2zero=False, head=rowheader_of_projects_table)

    for project in records:
        print "Entered By:", project['Email']

        #https://developers.google.com/gmail/api/guides/drafts
        message_text = "Status: " + project['Status']
        user_id = project['Email']
        sender = "bsdrummond@gmail.com"
        to = project['Email']
        subject = "googleEmailTest"

        message_body = CreateMessage(sender, to, subject, message_text)
        #draft = CreateDraft(service, user_id, message_body)

        # https://developers.google.com/gmail/api/guides/sending
        SendMessage(service, sender, message_body)

        True

    True

if __name__ == '__main__':
    main()