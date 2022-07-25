import os
import threading
from email.mime import application
from http import client

import google.auth
import slack
from apiclient import errors
from flask import Flask, Response, request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from slackeventsapi import SlackEventAdapter

# Initialize a Flask app
application = Flask(__name__)

# Initialize a Web API client
slack_event_adapter = SlackEventAdapter(
    os.environ['SIGNING_SECRET'],os.environ['SLACK_EVENTS'], application)
client = slack.WebClient(token=os.environ['SLACK_TOKEN'])
BOT_ID = client.api_call("auth.test")['user_id']

################################################################################
# SLASH COMMANDS
@application.route("/drivemate-sheet", methods=["POST"])
def create_sheet():
    data = request.form
    user_id = data.get('user_id')
    channel_id = data.get('channel_id')
    sheetTitle = data.get('text') if data.get('text') else "New Sheet"
    members = client.conversations_members(channel=channel_id)['members']
    collaboratorsEmails = []

    for member in members:
        user = client.users_info(user=member)['user']
        if not user['is_bot']:
            if user['id'] == user_id:
                creator = {
                    'userId' : user['id'],
                    'emailAddress': user['profile']['email']
                    }
            else: 
                email = user['profile']['email']
                collaboratorsEmails.append(email)

    googleSheetId = createGoogleSheet(sheetTitle)

    #Create thread to process remaining code 
    # x = threading.Thread(
    #     target=createGoogleDriveFilePermissions,
    #     args=(googleSheetId, collaboratorsEmails, "user", "reader",)
    # )
    # x.start()
    
    createGoogleDriveFilePermissions(googleSheetId, [creator['emailAddress']], "user", "writer")
    createGoogleDriveFilePermissions(googleSheetId, collaboratorsEmails, "user", "reader")
    googleSheetAccessLink = os.environ['GOOGLE_SHEETS'] + googleSheetId
    responseBlock = {
        "type": "section", 
        "text": {
            "type": "mrkdwn",
            "text": "{} <@{}> Created A New Google Sheet: <{}|{}>".format(':sparkle:', user_id, googleSheetAccessLink, sheetTitle)
            }
        }

    client.chat_postMessage(channel=channel_id, blocks=[responseBlock])
    return Response(), 200

################################################################################
#GOOGLE CLOUD API
def createGoogleSheet(title):
    """
    Creates the Sheet the user has access to.
    Load pre-authorized application service account user credentials from the environment.
    """

    creds, _ = google.auth.default()
    try:
        service = build('sheets', 'v4', credentials=creds)
        spreadsheet = {
            'properties': {
                'title': title
            }
        }
        spreadsheet = service.spreadsheets().create(body=spreadsheet,
                                                    fields='spreadsheetId') \
            .execute()
        print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
        return spreadsheet.get('spreadsheetId')
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

def createGoogleDriveFilePermissions(file_id, users, perm_type, role):
    """
    Insert a new permission.
    Args: k
    service: Drive API service instance.
    file_id: ID of the file to insert permission for.
    value: User or group e-mail address, domain name or None for 'default'
            type.
    perm_type: The value 'user', 'group', 'domain' or 'default'.
    role: The value 'owner', 'writer' or 'reader'.

    TODO: Currently setting role=writer not owner because POC is testing on 'gmail.com' 
    AKA user created emails and sharing with other user created emails. 
    During implimentation, creator will be in the same Organization in google workspace 
    so we can use 'transferOwnership' = True as paramater AND 'role' : 'owner' to transfer ownership.
    REF: https://developers.google.com/drive/api/guides/manage-sharing#transfer_file_ownership_to_another_google_workspace_account_in_the_same_organization
    Returns:
    The inserted permission if successful, None otherwise.
    """
  
    service = build('drive', 'v3')
    for user in users: 
        new_permission = {
            'type': perm_type,
            'role': role,
            'emailAddress': user
        }
        try:
            service.permissions().create(
                fileId=file_id, body=new_permission, sendNotificationEmail=False).execute()
        except errors.HttpError as error:
            print('An error occurred while creating permission: %s' % error)
    return None
    
################################################################################

if __name__ == "__main__":
    application.run(debug=True)
