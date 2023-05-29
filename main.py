# A Python script to pick up follow-up canvasser emails from your inbox and put them
# into a Google Sheet for logging. Setup: enable Gmail API; authorise via OAuth 2.0 as
# Desktop app

from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import base64
import re
from datetime import datetime, timedelta

GMAIL_QUERY = (
    "subject:'Labour Doorstep - Follow-up actions from your canvassing session'"
)
SPREADSHEET_ID = "1MIzNcxvvwLWGuRYtWMmbNyPqmG3oGblUlowfR16m6J8"

# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]


def auth(creds):
    """Authorises"""
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        return creds
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())
            return creds


def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    creds = auth(creds)
    # Call the Gmail API
    mail_service = build("gmail", "v1", credentials=creds)

    # Find all messages with subject "Labour Doorstep - Follow-up actions from your
    # canvassing session"
    message_response = (
        mail_service.users()
        .messages()
        .list(
            userId="me",
            q=GMAIL_QUERY,
            maxResults=500,
        )
        .execute()
    )
    email_messages = message_response.get("messages")
    email_messages.reverse()

    # Find last action date
    sheet_service = build("sheets", "v4", credentials=creds)
    # Call the Sheets API
    sheet = sheet_service.spreadsheets()
    result = (
        sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="ACTIONS!A:D").execute()
    )
    values = result.get("values", [])
    last_action_date = datetime.strptime((values[len(values) - 1])[0], "%Y-%m-%d")

    # Loop through each email and extract the relevant variables
    for email in email_messages:
        email = (
            mail_service.users()
            .messages()
            .get(userId="me", id=email["id"], format="full")
        ).execute()
        headers = email["payload"]["headers"]
        for header in headers:
            if header["name"] == "Date":
                date = header["value"][:-15]
                date = datetime.strptime(date, "%a, %d %B %Y")
                date_format = date.strftime("%Y-%m-%d")
        body = email["payload"]["parts"][0]["body"]["data"].strip()
        body_decoded = str(base64.urlsafe_b64decode(body))

        # Find an action
        for action in re.finditer(r"\<li\>(.+?)\ \-\ (.+?)\<\/li\>", body_decoded):
            name = action.group(1)
            address = action.group(2)

            # Check if the action date is later than last_action_date, accounting for
            # 2-day lag
            if date > last_action_date + timedelta(days=2):
                # Function to write date, name and address into the Google sheet
                write_to_sheet(creds, date, name, address)
            else:
                print(f"Old email from {date_format} skipped.")


def write_to_sheet(creds, date, name, address):
    sheet_service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = sheet_service.spreadsheets()

    # Set date, accounting for 2 day lag
    date = (date - timedelta(days=2)).strftime("%Y-%m-%d")
    values = [
        [date, name, address],
    ]
    body = {"values": values}
    result = (
        sheet.values()
        .append(
            spreadsheetId=SPREADSHEET_ID,
            range="ACTIONS!A:D",
            body=body,
            valueInputOption="USER_ENTERED",
        )
        .execute()
    )
    print(f"{(result.get('updates').get('updatedCells'))} cells appended.")
    return result


if __name__ == "__main__":
    main()
