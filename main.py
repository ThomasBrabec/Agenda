import os
import json
import requests

from msal import ConfidentialClientApplication
from google.oauth2 import service_account
from googleapiclient.discovery import build


# ============================================================
# MICROSOFT GRAPH
# ============================================================

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]

USER_ID = "bc0d3b0f-e615-4ae9-a4dd-f322e2506d97"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]


# 1. Inloggen bij Microsoft Entra
app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET,
)

token_result = app.acquire_token_for_client(scopes=SCOPE)

if "access_token" not in token_result:
    raise Exception(token_result)

access_token = token_result["access_token"]


# 2. Agenda ophalen
url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/calendar/events"

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json",
}

params = {
    "$select": "subject,start,end,location",
    "$top": 10,
}

response = requests.get(
    url,
    headers=headers,
    params=params,
)

response.raise_for_status()

data = response.json()


# ============================================================
# GOOGLE SHEETS
# ============================================================

# JSON credentials uit GitHub Secret halen
google_credentials = json.loads(
    os.environ["GOOGLE_CREDENTIALS"]
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets"
]

credentials = service_account.Credentials.from_service_account_info(
    google_credentials,
    scopes=SCOPES,
)

sheets = build(
    "sheets",
    "v4",
    credentials=credentials,
)


# ID van je Google Sheet
SPREADSHEET_ID = os.environ["GOOGLE_SPREADSHEET_ID"]

SHEET_RANGE = "Blad1!A:E"


# ============================================================
# AGENDA OMZETTEN NAAR RIJEN
# ============================================================

rows = [
    ["Datum", "Onderwerp", "Start", "Einde", "Locatie"]
]

for event in data.get("value", []):

    subject = event.get("subject", "")

    start = event.get("start", {}).get("dateTime", "")
    end = event.get("end", {}).get("dateTime", "")

    location = event.get("location", {}).get("displayName", "")

    # Datum en tijd splitsen
    start_date = start[:10]
    start_time = start[11:16]

    end_time = end[11:16]

    rows.append([
        start_date,
        subject,
        start_time,
        end_time,
        location,
    ])


# ============================================================
# GOOGLE SHEETS BIJWERKEN
# ============================================================

# Eerst bestaande inhoud wissen
sheets.spreadsheets().values().clear(
    spreadsheetId=SPREADSHEET_ID,
    range=SHEET_RANGE,
).execute()


# Daarna nieuwe gegevens schrijven
sheets.spreadsheets().values().update(
    spreadsheetId=SPREADSHEET_ID,
    range=SHEET_RANGE,
    valueInputOption="RAW",
    body={
        "values": rows
    },
).execute()


print("Google Sheet succesvol bijgewerkt!")

for row in rows:
    print(row)
