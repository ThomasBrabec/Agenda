import os
import json
import requests

from msal import ConfidentialClientApplication
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime
from zoneinfo import ZoneInfo


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


# ============================================================
# PERIODE BEPALEN
# ============================================================

amsterdam = ZoneInfo("Europe/Amsterdam")

today = datetime.now(amsterdam)

# Eerste dag van deze maand
first_of_this_month = today.replace(
    day=1,
    hour=0,
    minute=0,
    second=0,
    microsecond=0
)

# Eerste dag van vorige maand
if first_of_this_month.month == 1:
    first_of_previous_month = first_of_this_month.replace(
        year=first_of_this_month.year - 1,
        month=12
    )
else:
    first_of_previous_month = first_of_this_month.replace(
        month=first_of_this_month.month - 1
    )


# ============================================================
# AGENDA OPHALEN
# ============================================================

url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/calendar/calendarView"

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json",
}

params = {
    "startDateTime": first_of_previous_month.isoformat(),
    "endDateTime": first_of_this_month.isoformat(),
    "$select": "subject,start,end,location",
    "$top": 100,
}


# Alle pagina's ophalen
events = []

while url:

    response = requests.get(
        url,
        headers=headers,
        params=params,
    )

    response.raise_for_status()

    data = response.json()

    events.extend(data.get("value", []))

    # Volgende pagina
    url = data.get("@odata.nextLink")

    # Parameters alleen bij de eerste request meesturen
    params = None


print(
    f"Periode: {first_of_previous_month.strftime('%Y-%m-%d')}"
    f" t/m "
    f"{first_of_this_month.strftime('%Y-%m-%d')}"
)

print(f"Totaal aantal afspraken: {len(events)}")


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

utc = ZoneInfo("UTC")

for event in events:

    subject = event.get("subject", "")

    start = event.get("start", {}).get("dateTime", "")
    end = event.get("end", {}).get("dateTime", "")

    location = event.get("location", {}).get("displayName", "")

    # Alleen afspraken met een locatie meenemen
    if not location:
        continue

    # UTC omzetten naar Nederlandse tijd
    start = datetime.fromisoformat(start).replace(
        tzinfo=utc
    ).astimezone(amsterdam)

    end = datetime.fromisoformat(end).replace(
        tzinfo=utc
    ).astimezone(amsterdam)

    # Datum en tijd
    start_date = start.strftime("%Y-%m-%d")
    start_time = start.strftime("%H:%M")
    end_time = end.strftime("%H:%M")

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
