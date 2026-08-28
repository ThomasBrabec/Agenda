import os
import requests
from msal import ConfidentialClientApplication

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

# 3. Resultaat tonen
for event in data.get("value", []):
    print(
        event["subject"],
        "|",
        event["start"]["dateTime"],
        "-",
        event["end"]["dateTime"],
    )
