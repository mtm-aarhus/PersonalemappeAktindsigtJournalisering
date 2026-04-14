import requests
from requests_ntlm import HttpNtlmAuth
import xml.etree.ElementTree as ET
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
import json

def create_ntlm_session(username: str, password: str) -> requests.Session:
    session = requests.Session()
    session.auth = HttpNtlmAuth(username, password)
    return session

def get_site_digest(site_url: str, session: requests.Session) -> str:
    """Henter et FormDigestValue for det angivne web/scope."""
    endpoint = f"{site_url}/_api/contextinfo"
    r = session.post(endpoint, headers={"Accept": "application/json; odata=verbose"})
    r.raise_for_status()
    digest = r.json()["d"]["GetContextWebInformation"]["FormDigestValue"]
    print(f"Got digest for {site_url}: {digest[:20]}...")
    return digest

def search_sharepoint_user(root_api_url: str, session: requests.Session, digest: str, email: str):
    """Søger efter en bruger i PeoplePicker (kræver root-level digest)."""
    endpoint = f"{root_api_url}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser"

    headers = {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest
    }

    payload = {
        "queryParams": {
            "QueryString": email,
            "MaximumEntitySuggestions": 50,
            "AllowEmailAddresses": False,
            "AllowOnlyEmailAddresses": False,
            "PrincipalType": 1,
            "PrincipalSource": 15,
            "SharePointGroupID": 0
        }
    }

    r = session.post(endpoint, headers=headers, data=json.dumps(payload))
    r.raise_for_status()
    results = json.loads(r.json()["d"]["ClientPeoplePickerSearchUser"])
    for entity in results:
        entity_email = entity.get("EntityData", {}).get("Email")
        if entity_email and entity_email.lower() == email.lower():
            return entity
    return None

def update_case_field(api_url: str, session: requests.Session, digest: str, item_id: str, form_values: list):
    """Opdaterer felt(er) i sagslisten."""
    endpoint = (
        f"{api_url}/aktindsigt/_api/web/GetList(@a1)/items(@a2)/ValidateUpdateListItem()"
        f"?@a1='%2Faktindsigt%2FLists%2FCases1'&@a2='{item_id}'"
    )

    headers = {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest,
        "X-Sp-Requestresources": "listUrl=%2Faktindsigt%2FLists%2FCases1"
    }

    payload = {
        "formValues": form_values,
        "bNewDocumentUpdate": False,
        "checkInComment": None
    }

    print(f"Updating case {item_id}...")
    r = session.post(endpoint, headers=headers, data=json.dumps(payload))
    r.raise_for_status()
    return r.json()

def update_case_owner(api_url: str, username: str, password: str, case_id: str, item_id: str, email_sagsbehandler: str, email_anmoder: str):
    """Opdaterer sagens CaseOwner-felt korrekt med to forskellige digests."""
    session = create_ntlm_session(username, password)

    # 1️⃣ Root-digest (bruges til PeoplePicker)
    root_digest = get_site_digest(api_url, session)

    # 2️⃣ Aktindsigt-digest (bruges til feltopdatering)
    akt_digest = get_site_digest(f"{api_url}/aktindsigt", session)

    # Find bruger
    caseowner_entity = search_sharepoint_user(api_url, session, root_digest, email_sagsbehandler)
    supplerende_entity = search_sharepoint_user(api_url, session, root_digest, email_anmoder)
    if not caseowner_entity:
        raise ValueError(f"No SharePoint user found for {email_sagsbehandler}")
    elif not supplerende_entity:
        raise ValueError(f"No SharePoint user found for {email_anmoder}")
    form_values = [
    {
        "FieldName": "CaseOwner",
        "FieldValue": json.dumps([caseowner_entity]),
        "HasException": False,
        "ErrorMessage": None
    },
    {
        "FieldName": "SupplerendeSagsbehandlere",
        "FieldValue": json.dumps([supplerende_entity]),
        "HasException": False,
        "ErrorMessage": None
    },
    {
        "FieldName": "CaseCategory",
        "FieldValue": "Kun sagsbehandlere på sagen",
        "HasException": False,
        "ErrorMessage": None
    }
]

    # Opdater felt
    result = update_case_field(api_url, session, akt_digest, item_id, form_values)
    return result


# --- Eksekvering ---
orchestrator_connection = OrchestratorConnection("AktbobJournaliser", os.getenv("OpenOrchestratorSQL"), os.getenv("OpenOrchestratorKey"), None)
orchestrator_connection.log_trace("Running process.")

API_URL = orchestrator_connection.get_constant("GOApiURL").value
go_api_login = orchestrator_connection.get_credential("GOAktApiUser")

USERNAME = go_api_login.username
PASSWORD = go_api_login.password
CASE_ID = "AKT-2023-000534"
ITEM_ID = "1249"
EMAIL_sagsbehandler = "jadt@aarhus.dk"
email_anmoder = "balas@aarhus.dk"

update_case_owner(API_URL, USERNAME, PASSWORD, CASE_ID, ITEM_ID, EMAIL_sagsbehandler, email_anmoder)
