from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import json
import requests
import smtplib
from email.message import EmailMessage
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from requests_ntlm import HttpNtlmAuth
import openpyxl
import io
import re
import pandas as pd
from datetime import datetime
import time
import os
from urllib.parse import unquote, urlparse


def create_case(go_api_url, SagsTitel, SagsID, session):
    '''
    Function for creating case in GetOrganized for the applicant to access
    '''
    url = f"{go_api_url}/geosager/_goapi/Cases"

    payload = json.dumps({
    "CaseTypePrefix": "GEO",
    "MetadataXml": f"<z:row xmlns:z=\"#RowsetSchema\" ows_Title=\"Journaliseret {SagsTitel}\" ows_CaseStatus=\"Åben\" ows_EksterntSagsID=\"{SagsID}\" ows_EksterntSystemID=\"TestSystemID\" />",
    "ReturnWhenCaseFullyCreated": True
    })
    headers = {
    'Content-Type': 'application/json'
    }

    response = session.post(url, headers=headers, data=payload)

    return response.text

def upload_document_go(go_api_url, payload, session):
    '''
    Uploades document to case in GO
    '''
    url = f"{go_api_url}/_goapi/Documents/AddToCase"
    response = session.post(url, data=payload, timeout=1200)
    response.raise_for_status()
    return response.json()

def create_session (Username, PasswordString):
    # Create a session
    session = requests.Session()
    session.auth = HttpNtlmAuth(Username, PasswordString)
    return session

def download_file(go_url, file_path, DokumentID, GoUsername, GoPassword):
    try:
        max_retries = 2
        for attempt in range(max_retries):
            try:
                # Hent metadata for at finde dokumentets URL
                metadata_url = f"{go_url}/_goapi/Documents/MetadataWithSystemFields/{DokumentID}/True"
                metadata_response = requests.get(
                    metadata_url,
                    auth=HttpNtlmAuth(GoUsername, GoPassword),
                    headers={"Content-Type": "application/json"},
                    timeout=60
                )

                content = metadata_response.text
                DocumentURL = content.split("ows_EncodedAbsUrl=")[1].split('"')[1]
                DocumentURL = DocumentURL.split("\\")[0].replace("test.go.aarhus", "testad.go.aarhus")

                # Download selve filen
                handler = requests.Session()
                handler.auth = HttpNtlmAuth(GoUsername, GoPassword)
    
                with handler.get(DocumentURL, stream=True) as download_response:
                    download_response.raise_for_status()
                    with open(file_path, "wb") as file:
                        for chunk in download_response.iter_content(chunk_size=8192):
                            file.write(chunk)

                break

            except Exception as retry_exception:
                print(f"Retry {attempt + 1} failed: {retry_exception}")
                if attempt == max_retries - 1:
                    raise RuntimeError(
                        f"Failed to download file after {max_retries} retries. "
                        f"DokumentID: {DokumentID}, Path: {file_path}"
                    )
                time.sleep(5)

    except RuntimeError as nested_exception:
        print(f"An unrecoverable error occurred: {nested_exception}")
        raise nested_exception
def delete_case_go(go_api_url, session, sagsnummer):
    '''
    Deletes case in go
    '''
    url = f"{go_api_url}/geosager/_goapi/Cases/{sagsnummer}"
    response = session.delete(url, data= {"Data": ""}, timeout=1200)
    response.raise_for_status()
    return response.json()

def delete_local_file(filsti):
    """
    Sletter en lokal fil ud fra stien.
    Returnerer True hvis slettet, False hvis filen ikke fandtes.
    """
    try:
        os.remove(filsti)
    except FileNotFoundError:
        print(f"Filen findes ikke: {filsti}")
    except Exception as e:
        print(f"Fejl ved sletning af {filsti}: {e}")

def make_payload_document(ows_dict: dict, caseID: str, FolderPath: str, byte_arr: list, filename):
    ows_str = ' '.join([f'ows_{k}="{v}"' for k, v in ows_dict.items()])
    MetaDataXML = f'<z:row xmlns:z="#RowsetSchema" {ows_str}/>'

    return {
        "Bytes": byte_arr,
        "CaseId": caseID,
        "ListName": "Dokumenter",
        "FolderPath": FolderPath.replace("\\","/"),
        "FileName": filename,
        "Metadata": MetaDataXML,
        "Overwrite": True
    }

def get_case_metadata(gourl, sagsnummer, session):
    url = f"{gourl}/_goapi/Cases/Metadata/{sagsnummer}"

    session.headers.update({"Content-Type": "application/json"})

    response = session.get( url)

    return response.text

import json
import re

def get_case_documents(session, GOAPI_URL, SagsURL, SagsID):
    print(f'Goapiurl {GOAPI_URL}, Sagsurl {SagsURL}, SagsID {SagsID}')

    Akt = SagsURL.split("/")[1]
    encoded_sags_id = SagsID.replace("-", "%2D")
    ListURL = f"%27%2Fcases%2F{Akt}%2F{encoded_sags_id}%2FDokumenter%27"

    ViewId = None
    ikke_journaliseret_id = None
    journaliseret_id = None
    view_ids_to_use = []
    all_rows = []

    response = session.get(f"{GOAPI_URL}/{SagsURL}/_goapi/Administration/GetLeftMenuCounter")
    response.raise_for_status()
    ViewsIDArray = json.loads(response.text)

    for item in ViewsIDArray:
        if item["ViewName"] == "UdenMapper.aspx":
            ViewId = item["ViewId"]
            break

        elif item["ViewName"] == "Ikkejournaliseret.aspx":
            ikke_journaliseret_id = item["ViewId"]
            if ikke_journaliseret_id is None:
                print('None detecteret')
                LinkURL = item["LinkUrl"]
                response = session.get(f'{GOAPI_URL}{LinkURL}')
                response.raise_for_status()

                match = re.search(r'_spPageContextInfo\s*=\s*({.*?});', response.text, re.DOTALL)
                if not match:
                    raise ValueError("Kunne ikke finde _spPageContextInfo i HTML")

                context_info = json.loads(match.group(1))
                ikke_journaliseret_id = context_info.get("viewId", "").strip("{}")

        elif item["ViewName"] == "Journaliseret.aspx":
            journaliseret_id = item["ViewId"]
            if journaliseret_id is None:
                print('None detecteret')
                LinkURL = item["LinkUrl"]
                response = session.get(f'{GOAPI_URL}{LinkURL}')
                response.raise_for_status()

                match = re.search(r'_spPageContextInfo\s*=\s*({.*?});', response.text, re.DOTALL)
                if not match:
                    raise ValueError("Kunne ikke finde _spPageContextInfo i HTML")

                context_info = json.loads(match.group(1))
                journaliseret_id = context_info.get("viewId", "").strip("{}")

    if ViewId is None:
        view_ids_to_use = [vid for vid in [ikke_journaliseret_id, journaliseret_id] if vid]

    views = [ViewId] if ViewId else view_ids_to_use

    if not views:
        raise ValueError("Ingen gyldige ViewId fundet")

    for current_view_id in views:
        firstrun = True
        MorePages = True
        NextHref = None

        while MorePages:
            url = f"{GOAPI_URL}/{SagsURL}/_api/web/GetList(@listUrl)/RenderListDataAsStream"

            if firstrun:
                full_url = f"{url}?@listUrl={ListURL}&View={current_view_id}"
            else:
                full_url = f"{url}?@listUrl={ListURL}{NextHref.replace('?', '&')}"

            response = session.post(full_url, timeout=500)
            response.raise_for_status()

            dokumentliste_json = response.json()
            dokumentliste_rows = dokumentliste_json.get("Row", [])
            all_rows.extend(dokumentliste_rows)

            NextHref = dokumentliste_json.get("NextHref")
            MorePages = bool(NextHref)
            firstrun = False

    return all_rows