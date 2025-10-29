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
    "MetadataXml": f"<z:row xmlns:z=\"#RowsetSchema\" ows_Title=\"Journaliseret aktindsigtssag {SagsID} - {SagsTitel}\" ows_CaseStatus=\"Åben\" ows_EksterntSagsID=\"{SagsID}\" ows_EksterntSystemID=\"TestSystemID\" />",
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

def get_case_documents(session, GOAPI_URL, SagsURL, SagsID ):
    # Initialize variables
    Akt = SagsURL.split("/")[1]  
    encoded_sags_id = SagsID.replace("-", "%2D")
    ListURL = f"%27%2Fcases%2F{Akt}%2F{encoded_sags_id}%2FDokumenter%27"
    ViewId = None
    view_ids_to_use = []  # To handle combined views
    response = session.get(f"{GOAPI_URL}/{SagsURL}/_goapi/Administration/GetLeftMenuCounter")
    ViewsIDArray = json.loads(response.text) # Parse the JSON

    # Check for "UdenMapper.aspx"
    for item in ViewsIDArray:
        if item["ViewName"] == "UdenMapper.aspx":
            ViewId = item["ViewId"]
            break
        elif item["ViewName"] == "Ikkejournaliseret.aspx":
            ikke_journaliseret_id = item["ViewId"]    
            if ikke_journaliseret_id is None: 
                print('None detecteret')
                LinkURL = item["LinkUrl"]
                reponse = session.get(f'{GOAPI_URL}{LinkURL}')
                                
                # Find _spPageContextInfo JavaScript-objektet
                match = re.search(r'_spPageContextInfo\s*=\s*({.*?});', reponse.text, re.DOTALL)
                if not match:
                    raise ValueError("Kunne ikke finde _spPageContextInfo i HTML")

                # Pars JSON-delen
                context_info = json.loads(match.group(1))

                # Hent ViewId og fjern {} hvis tilstede
                view_id = context_info.get("viewId", "").strip("{}")

                ikke_journaliseret_id = view_id
        elif item["ViewName"] == "Journaliseret.aspx":
            journaliseret_id = item["ViewId"]
            if journaliseret_id is None:
                print('None detecteret')
                LinkURL = item["LinkUrl"]
                reponse = session.get(f'{GOAPI_URL}{LinkURL}')
                                
                # Find _spPageContextInfo JavaScript-objektet
                match = re.search(r'_spPageContextInfo\s*=\s*({.*?});', reponse.text, re.DOTALL)
                if not match:
                    raise ValueError("Kunne ikke finde _spPageContextInfo i HTML")

                # Pars JSON-delen
                context_info = json.loads(match.group(1))
                view_id = context_info.get("viewId", "").strip("{}")

                ikke_journaliseret_id = view_id


    # # If "UdenMapper.aspx" doesn't exist, combine views
    # if ViewId is None:
    #     view_ids_to_use = [ikke_journaliseret_id, journaliseret_id]
    #     print(view_ids_to_use)

    # Iterate through views
    for current_view_id in ([ViewId] if ViewId else view_ids_to_use):
        firstrun = True
        MorePages = True

        while MorePages:

            # If not the first run, fetch the next page
            if not firstrun:
                url = f"{GOAPI_URL}/{SagsURL}/_api/web/GetList(@listUrl)/RenderListDataAsStream"
                url_with_query = f"{url}?@listUrl={ListURL}{NextHref.replace('?', '&')}"

                response = session.post(url_with_query, timeout=500)
                response.raise_for_status()
                Dokumentliste = response.text  # Extract the content
            else:
                # If first run, fetch the first page for the current view
                url = f"{GOAPI_URL}/{SagsURL}/_api/web/GetList(@listUrl)/RenderListDataAsStream"
                query_params = f"?@listUrl={ListURL}&View={current_view_id}"
                full_url = url + query_params

                response = session.post(full_url, timeout=500)
                response.raise_for_status()
                Dokumentliste = response.text  # Extract the content

            # Deserialize response
            dokumentliste_json = json.loads(Dokumentliste)
            dokumentliste_rows = dokumentliste_json.get("Row", [])

            # Check for additional pages
            NextHref = dokumentliste_json.get("NextHref")
            MorePages = "NextHref" in dokumentliste_json

    return dokumentliste_rows