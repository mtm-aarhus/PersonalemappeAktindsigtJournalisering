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
import uuid
import xml.etree.ElementTree as ET
def decrypt(b64_ciphertext: str) -> str:
    EncryptionKey = os.getenv("PERSONALEINDSIGTENCRYPTIONKEY")
    if not EncryptionKey:
        return None
    combined = base64.b64decode(b64_ciphertext)
    iv = combined[:16]
    ciphertext = combined[16:]

    key = base64.b64decode(EncryptionKey) 
    cipher = Cipher(algorithms.AES(key), modes.CBC(iv), backend=default_backend())

    decryptor = cipher.decryptor()
    padded_plaintext = decryptor.update(ciphertext) + decryptor.finalize()

    unpadder = padding.PKCS7(128).unpadder()
    plaintext = unpadder.update(padded_plaintext) + unpadder.finalize()
    return plaintext.decode('utf-8')


def create_case(go_api_url, SagsTitel, SagsID, session):
    '''
    Function for creating case in GetOrganized for the applicant to access
    '''
    url = f"{go_api_url}/aktindsigt/_goapi/Cases"

    payload = json.dumps({"CaseTypePrefix": "AKT", "MetadataXml": f"<z:row xmlns:z=\"#RowsetSchema\" ows_Title=\"Journaliseret {SagsTitel}\" ows_CaseStatus=\"Åben\" />"})
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
                DocumentURL = DocumentURL.split("\\")[0].replace("go.aarhus", "ad.go.aarhus")

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

def fetch_document_info_go(api_url, DokumentID, session, AktID, Titel):
    url = f"{api_url}/_goapi/Documents/Data/{DokumentID}"
    response = session.get(url)
    data = json.loads(response.text)
    item_properties = data.get("ItemProperties", "")
    file_type_match = re.search(r'ows_File_x0020_Type="([^"]+)"', item_properties)
    version_ui_match = re.search(r'ows__UIVersionString="([^"]+)"', item_properties)
    DokumentType = file_type_match.group(1) if file_type_match else "unknown"
    VersionUI = version_ui_match.group(1) if version_ui_match else "Not found"
    file_title = f"{AktID} - {DokumentID} - {Titel}"
    return {"DokumentType": DokumentType, "VersionUI": VersionUI, "file_title": file_title}

def fetch_document_bytes(api_url, session, DokumentID, file_path=None, max_retries=30, retry_interval=5):
    url = f"{api_url}/_goapi/Documents/DocumentBytes/{DokumentID}"
    ByteResult = None
    for attempt in range(max_retries):
        try:
            response = session.get(url, timeout=180)
            response.raise_for_status()
            if b"HTTP Error 503. The service is unavailable." in response.content:
                print(f"Attempt {attempt + 1}: 503 fejl")
                time.sleep(retry_interval)
                continue
            ByteResult = response.content
            break
        except Exception as e:
            print(f"Attempt {attempt + 1}: {e}")
            time.sleep(retry_interval)
    if file_path and ByteResult:
        with open(file_path, "wb") as f:
            f.write(ByteResult)
    return ByteResult

def GOPDFConvert(api_url, DokumentID, VersionUI, GoUsername, GoPassword):
    try:
        url = f"{api_url}/_goapi/Documents/ConvertToPDF/{DokumentID}/{VersionUI}"
        response = requests.get(
            url,
            auth=HttpNtlmAuth(GoUsername, GoPassword),
            headers={"Content-Type": "application/json"},
            timeout=None
        )
        if "Document could not be converted" in response.text:
            return None
        return response.content
    except Exception:
        return None

def try_convert_go_file_to_pdf(api_url, DokumentID, session, GoUsername, GoPassword, GOUrl, file_path, orchestrator_connection=None):
    metadata = fetch_document_info_go(api_url, DokumentID, session, 0, "temp")
    VersionUI = metadata["VersionUI"]
    DokumentType = metadata["DokumentType"]
    titel = os.path.basename(file_path)

    if DokumentType.lower() == "pdf":
        if orchestrator_connection:
            orchestrator_connection.log_info(f"{DokumentID} er allerede PDF")
        byte_result = fetch_document_bytes(api_url, session, DokumentID)
        return byte_result, True, None

    result = GOPDFConvert(api_url, DokumentID, VersionUI, GoUsername, GoPassword)
    if result:
        if orchestrator_connection:
            orchestrator_connection.log_info(f"{DokumentID} konverteret via GO")
        return result, True, None

    if orchestrator_connection:
        orchestrator_connection.log_info(f"{DokumentID} GO-konvertering fejlede, forsøger fetch_document_bytes")
    byte_result = fetch_document_bytes(api_url, session, DokumentID, file_path=file_path)
    if byte_result:
        return byte_result, False, titel

    if orchestrator_connection:
        orchestrator_connection.log_info(f"{DokumentID} fetch_document_bytes fejlede, forsøger metadata-URL")
    try:
        download_file(go_url=GOUrl, file_path=file_path, DokumentID=DokumentID, GoUsername=GoUsername, GoPassword=GoPassword)
        with open(file_path, "rb") as f:
            byte_result = f.read()
        return byte_result, False, titel
    except Exception as e:
        if orchestrator_connection:
            orchestrator_connection.log_info(f"{DokumentID} alle download-metoder fejlede: {e}")
        return None, False, titel

def chunk_uploaded(offset, total_size, orchestrator_connection: OrchestratorConnection):
    orchestrator_connection.log_info(f"Uploaded {offset} out of {total_size} bytes")

def chunked_file_upload(APIURL, case_url, binary, file_name, session, request_digest, folder_path, orchestrator_connection: OrchestratorConnection):
    chunk_size_bytes = 1024 * 10240
    session.headers.update({
        'X-FORMS_BASED_AUTH_ACCEPTED': 'f',
        'X-RequestDigest': request_digest
    })

    web_url = APIURL + "/" + case_url
    if folder_path:
        target_folder_url = f"/{case_url}/Dokumenter/{folder_path}".replace("\\", "/")
    else:
        target_folder_url = f"/{case_url}/Dokumenter"

    create_file_request_url = f"{web_url}/_api/web/GetFolderByServerRelativePath(DecodedUrl=@p)/Files/add(url=@f,overwrite=true)?@p='{target_folder_url}'&@f='{file_name}'"
    response = session.post(create_file_request_url)
    response.raise_for_status()

    target_url = f"{target_folder_url}%2F{file_name}"
    upload_id = str(uuid.uuid4())
    offset = 0
    total_size = len(binary)

    with io.BytesIO(binary) as input_stream:
        first_chunk = True
        while True:
            buffer = input_stream.read(chunk_size_bytes)
            if not buffer:
                break

            if first_chunk and len(buffer) == total_size:
                endpoint_url = f"{web_url}/_api/web/GetFileByServerRelativePath(DecodedUrl=@u)/startUpload(uploadId=guid'{upload_id}')?@u='{target_url}'"
                session.post(endpoint_url, data=buffer).raise_for_status()
                endpoint_url = f"{web_url}/_api/web/GetFileByServerRelativePath(DecodedUrl=@u)/finishUpload(uploadId=guid'{upload_id}',fileOffset={offset})?@u='{target_url}'"
                session.post(endpoint_url, data=buffer).raise_for_status()
                break
            elif first_chunk:
                endpoint_url = f"{web_url}/_api/web/GetFileByServerRelativePath(DecodedUrl=@u)/startUpload(uploadId=guid'{upload_id}')?@u='{target_url}'"
                session.post(endpoint_url, data=buffer).raise_for_status()
                first_chunk = False
            elif input_stream.tell() == total_size:
                endpoint_url = f"{web_url}/_api/web/GetFileByServerRelativePath(DecodedUrl=@u)/finishUpload(uploadId=guid'{upload_id}',fileOffset={offset})?@u='{target_url}'"
                session.post(endpoint_url, data=buffer).raise_for_status()
            else:
                endpoint_url = f"{web_url}/_api/web/GetFileByServerRelativePath(DecodedUrl=@u)/continueUpload(uploadId=guid'{upload_id}',fileOffset={offset})?@u='{target_url}'"
                session.post(endpoint_url, data=buffer).raise_for_status()

            offset += len(buffer)
            chunk_uploaded(offset, total_size, orchestrator_connection)

def request_form_digest(APIURL, case_url, session):
    endpoint_url = f"{APIURL}/{case_url}/_api/contextinfo"
    session.headers.update({'Accept': 'application/json; odata=verbose'})
    response = session.post(endpoint_url)
    response.raise_for_status()
    return response.json()['d']['GetContextWebInformation']['FormDigestValue']

def get_docid(file_name, APIURL, case_url, folder_path, session, orchestrator_connection: OrchestratorConnection):
    orchestrator_connection.log_info(f'Fetching doc_id for {file_name}')
    response = session.get(f'{APIURL}/{case_url}/_goapi/Administration/GetLeftMenuCounter')
    response.raise_for_status()
    data = response.json()

    ViewId = None
    for item in data:
        if item.get("ViewName") == "AllItems.aspx" and item.get("ListName") == "Dokumenter":
            ViewId = item.get("ViewId")
            break

    if ViewId is None:
        raise ValueError("ViewId for AllItems.aspx not found.")

    list_url = f"'/{case_url}/Dokumenter'"
    root_folder = f"/{case_url}/Dokumenter/{folder_path.replace(chr(39), chr(39)*2)}" if folder_path else f"/{case_url}/Dokumenter"
    headers = {'content-type': 'application/json;odata=verbose'}
    url = f"{APIURL}/{case_url}/_api/web/GetList(@listUrl)/RenderListDataAsStream?@listUrl={list_url}&View={ViewId}&RootFolder={root_folder}"

    while True:
        payload = json.dumps({"parameters": {"__metadata": {"type": "SP.RenderListDataParameters"}, "ViewXml": "<View><RowLimit Paged=\"TRUE\">100</RowLimit></View>"}})
        response = session.post(url, headers=headers, data=payload)
        response.raise_for_status()
        data = response.json()

        for row in data.get('Row', []):
            if str(row.get('FileLeafRef')).lower() == str(file_name).lower():
                orchestrator_connection.log_info(f'DocID: {row.get("DocID")}')
                return row.get('DocID')

        next_href = data.get('NextHref')
        if next_href:
            url = f"{APIURL}/{case_url}/_api/web/GetList(@listUrl)/RenderListDataAsStream?@listUrl={list_url}{next_href.replace('?', '&')}"
        else:
            orchestrator_connection.log_info("DocID not found.")
            return None

def get_case_type(APIURL, session, case_id):
    response = session.get(f"{APIURL}/_goapi/Cases/Metadata/{case_id}/False")
    metadata = response.json()["Metadata"]
    root = ET.fromstring(metadata)
    return root.attrib.get('ows_CaseUrl')

def update_metadata(APIURL, docid, session, metadata, orchestrator_connection: OrchestratorConnection):
    start_index = metadata.find('ows_Dato="') + len('ows_Dato="')
    end_index = metadata.find('"', start_index)
    date_str = metadata[start_index:end_index]
    day, month, year = date_str.split('-')
    flipped_date = f'{month}-{day}-{year}'
    metadata = metadata.replace(date_str, flipped_date)
    payload = {"DocId": docid, "MetadataXml": metadata}
    response = session.post(f'{APIURL}/_goapi/Documents/Metadata', data=payload, timeout=600)
    response.raise_for_status()

def upload_large_document(APIURL, payload, session, binary, orchestrator_connection: OrchestratorConnection):
    case_id = payload["CaseId"]
    folder_path = payload["FolderPath"]
    file_name = payload["FileName"]
    file_name2 = file_name
    metadata = payload["Metadata"]
    case_url = get_case_type(APIURL, session, case_id)
    request_digest = request_form_digest(APIURL, case_url, session)
    file_name = file_name.replace("'", "''")
    folder_path = folder_path.replace("'", "''")
    chunked_file_upload(APIURL, case_url, binary, file_name, session, request_digest, folder_path, orchestrator_connection)
    time.sleep(5)
    docid = get_docid(file_name2, APIURL, case_url, folder_path, session, orchestrator_connection)
    if docid is not None:
        update_metadata(APIURL, docid, session, metadata, orchestrator_connection)
        return f'{{"DocId":{docid}}}'
    else:
        return 'Failed to get DocId'