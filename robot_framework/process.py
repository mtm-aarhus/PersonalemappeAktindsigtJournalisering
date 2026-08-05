from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
from datetime import datetime
import json
import time
from Funktioner import *
import xml.etree.ElementTree as ET
from sqlalchemy import create_engine, text
from datetime import datetime
from urllib.parse import quote_plus
from datetime import datetime
from mail_journaliser import *
from GoBrugerstyring import *
import base64
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives import padding
from cryptography.hazmat.backends import default_backend
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:


    go_ad_url = orchestrator_connection.get_constant("GOApiURL").value
    go_ad_login = orchestrator_connection.get_credential("GOAktApiUser")
    go_ad_username = go_ad_login.username
    go_ad_password = go_ad_login.password

    # SharePoint-credentials (samme opsætning som i Aktbob2-Journaliser)
    sharepoint_certification = orchestrator_connection.get_credential("SharePointCert")
    sharepoint_api = orchestrator_connection.get_credential("SharePointAPI")
    sp_tenant = sharepoint_api.username
    sp_client_id = sharepoint_api.password
    sp_thumbprint = sharepoint_certification.username
    sp_cert_path = sharepoint_certification.password

    specific_content = json.loads(queue_element.data)

    Udleveringsmappelink = specific_content.get('Udleveringsmappelink') 
    if Udleveringsmappelink:
        SagsNummer = Udleveringsmappelink.rsplit("/")[-1]
    SagsID = specific_content.get('caseid') 
    SagsTitel = specific_content.get('PersonaleSagsTitel') 
    Journaliseringsmappelink = specific_content.get('Journaliseringsmappelink')
    EmailBody = specific_content.get('EmailBody')
    MailModtager = specific_content.get("MailModtager")
    MailAfsender = specific_content.get("MailAfsender")
    Beskrivelse = specific_content.get("Beskrivelse")
    Modtagelsesdato = specific_content.get("Modtagelsesdato")
    SharepointMappelink = specific_content.get("SharepointMappelink")

    #Making go session
    session = create_session(go_ad_username, go_ad_password)
    if Journaliseringsmappelink:
            #hvis der allerede ligger en journaliseringsmappe skal den slettes for ikke at have dobbeltmapper til at ligge
        JournaliseringsmappeID = Journaliseringsmappelink.rsplit("/")[-1]
        print(f'Gammel journaliseringsmappe detekteret {JournaliseringsmappeID}')
        try:
            delete_case_go(go_ad_url, session, JournaliseringsmappeID)
            print(f'Gammel delingsmappe slettet for sag {JournaliseringsmappeID}')
        except Exception as e:
            print(f"Tried to delete old journaliseringsmappe, but failed {e}")

    #1 - definer stuff
    today_date = datetime.now().strftime("%d-%m-%Y")

    #Hent filoplysninger på færdigbehandlet go-sag
    if Udleveringsmappelink:
        SagsMetaData = get_case_metadata(go_ad_url, SagsNummer, session)
        SagsMetaData = json.loads(SagsMetaData).get("Metadata")
        xdoc = ET.fromstring(SagsMetaData)

        # Extract attributes
        RelativeSagsUrl = xdoc.attrib.get("ows_CaseUrl")
        SagsTitel = xdoc.attrib.get("ows_Title")

        #Hent info om sag, der skal journaliseres
        casefiles = get_case_documents(session, go_ad_url, SagsURL= RelativeSagsUrl, SagsID = RelativeSagsUrl.rsplit('/')[-1])

    #Lav ny sag til at journalisere ind i
    session = create_session(go_ad_username, go_ad_password)
    CreatedCase = json.loads(create_case(go_ad_url, SagsTitel, SagsID, session))
    RelativeSagsUrl = CreatedCase['CaseRelativeUrl']
    CaseID = CreatedCase['CaseID']
    AktNr = RelativeSagsUrl.split('/')[-2]
    CaseUrl_new = f'{go_ad_url}/{RelativeSagsUrl}'
    

    #Sagsbehandler sættes som det første
    mailHR = orchestrator_connection.get_constant('balas').value
    try:
        update_case_owner(go_ad_url, go_ad_username, go_ad_password, CaseID, MailAfsender, mailHR, AktNr)
    except Exception as e:
        orchestrator_connection.log_error("Kunne ikke sætte sagsbehandler")
        raise e

    ikke_konverterede_filer = []
    fejlede_uploads = []
    uploaded_doc_ids = []
    if Udleveringsmappelink:
        for item in casefiles:
            DokTitle = item.get("Title", "")
            DokID = str(item.get("DocID"))
            file_path = os.path.join(os.getcwd(), DokTitle)

            byte_result, is_pdf, ikke_konverteret = try_convert_go_file_to_pdf(
                go_ad_url, DokID, session, go_ad_username, go_ad_password, go_ad_url, file_path
            )

            if ikke_konverteret:
                ikke_konverterede_filer.append(ikke_konverteret)

            if byte_result is None:
                print(f"Kunne ikke hente fil {DokTitle} - springer over")
                fejlede_uploads.append(DokTitle)
                continue

            if is_pdf:
                file_path = f"{os.path.splitext(DokTitle)[0]}.pdf"

            filename = os.path.basename(file_path)
            byte_arr = list(byte_result)

            ows_dict = {
                "Title": filename,
                "CaseID": CaseID,
                "Beskrivelse": "Uploaded af personaleaktbob",
                "Korrespondance": "Udgående",
                "Dato": today_date,
                "CCMMustBeOnPostList": "0"
            }
            payload = make_payload_document(
                ows_dict=ows_dict, caseID=CaseID, FolderPath="", byte_arr=byte_arr, filename=filename
            )

            try:
                if (len(byte_result) / (1024 * 1024)) > 10:
                    raise Exception("Fil er større end 10 MB, forsøger chunk-upload")
                response = upload_document_go(go_ad_url, payload=payload, session=session)

                if "DocId" not in response:
                    raise Exception("No DocId i response")
                uploaded_doc_ids.append(response["DocId"])

            except Exception as e:
                print(f"Normal upload fejlede for {filename}: {e}")
                for attempt in range(1, 4):
                    try:
                        print(f"Chunk-upload forsøg {attempt} for {filename}")
                        large_response = upload_large_document(
                            go_ad_url, payload, session, byte_result, orchestrator_connection
                        )
                        large_response_json = json.loads(large_response)
                        if "DocId" not in large_response_json:
                            raise Exception(f"Ingen DocId i chunk-response for {filename}")
                        uploaded_doc_ids.append(large_response_json["DocId"])
                        break
                    except Exception as retry_exception:
                        print(f"Chunk-upload forsøg {attempt} fejlede: {retry_exception}")
                        if attempt == 3:
                            print(f"Alle upload-metoder fejlede for {filename}")
                            fejlede_uploads.append(filename)

            delete_local_file(filsti=file_path)
    if Beskrivelse:
        application_pdf_path = save_application_pdf("Anmodning om aktindsigt", MailAfsender, Beskrivelse, Modtagelsesdato)
        with open(application_pdf_path, "rb") as local_file:
                file_content_application = local_file.read()
                byte_arr_mail = list(file_content_application)
        ows_dict_mail = {
                        "Title": "Anmodning om aktindsigt.pdf",
                        "CaseID": CaseID,  # Replace with your case ID
                        "Beskrivelse": "Uploaded af personaleaktbob",  # Add relevant description
                        "Korrespondance": "Udgående",
                        "Dato": today_date,
                        "CCMMustBeOnPostList": "0"
                        }
        payload_mail = make_payload_document(ows_dict= ows_dict_mail, caseID= CaseID, FolderPath= "", byte_arr= byte_arr_mail, filename= "Anmodning.pdf")
        response_besk = upload_document_go(go_ad_url, payload = payload_mail, session = session)
        if "DocId" in response_besk:
            uploaded_doc_ids.append(response_besk["DocId"])
        else:
            raise Exception("No docid in response_besk")
        delete_local_file(filsti = application_pdf_path)


    #Journalising the answer
    if EmailBody:
        sent_mail_pdf_path = save_communication_pdf("Vedr. din anmodning om aktindsigt", MailModtager, MailAfsender, EmailBody)
        with open(sent_mail_pdf_path, "rb") as local_file:
                file_content_mail = local_file.read()
                byte_arr_mail = list(file_content_mail)
        ows_dict_mail = {
                        "Title": "Vedr. din anmodning om aktindsigt.pdf",
                        "CaseID": CaseID,  # Replace with your case ID
                        "Beskrivelse": "Uploaded af personaleaktbob",  # Add relevant description
                        "Korrespondance": "Udgående",
                        "Dato": today_date,
                        "CCMMustBeOnPostList": "0"
                        }
        payload_mail = make_payload_document(ows_dict= ows_dict_mail, caseID= CaseID, FolderPath= "", byte_arr= byte_arr_mail, filename= "Svar på anmodning.pdf")
        response_mail = upload_document_go(go_ad_url, payload = payload_mail, session = session)
        if "DocId" in response_mail:
            uploaded_doc_ids.append(response_mail["DocId"])
        else:
            raise Exception("No doc id found for emailbody")
        delete_local_file(filsti = sent_mail_pdf_path)

    if SharepointMappelink:
        try:
            sp_site_url, sp_relative_folder = parse_sharepoint_url(SharepointMappelink)
            sharepoint_doc_ids = process_sharepoint_folders(
                sharepoint_site_url=sp_site_url,
                folders=[sp_relative_folder],
                go_api_url=go_ad_url,
                tenant=sp_tenant,
                client_id=sp_client_id,
                thumbprint=sp_thumbprint,
                cert_path=sp_cert_path,
                session=session,
                orchestrator_connection=orchestrator_connection,
                case_id=CaseID
            )
            uploaded_doc_ids.extend(sharepoint_doc_ids)
        except Exception as e:
            orchestrator_connection.log_error(f"Fejl ved hentning/upload af filer fra SharePoint-link {SharepointMappelink}: {e}")
            raise e

    # Journalisér alle uploadede dokumenter inden lukning
    if uploaded_doc_ids:
        try:
            payload_journaliser = {"DocumentIds": uploaded_doc_ids}
            url = f"{go_ad_url}/_goapi/Documents/MarkMultipleAsCaseRecord/ByDocumentId"
            response = session.post(url, data=json.dumps(payload_journaliser), headers={"Content-Type": "application/json"})
            response.raise_for_status()
            print(f"Journaliserede {len(uploaded_doc_ids)} dokumenter.")
        except Exception as e:
            print(f"Journalisering fejlede: {e}")

        try:
            finaliser_dokumenter(go_ad_url, uploaded_doc_ids, session, orchestrator_connection)
        except Exception as e:
            print(f"Endeliggørelse fejlede: {e}")

    close_case(CaseID, session, go_ad_url)


    SQL_SERVER = orchestrator_connection.get_constant('SqlServer').value 
    DATABASE_NAME = "AktindsigterPersonalemapper"

    odbc_str = (
        "DRIVER={SQL Server};"
        f"SERVER={SQL_SERVER};"
        f"DATABASE={DATABASE_NAME};"
        "Trusted_Connection=yes;"
    )

    odbc_str_quoted = quote_plus(odbc_str)
    engine = create_engine(f"mssql+pyodbc:///?odbc_connect={odbc_str_quoted}", future=True)

    sql = text("""
        UPDATE dbo.cases
        SET Journaliseringsmappelink = :link,
            last_run_complete = :ts
        WHERE aktid = :caseid
    """)

    with engine.begin() as conn:
        result = conn.execute(sql, {
            "link": CaseUrl_new.replace("ad.", "", 1),
            "ts": datetime.now(),
            "caseid": str(SagsID)
        })
        if result.rowcount == 0:
            print(f"⚠️ Ingen sag fundet med aktid={SagsID}")
        else:
            print(f"✅ Opdateret sag {SagsID} med journaliseringslink:")

def sharepoint_client(tenant: str, client_id: str, thumbprint: str, cert_path: str, sharepoint_site_url: str, orchestrator_connection: OrchestratorConnection) -> ClientContext:
    """
    Creates and returns a SharePoint client context.
    """
    # Authenticate to SharePoint
    cert_credentials = {
        "tenant": tenant,
        "client_id": client_id,
        "thumbprint": thumbprint,
        "cert_path": cert_path
    }
    ctx = ClientContext(sharepoint_site_url).with_client_certificate(**cert_credentials)

    # Load and verify connection
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    orchestrator_connection.log_info(f"Authenticated successfully. Site Title: {web.properties['Title']}")
    return ctx

def parse_sharepoint_url(full_url: str):

    parsed = urlparse(full_url)
    path_parts = [p for p in unquote(parsed.path).split("/") if p]

    if len(path_parts) < 2 or path_parts[0].lower() not in ("sites", "teams"):
        raise ValueError(f"Kunne ikke genkende SharePoint-linket som en sites-/teams-URL: {full_url}")

    site_type = path_parts[0].lower()
    site_name = path_parts[1]
    site_url = f"{parsed.scheme}://{parsed.netloc}/{site_type}/{site_name}"
    relative_folder = "/" + "/".join(path_parts)

    return site_url, relative_folder


def print_download_progress(offset, orchestrator_connection: OrchestratorConnection):
    orchestrator_connection.log_info(f"Downloadet {offset} bytes...")


def delete_document_go(go_api_url, doc_id, session):
    url = f"{go_api_url}/_goapi/Documents/ByDocumentId"
    payload = {
        "DocId": doc_id,
        "ForceDelete": True
    }
    response = session.delete(url, json=payload, timeout=1200)
    response.raise_for_status()


def create_and_delete_placeholder(go_api_url, case_id, folder_path, session, orchestrator_connection: OrchestratorConnection):
    """
    GO's API har ikke et rent 'opret mappe'-endpoint, så vi opretter mappen
    ved at uploade en lille placeholder-fil og slette den igen bagefter.
    """
    file_content = b"A"
    byte_array = list(file_content)

    ows_dict = {
        "Beskrivelse": "Leveret via SharePoint-journalisering",
        "CCMMustBeOnPostList": "0"
    }
    metadata_xml = ' '.join([f'ows_{k}="{v}"' for k, v in ows_dict.items()])
    metadata = f'<z:row xmlns:z="#RowsetSchema" {metadata_xml}/>'

    payload = {
        "Bytes": byte_array,
        "CaseId": case_id,
        "ListName": "Dokumenter",
        "FolderPath": folder_path,
        "FileName": "CreateFolder.txt",
        "Metadata": metadata,
        "Overwrite": True
    }

    try:
        orchestrator_connection.log_info("Uploader placeholder-fil for at oprette mappe...")
        upload_response = upload_document_go(go_api_url, payload, session)
        doc_id = upload_response.get("DocId")
        orchestrator_connection.log_info(f"Placeholder-fil uploadet med DocId: {doc_id}")

        orchestrator_connection.log_info("Sletter placeholder-fil igen...")
        delete_document_go(go_api_url, doc_id, session)
    except Exception as e:
        orchestrator_connection.log_info(f"Fejl ved oprettelse/sletning af placeholder-mappe: {e}")


def fetch_files_in_folder(ctx: ClientContext, folder_url, base_folder=""):
    files_array = []
    folder = ctx.web.get_folder_by_server_relative_url(folder_url).execute_query()
    files = folder.files.get().execute_query()
    folders = folder.folders.get().execute_query()


    for file in files:
        files_array.append({
            "ServerRelativeUrl": file.serverRelativeUrl,
            "UniqueId": file.unique_id,
            "Name": file.name,
            "FolderPath": base_folder
        })

    for subfolder in folders:
        subfolder_name = os.path.join(base_folder, subfolder.name)
        files_array.extend(fetch_files_in_folder(ctx, subfolder.serverRelativeUrl, subfolder_name))

    return files_array

def process_sharepoint_folders(sharepoint_site_url, folders, go_api_url, tenant, client_id, thumbprint, cert_path, session, orchestrator_connection: OrchestratorConnection, case_id):
    ctx = sharepoint_client(tenant, client_id, thumbprint, cert_path,sharepoint_site_url, orchestrator_connection)

    created_folders = set()  # Keep track of created folders
    today_date = datetime.now().strftime("%d-%m-%Y")

    timestamp = time.time()

    all_doc_ids = []  # Samles op og returneres, så process() kan journalisere/finalisere alt samlet til sidst

    for folder_url in folders:
        orchestrator_connection.log_info(f"Processing top-level folder: {folder_url}")
        files = fetch_files_in_folder(ctx, folder_url)

        # Group files by their subfolder paths
        files_by_subfolder = {}
        for file in files:
            folder_path = file["FolderPath"]
            if folder_path not in files_by_subfolder:
                files_by_subfolder[folder_path] = []
            files_by_subfolder[folder_path].append(file)

        # Process each subfolder separately
        for folder_path, folder_files in files_by_subfolder.items():
            orchestrator_connection.log_info(f"Processing subfolder: {folder_path}")

            folder_doc_ids = []

            for file in folder_files:
                elapsed = time.time() - timestamp
                if elapsed >= 30 * 60:  # 30 minutes in seconds
                    orchestrator_connection.log_info("30 minutter er gået, henter ny SharePoint-klient og nulstiller timestamp.")
                    ctx = sharepoint_client(tenant, client_id, thumbprint, cert_path,sharepoint_site_url, orchestrator_connection)

                    timestamp = time.time()
                orchestrator_connection.log_info(f"Uploading file: {file['Name']} in {folder_path}")

                # Download the file content
                try:
                    sp_file = File.open_binary(ctx, file['ServerRelativeUrl'])
                    file_content = sp_file.content
                except Exception:
                    orchestrator_connection.log_info("Downloading file failed, trying large file download from unique id")
                    large_file = ctx.web.get_file_by_id(file['UniqueId'])
                    local_filename = file['Name']

                    # Download large file to local storage
                    with open(local_filename, "wb") as local_file:
                        large_file.download_session(
                            local_file,
                            lambda offset: print_download_progress(offset, orchestrator_connection)
                        ).execute_query()

                    # Read the file content from the saved file
                    with open(local_filename, "rb") as local_file:
                        file_content = local_file.read()
                    
                    os.remove(local_filename)
                                
                
                byte_array = list(file_content)

                # Tilføjer '- Sharepoint' til filnavnet, så filen ikke risikerer at blive
                # overskrevet af en fil med samme navn fra en anden kilde (Overwrite=True
                # i make_payload_document gælder pr. filnavn+mappe)
                base_name, file_ext = os.path.splitext(file['Name'])
                upload_filename = f"{base_name} - Sharepoint{file_ext}"

                # Prepare metadata
                ows_dict = {
                    "Title": f"{base_name} - Sharepoint",
                    "CaseID": case_id,  # Replace with your case ID
                    "Beskrivelse": "Uploadet af Aktbob",  # Add relevant description
                    "Korrespondance": "Udgående",
                    "Dato": today_date,
                    "CCMMustBeOnPostList": "0"
                }

                # Create payload
                payload = make_payload_document(ows_dict, case_id, folder_path, byte_array, upload_filename)

                try:
                    if (len(file_content) / (1024 * 1024)) > 10:
                        raise Exception("File is larger than 10 MB, skipping normal upload to avoid errors")
                    # Attempt upload
                    response = upload_document_go(go_api_url, payload, session)
                    if "DocId" in response:
                        folder_doc_ids.append((response["DocId"], file['Name']))
                        if folder_path not in created_folders:
                            created_folders.add(folder_path)
                    else:
                        raise Exception("No DocId")
                except Exception as e:
                    orchestrator_connection.log_info(f"Failed to upload {file['Name']}: {e}")
                    max_retries = 3
                    for attempt in range(1, max_retries + 1):
                        try:
                            orchestrator_connection.log_info(f"Retry attempt {attempt} for {file['Name']} after error: {e}")
                            orchestrator_connection.log_info("Retrying with large upload...")

                            if folder_path not in created_folders:
                                orchestrator_connection.log_info(f"Creating folder: {folder_path}")
                                create_and_delete_placeholder(
                                    go_api_url,
                                    case_id,
                                    str(folder_path).replace("\\", "/"),
                                    session,
                                    orchestrator_connection
                                )
                                created_folders.add(folder_path)

                            large_response = upload_large_document(
                                go_api_url,
                                payload,
                                session,
                                file_content,
                                orchestrator_connection
                            )
                            large_response_json = json.loads(large_response)

                            if "DocId" in large_response_json:
                                folder_doc_ids.append((large_response_json["DocId"], file['Name']))
                                break  
                            else:
                                raise Exception(f"Failed upload for file: {file['Name']} in {folder_path}")
                        except Exception as retry_exception:
                            if attempt == max_retries:
                                raise retry_exception  
                            else:
                                orchestrator_connection.log_info(f"Retry {attempt} failed: {retry_exception}")

            # Sortér filer efter filnavn og læg doc-id'erne i den samlede liste.
            # Journalisering/finalisering sker ikke her - det sker samlet i process()
            # sammen med de øvrige uploadede dokumenter (GO-udleveringsmappe, mailbilag mv.)
            folder_doc_ids.sort(key=lambda x: x[1])
            all_doc_ids.extend(doc_id for doc_id, _ in folder_doc_ids)

    return all_doc_ids
