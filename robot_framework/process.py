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

    specific_content = json.loads(queue_element.data)

    Udleveringsmappelink = specific_content.get('Udleveringsmappelink') 
    SagsNummer = Udleveringsmappelink.rsplit("/")[-1] 
    SagsID = specific_content.get('caseid') 
    SagsTitel = specific_content.get('PersonaleSagsTitel') 
    Journaliseringsmappelink = specific_content.get('Journaliseringsmappelink')
    EmailBody = specific_content.get('EmailBody')
    MailModtager = specific_content.get("MailModtager")
    MailAfsender = specific_content.get("MailAfsender")
    Beskrivelse = specific_content.get("Beskrivelse")
    Modtagelsesdato = specific_content.get("Modtagelsesdato")

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
    SagsMetaData = get_case_metadata(go_ad_url, SagsNummer, session)
    SagsMetaData = json.loads(SagsMetaData).get("Metadata")
    xdoc = ET.fromstring(SagsMetaData)

    # Extract attributes
    RelativeSagsUrl = xdoc.attrib.get("ows_CaseUrl")
    SagsTitel = xdoc.attrib.get("ows_Title")

    #Hent info om sag, der skal journaliseres
    casefiles = get_case_documents(session, go_ad_url, SagsURL= RelativeSagsUrl, SagsID = RelativeSagsUrl.rsplit('/')[-1])

    #Lav ny sag til at journalisere ind i
    session.headers.clear()
    CreatedCase = json.loads(create_case(go_ad_url, SagsTitel, SagsID, session))
    RelativeSagsUrl = CreatedCase['CaseRelativeUrl']
    CaseID = CreatedCase['CaseID']
    CaseUrl_new = f'{go_ad_url}/{RelativeSagsUrl}'

    #Sagsbehandler sættes som det første
    mailHR = orchestrator_connection.get_constant('balas').value
    try:
        update_case_owner(go_ad_url, go_ad_username, go_ad_password, CaseID, MailAfsender, mailHR)
    except Exception as e:
        orchestrator_connection.log_error("Kunne ikke sætte sagsbehandler")
        raise e

    ikke_konverterede_filer = []
    fejlede_uploads = []
    uploaded_doc_ids = []

    for item in casefiles:
        DokTitle = item.get("Title", "")
        DokID = str(item.get("DocID"))
        file_path = DokTitle

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
            "link": CaseUrl_new,
            "ts": datetime.now(),
            "caseid": str(SagsID)
        })
        if result.rowcount == 0:
            print(f"⚠️ Ingen sag fundet med aktid={SagsID}")
        else:
            print(f"✅ Opdateret sag {SagsID} med journaliseringslink:")