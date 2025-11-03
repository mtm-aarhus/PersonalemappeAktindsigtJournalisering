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

def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:

    gotesturl = orchestrator_connection.get_constant('GOApiTESTURL').value
    # go_api_url = orchestrator_connection.get_constant("GOApiURL").value
    # go_api_login = orchestrator_connection.get_credential("GOAktApiUser")
    # robot_user = orchestrator_connection.get_credential("Robot365User")
    # username = robot_user.username
    # password = robot_user.password
    # go_username = go_api_login.username
    # go_password = go_api_login.password
    go_test_login = orchestrator_connection.get_credential("GOTestApiUser")
    go_username_test = go_test_login.username
    go_password_test = go_test_login.password

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
    session = create_session(go_username_test, go_password_test)
    if Journaliseringsmappelink:
            #hvis der allerede ligger en journaliseringsmappe skal den slettes for ikke at have dobbeltmapper til at ligge
        JournaliseringsmappeID = Journaliseringsmappelink.rsplit("/")[-1]
        print(f'Gammel journaliseringsmappe detekteret {JournaliseringsmappeID}')
        try:
            delete_case_go(gotesturl, session, JournaliseringsmappeID)
            print(f'Gammel delingsmappe slettet for sag {JournaliseringsmappeID}')
        except Exception as e:
            print(f"Tried to delete old journaliseringsmappe, but failed {e}")

    #1 - definer stuff
    today_date = datetime.now().strftime("%d-%m-%Y")

    #Hent filoplysninger på færdigbehandlet go-sag
    SagsMetaData = get_case_metadata(gotesturl, SagsNummer, session)
    SagsMetaData = json.loads(SagsMetaData).get("Metadata")
    xdoc = ET.fromstring(SagsMetaData)

    # Extract attributes
    RelativeSagsUrl = xdoc.attrib.get("ows_CaseUrl")
    SagsTitel = xdoc.attrib.get("ows_Title")

    #Hent info om sag, der skal journaliseres
    casefiles = get_case_documents(session, gotesturl, SagsURL= RelativeSagsUrl, SagsID = RelativeSagsUrl.rsplit('/')[-1])

    #Lav ny sag til at journalisere ind i
    session.headers.clear()
    CreatedCase = json.loads(create_case(gotesturl, SagsTitel, SagsID, session))
    RelativeSagsUrl = CreatedCase['CaseRelativeUrl']
    CaseID = CreatedCase['CaseID']
    CaseUrl_new = f'{gotesturl}/{RelativeSagsUrl}'

    for item in casefiles:
        DokTitle= item.get("Title", "")
        DokID = str(item.get("DocID"))
        #funktionen skal rettes tilbage til ikke testversion i drift
        download_file(go_url = gotesturl, file_path= DokTitle, DokumentID= DokID, GoUsername= go_username_test, GoPassword= go_password_test )

        with open(DokTitle, "rb") as local_file:
            print(f'Opened file {DokTitle}')
            file_content = local_file.read()
            byte_arr = list(file_content)

        ows_dict = {
                    "Title": DokTitle,
                    "CaseID": CaseID,  # Replace with your case ID
                    "Beskrivelse": "Uploaded af personaleaktbob",  # Add relevant description
                    "Korrespondance": "Udgående",
                    "Dato": today_date,
                    "CCMMustBeOnPostList": "0"
                    }
            
        payload = make_payload_document(ows_dict= ows_dict, caseID = CaseID, FolderPath= "", byte_arr= byte_arr, filename = DokTitle )
        upload_document_go(gotesturl, payload = payload, session = session)
        delete_local_file(filsti = DokTitle)

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
    upload_document_go(gotesturl, payload = payload_mail, session = session)
    delete_local_file(filsti = application_pdf_path)


    #Journalising the answer
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
    upload_document_go(gotesturl, payload = payload_mail, session = session)
    delete_local_file(filsti = sent_mail_pdf_path)

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