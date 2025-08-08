from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
from datetime import datetime
import json
import time
from Funktioner import *
import xml.etree.ElementTree as ET

# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:

    gotesturl = orchestrator_connection.get_constant('GOApiTESTURL').value
    go_api_url = orchestrator_connection.get_constant("GOApiURL").value
    go_api_login = orchestrator_connection.get_credential("GOAktApiUser")
    robot_user = orchestrator_connection.get_credential("Robot365User")
    username = robot_user.username
    password = robot_user.password
    go_username = go_api_login.username
    go_password = go_api_login.password
    go_test_login = orchestrator_connection.get_credential("GOTestApiUser")
    go_username_test = go_test_login.username
    go_password_test = go_test_login.password

    specific_content = json.loads(queue_element.data)

    Udleveringsmappe = specific_content.get('Udleveringsmappe')
    SagsNummer = Udleveringsmappe
    SagsID = specific_content.get('caseid')
    SagsTitel = specific_content.get('PersonaleSagsTitel')

    #1 - definer sharepointsite url og mapper
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    today_date = datetime.now().strftime("%d-%m-%Y")

    #Hent filoplysninger på færdigbehandlet go-sag
    session = create_session(go_username_test, go_password_test)
    SagsMetaData = get_case_metadata(gotesturl, SagsNummer, session)
    SagsMetaData = json.loads(SagsMetaData).get("Metadata")
    xdoc = ET.fromstring(SagsMetaData)

    # Extract attributes
    RelativeSagsUrl = xdoc.attrib.get("ows_CaseUrl")
    CaseUrl = f'{gotesturl}/{RelativeSagsUrl}'
    SagsTitel = xdoc.attrib.get("ows_Title")

    #Hent info om sag, der skal journaliseres
    casefiles = get_case_documents(session, gotesturl, SagsURL= RelativeSagsUrl, SagsID = SagsTitel)

    #Lav ny sag til at journalisere ind i
    CreatedCase = json.loads(create_case(gotesturl, SagsTitel, SagsID, session))
    RelativeSagsUrl = CreatedCase['CaseRelativeUrl']
    CaseID = CreatedCase['CaseID']
    CaseUrl_new = f'{gotesturl}/{RelativeSagsUrl}'

    for item in casefiles:
        DokTitle= item.get("Title", "")
        DokID = str(item.get("DocID"))
        filepath = f'{downloads_folder}\{DokTitle}'
        #funktionen skal rettes tilbage til ikke testversion i drift
        download_file(go_url = gotesturl, file_path= filepath, DokumentID= DokID, GoUsername= go_username_test, GoPassword= go_password_test )

        with open(filepath, "rb") as local_file:
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
        session = create_session(go_username_test, go_password_test)
        upload_document_go(gotesturl, payload = payload, session = session)
        delete_local_file(filsti = filepath)

    # url = "/api/callback"
    # headers = {
    #     "Content-Type": "application/json",
    #     "X-API-Key": "secret_key"
    # }
    # payload = {
    #     "title": "journalisering",
    #     "Journaliseringsmappelink": CaseUrl_new
    # }

    # response = requests.post(url, json=payload, headers=headers)

    # print("Status:", response.status_code)
    # print("Respons:", response.json())
