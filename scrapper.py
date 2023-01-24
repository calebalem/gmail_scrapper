from __future__ import print_function
import base64
import datetime
import json
import os
import pathlib
import random
import re
import shutil
import time
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from azure.identity import TokenCachePersistenceOptions ,UsernamePasswordCredential,ClientSecretCredential,InteractiveBrowserCredential
from msgraph.core import GraphClient
from xls2xlsx import XLS2XLSX
from progress.bar import ChargingBar


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

FILE_EXT = [".xls"]
client = None
service = None


def main():
    try:
        global service
        config = json.load(open("config.json","r"))
        wait = config["time"]
        seconds_to_wait = wait * 3600
        if config["logout_google"] == True:
            logout_google(config)
        google_signin()
        labels = get_labels()
        current_user = create_profile()
        if config["clear_cache"] == True:
            clear_cache(config,current_user)
        one_drive_login()
        print("Login to your microsoft account.")
        test_upload()
        scraped_messages = get_scraped_messages(current_user)
        data = get_messages(current_user)
        print("\nScraper started...\n")
        if not data["messages"]:
            print('No messages found.')
            return
        files = []
        stop_scrapping = False
        progress_bar = ChargingBar(max=labels["total"]-len(scraped_messages),suffix='%(percent)d%% %(index)d/%(max)d')
        os.terminal_size([90,23])
        progress_bar.width = 20
        while True:
            stop_scrapping = False
            while not stop_scrapping:
                for i in range(len(data["messages"])):
                    message_id = data["messages"][i]['id']
                    if message_id not in scraped_messages:
                        res = service.users().messages().get(userId='me',id=message_id).execute()
                        subject = get_email_subject(res["payload"]["headers"])
                        progress_bar.bar_prefix = f"""Scrapping: {subject[0:20]}... """
                        progress_bar.next()
                        try:
                            parts = res["payload"]["parts"]
                        except Exception as e:
                            continue
                        for part in parts:
                            fname = part["filename"]
                            ext = pathlib.Path(fname).suffix
                            if fname != "" and ext.lower() in FILE_EXT and subject != "":
                                progress_bar.bar_prefix = f"""Found {fname}"""
                                files.append({"name":fname,"attachmentId":part["body"]["attachmentId"],"messageId":message_id,"title":f"""{subject}""","ext":ext.lower()})
                    else:
                        progress_bar.bar_prefix = "Done scrapping emails."
                        if len(files) == 0:
                            progress_bar.bar_prefix = "No new files found."
                        stop_scrapping = True
                        break
                if len(files) > 0:
                    print("\n\nDownloading .xls files in excel/")
                    save_xls_files(files)
                    print("\n\nConverting .xls files in excel/")
                    files_for_upload = convert_xls_to_xlsx(files)
                    print("\n\nUploading .xlsx files in converted/ to one drive")
                    upload_to_onedrive(files_for_upload,current_user)
                    files.clear()
                if not stop_scrapping:
                    scraped_messages = get_scraped_messages(current_user)
                    print("\nGetting next batch of emails to scrape...\n")
                    data = get_messages(current_user,data["next_page"])
                    if not data["messages"]:
                        print("No more messages to scrape...")
                        break
                else:
                    print("\nScrapping done for session...")
            print(f"\nSleeping for {wait} hours before next run...")
            time.sleep(seconds_to_wait)
            scraped_messages = get_scraped_messages(current_user)
            data = get_messages(current_user)
            if not data["messages"]:
                print('\nNo messages found.')
    except Exception as e:
        log_error(e)
        input("Press any key to quit")



def get_email_subject(headers):
    subject = ""
    try:
        for header in headers:
            if header["name"] == "Subject":
                val = header["value"]
                if len(val) < 3:
                    return ""
                if len(val) >= 71:
                    val = val[0:70]
                subject = re.sub('[^a-zA-Z0-9]','_',val)
                subject = f"""{subject}"""
        return subject
    except Exception as e:
        log_error(e)
        return subject



def google_signin():
    try:
        
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        try:
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'credentials.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('token.json', 'w') as token:
                    token.write(creds.to_json())
        except Exception as e:
            print(e)

        try:
            # Call the Gmail API
            global service
            service = build('gmail', 'v1', credentials=creds)
        except HttpError as error:
            # TODO(developer) - Handle errors from gmail API.
            print(f'An error occurred: {error}')
    except Exception as e:
        log_error(e)
      
def clear_cache(config,current_user):
    try:
        shutil.rmtree(current_user)
        file = open("config.json","w")
        config["clear_cache"] = False
        file.write(json.dumps(config))
        file.flush()
        file.close()
        create_profile()
        print(f"Cache clear for user with email {current_user}")
    except Exception as e:
        log_error(e)
       

def create_profile():
    try:
        res = service.users().getProfile(userId="me").execute()
        current_user = res["emailAddress"].split("@")[0]
        if not os.path.exists(current_user):
            os.mkdir(current_user)
        return current_user
    except Exception as e:
        log_error(e)
       

def logout_google(config):
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
            print("You're now logged out of gmail.")
            file = open("config.json","w")
            config["logout_google"] = False
            file.write(json.dumps(config))
            file.flush()
            file.close()
        else:
            print("You're already logged out.")
    except Exception as e:
        log_error(e)
       

def get_messages(current_user,page = 0):
    try:
        results = None
        if not page:
            results = service.users().messages().list(userId='me',maxResults=100).execute()
        else:
            results = service.users().messages().list(userId='me',maxResults=100,pageToken=page).execute()
        messages = results.get('messages', [])
        nextPageToken = results.get('nextPageToken')
        scraped_messages = open(f"{current_user}/scrapped_messages.txt","a")
        for i in range(len(messages)):
            message_id = messages[i]['id']
            scraped_messages.write(f"{message_id}\n")
        return {"messages":messages,"next_page":nextPageToken}
    except Exception as e:
        log_error(e)
        

def get_scraped_messages(current_user):
    try:
        if not os.path.exists(f"{current_user}/scrapped_messages.txt"):
            return []
        scraped_messages_file = open(f"{current_user}/scrapped_messages.txt","r")
        scraped_messages = scraped_messages_file.read()
        if scraped_messages == "":
            return []
        message_ids = scraped_messages.splitlines()
        return message_ids
    except Exception as e:
        log_error(e)
        return []
       

def save_xls_files(files):
    progress_bar = ChargingBar(max= len(files),suffix='%(percent)d%% %(index)d/%(max)d')
    try:
        if not os.path.exists("excel"):
            os.mkdir("excel")
        for file in files:
           
                progress_bar.bar_prefix = f"""Downloading {file["title"][0:20]}{file["ext"]}"""
                res = service.users().messages().attachments().get(userId='me',messageId=file["messageId"],id=file["attachmentId"]).execute()
                xls_data = base64.urlsafe_b64decode(res['data'])
                with open(f"""excel/{file["title"]}{file["ext"]}""", 'wb') as f:
                    f.write(xls_data)
                    f.flush()
                    f.close()
                    progress_bar.next()
            
    except Exception as e:
        log_error(e)
       

def convert_xls_to_xlsx(files):
    files_for_upload = []
    progress_bar = ChargingBar(max= len(files),suffix='%(percent)d%% %(index)d/%(max)d')
    try:
        if not os.path.exists("converted_files"):
            os.mkdir("converted_files")
        for file in files:
            try:
                progress_bar.bar_prefix = f"""Converting {file["title"][0:20]}{file["ext"]} to .xlsx"""
                xls = XLS2XLSX(f"""excel/{file["title"]}{file["ext"]}""")
                xls.ignore_workbook_corruption = True
                xls.to_xlsx(f"""converted_files/{file["title"]}.xlsx""")
                # df = pd.read_excel(f"""excel/{file["title"]}{file["ext"]}""")
                # df.to_excel(f"""converted_files/{file["title"]}.xlsx""")
                # fname = f"""excel/{file["title"]}{file["ext"]}"""
                # excel = win32.gencache.EnsureDispatch('Excel.Application')
                # wb = excel.Workbooks.Open(fname)
                # wb.SaveAs(f"""converted_files/{file["title"]}.xlsx""", FileFormat = 51)
                # wb.Close()
                # excel.Application.Quit()   
                progress_bar.next()
                files_for_upload.append(file)
            except Exception as e:
                log_error(e)
                pass
        return files_for_upload
    except Exception as e:
        log_error(e)
        return files_for_upload
        

def get_labels():
    res = service.users().labels().get(userId="me",id="INBOX").execute()
    total_messages = res["messagesTotal"]
    unread_messages = res["messagesUnread"]
    return {"total":total_messages,"unread":unread_messages}

def one_drive_login():
    try:
        global client
        cache = TokenCachePersistenceOptions(allow_unencrypted_storage=True)
        credential = InteractiveBrowserCredential(client_id="Your client_id",cache_persistence_options=cache)
        client = GraphClient(credential=credential)
       
    except Exception as e:
        log_error(e)

def log_error(e):
     file = ""
     if not os.path.exists("error.log"):
        file = open("error.log","w")
     else:
        file = open("error.log","a")
     file.write(f"""\ntime: {datetime.datetime.now()}\n error: {str(e)}\n""")
     file.flush()
     file.close()      

def test_upload():
    try:
        file = open("this_file_is_to_authenticate_onedrive_for_test.txt","w")
        file.write("to authenticate one drive at the begining of app start instead of upload")
        file.flush()
        file.close()
        client.put(f"""/me/drive/root:/temp/this_file_is_to_authenticate_onedrive_for_test.txt:/content""",headers={"Content-Type":"text/plain"})
    except Exception as e:
        log_error(e)  

def upload_to_onedrive(files,current_user):
    progress_bar = ChargingBar(max =len(files))
    if not os.path.exists("failed_uploads"):
        os.mkdir("failed_uploads")
    try:
        global client
        for file in files:
            progress_bar.bar_prefix = f"""Uploading {file["title"][0:20]}.xlsx to one drive... """
            data = open(f"""converted_files/{file["title"]}.xlsx""","rb")
            try:
                res = client.put(f"""/me/drive/root:/xlsx/{file["title"]}.xlsx/createUploadSession""",data.read(),headers={"Content-Tyep":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
                progress_bar.next()
            except Exception as e:
                 data.close()
                 shutil.move(f"""converted_files/{file["title"]}.xlsx""",f"""failed_uploads/{file["title"]}.xlsx""")
                 log_error(e)
                 continue
        print("Removing .xlsx files in local...")
        shutil.rmtree("converted_files")
        print("Removing .xls files in local...")
        shutil.rmtree("excel")     
    except Exception as e:
        log_error(e)
       
   
main()
