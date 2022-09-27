import sys
import win32com.client as client 
from datetime import datetime, timedelta
import os
import csv
import time
from pathlib import Path


def save_csv(path,data_list): 
    with open(f'{path}\\log.csv', "a", newline='', encoding='utf-8-sig') as fp:
        wr = csv.writer(fp, dialect='excel')
        wr.writerow(data_list)


# Remove backslash
def modify_path_name(path_name): 
    mod_path_name = path_name.replace('\\',r"/")
    print("The path name is modified successfully!")
    return mod_path_name

# Create Desired Directory 
def create_folder(PATH):
    Path(PATH).mkdir(parents=True, exist_ok=True)



# Download Attachments 
def download_attachments(path_name,date_today,status):
    total_accounts = 0
    total_folder = 0 
    total_messages = 0
    total_attachments = 0

    try:
        outlook = client.Dispatch("Outlook.Application")
        print("Called client dispatch")

        mapi = outlook.GetNamespace("MAPI")
        print("Got the Mapi object")
    except: 
        sys.exit("Mapi Object could not be created!")

    # Iterate through all accounts 
    for account in mapi.Accounts:
        total_accounts +=1
        email = account.DeliveryStore.DisplayName
        sender_name = email.split("@")[0]
        path_original_name = f"{path_name}/{sender_name}/{date_today}"
        all_folders = mapi.Folders(email).Folders

        for each_folder in all_folders:

                if each_folder.name == 'Inbox' or each_folder.name == 'Outbox' or each_folder.name == 'Drafts' or each_folder.name == '[Gmail]' or each_folder.name == 'RSS Feeds':
                    pass 
                else:
                    total_folder +=1
                    messages = each_folder.Items
                    print("Got all the messages!")


                try:
                    for message in list(messages):
                        print("Checking read/unread status")
                        if status.lower() == "read": 
                            if message.UnRead == True: 
                                continue 
                            else:
                                pass 
                        elif status.lower() == 'unread': 
                            if message.UnRead == True: 
                                pass 
                            else: 
                                continue 
                        else: 
                            pass
                        total_messages +=1

                        print("Downloading Attachaments...")
                        try:
                            s = message.sender
                            for idx,attachment in enumerate(message.Attachments):
                                if idx==0: 
                                    create_folder(path_original_name)
                                    print(f"Folder created for account {sender_name}")
                                attachment.SaveASFile(os.path.join(path_original_name, attachment.FileName))
                                print(f"attachment {attachment.FileName} from {s} saved")
                                total_attachments +=1
                        except Exception as e:
                            print("Error when saving the attachment:" + str(e))
                            print(path_original_name)
                
                except Exception as e:
                    print("Error when processing emails messages:" + str(e))

    return total_accounts,total_folder,total_messages,total_attachments       



def main(): 

    path_name = input("Please enter the path: ")
    status = "all"
    date_time_today =  datetime.today().strftime('%Y-%m-%d %H:%M:%S')
    date_today = f"{date_time_today}".split(" ")[0]
    path_name_modified = modify_path_name(path_name)
    start_time = time.time()
    try:
        account_count,folders_count,messages_count,attachments_count = download_attachments(path_name_modified,date_today,status)
    except Exception as e:
        sys.exit(e) 

    execution_time = f"{round(time.time() - start_time,2)}"
    try:
        save_csv(path_name,[f'{date_time_today}',account_count,folders_count,messages_count,attachments_count,execution_time])
    except Exception as e: 
        print(e)
main()