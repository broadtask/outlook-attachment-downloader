import win32com.client as client 
from datetime import datetime, timedelta
import os
from pathlib import Path

# Remove backslash
def modify_path_name(path_name): 
    mod_path_name = path_name.replace('\\',r"/")
    print("The path name is modified successfully!")
    return mod_path_name

# Create Desired Directory 
def create_folder(PATH):
    Path(PATH).mkdir(parents=True, exist_ok=True)


# Download Attachments 
def download_attachments(path_name,sender_name,date_today,given_email):

    outlook = client.Dispatch("Outlook.Application")
    print("Called client dispatch")

    mapi = outlook.GetNamespace("MAPI")
    print("Got the Mapi object")

    # Check Username 
    isAccountFound = False
    for account in mapi.Accounts:
        email = account.DeliveryStore.DisplayName

        if email == given_email:
            isAccountFound = True
            print("Email Matched with given email!")

            sender = sender_name.split("@")[0]
            path_original_name = f"{path_name}/{sender}/{date_today}"
            create_folder(path_original_name)
            print("Created folder with given path name")

            print("Trying to download the attachments")
            inbox = mapi.Folders(given_email).Folders("Inbox")
            print("Got inbox object for given email")
            messages = inbox.Items 
            print("Got messages in the folder")
            try:
                for message in list(messages):

                    if message.SenderEmailAddress == sender_name:
                        try:
                            s = message.sender
                            for attachment in message.Attachments:
                                attachment.SaveASFile(os.path.join(path_original_name, attachment.FileName))
                                print(f"attachment {attachment.FileName} from {s} saved")
                        except Exception as e:
                            print("Error when saving the attachment:" + str(e))
            except Exception as e:
                print("Error when processing emails messages:" + str(e))
    if isAccountFound == False: 
        print("Sorry your account is not present in outlook!")  



def main(): 
    email = "Please enter email here"
    path_name = r"Path name here"
    sender_name = "Sender email here"
    date_today =  datetime.today().strftime('%Y-%m-%d')
    print(f"Todays Date: {date_today}")
    path_name_modified = modify_path_name(path_name)
    download_attachments(path_name_modified,sender_name,date_today,email)
main()