
from pathlib import Path
from imap_tools import MailBox,A
from datetime import datetime


# Remove backslash
def modify_path_name(path_name): 
    return path_name.replace('\\',r"/")

# Create Desired Directory 
def create_folder(PATH):
    Path(PATH).mkdir(parents=True, exist_ok=True)


# Scraper Function 
def scrape_attachments(path_name,sender_name,date_today,email,password):
    with MailBox('imap.outlook.com').login(email, password) as mailbox:
        q = A(new=True,from_=sender_name)
        for msg in mailbox.fetch(q): 
            sender = msg.from_values.email
            sender_name = sender.split("@")[0]
            path_original_name = f"{path_name}/{sender_name}/{date_today}"
            create_folder(path_original_name)
            if sender == sender:
                attachments = msg.attachments
                for att in attachments:
                    print(msg)
                    filename = att.filename.replace("\\", "_")
                    with open(f'{path_original_name}/{filename}', 'wb') as f:
                        f.write(att.payload)



def main(): 
    email = "Outlook email"
    password = "Password"
    path_name = r"The path of the directory"
    sender_name = "sender email address"
    date_today =  datetime.today().strftime('%Y-%m-%d')
    path_name_modified = modify_path_name(path_name)
    scrape_attachments(path_name_modified,sender_name,date_today,email,password)


main()