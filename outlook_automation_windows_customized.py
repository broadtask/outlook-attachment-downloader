"""Import modules"""
import sys
from datetime import datetime
import os
from os import path
from pathlib import Path
import pandas as pd
from win32com import client
MAIN_LIST = []


def df_to_excel_main_list(data_list, date_and_time):
    """Save data to excel"""
    filename = modify_date_time(date_and_time).replace(":", "_")
    my_python_list = data_list
    d_f = pd.DataFrame(columns=['Date Time',
                                'Folder Names',
                                'Message Count',
                                'Attachment Count'], data=my_python_list)
    # pylint: disable=E0110:
    with pd.ExcelWriter(f'log_{filename}.xlsx', engine='xlsxwriter') as writer:
        d_f.to_excel(writer, sheet_name='Sheet1', index=False)


def modify_path_name(path_name):
    """Remove backslash"""
    mod_path_name = path_name.replace('\\', r"/")
    print("The path name is modified successfully!")
    return mod_path_name


def modify_date_time(date_and_time):
    """Modify date time to text and remove the seconds"""
    date_time_text_split = f"{date_and_time}".split(":")
    date_time_text_split.pop()
    date_and_time_text = ":".join(date_time_text_split)
    return date_and_time_text


def create_folder(path_name):
    """Create desired directory"""
    Path(path_name).mkdir(parents=True, exist_ok=True)


def check_similar_file(filename, pathname):
    """Check similar file in the given directory"""
    count_data = 1
    new_filename = filename
    while True:

        if path.exists(pathname+'\\' + new_filename):
            new_filename = f"{filename} ({count_data})"
            count_data = count_data+1
            continue

        return new_filename


def move_message(folders_object_data, date_and_time, message):
    """Move messages to different folder"""
    date_and_time_text = modify_date_time(date_and_time)

    try:
        folders_object_data[date_and_time_text]
    except:  # pylint: disable=W0702
        folders_object_data.Add(date_and_time_text)

    message.Move(folders_object_data[date_and_time_text])


# pylint: disable=R1702

def download_attachments(path_name, date_today, status, date_and_time):  # pylint: disable=R0914
    # pylint: disable=R0915
    """Attachment downloading"""
    avoidable_folders = ['Inbox', 'Outbox',
                         'Drafts', '[Gmail]', 'RSS Feeds', '']
    try:
        outlook = client.Dispatch("Outlook.Application")
        print("Called client dispatch")

        mapi = outlook.GetNamespace("MAPI")
        print("Got the Mapi object")
    except:  # pylint: disable=W0702
        sys.exit("Mapi Object could not be created!")
    # Iterate through all accounts
    for account in mapi.Accounts:
        email = account.DeliveryStore.DisplayName
        sender_name = email.split("@")[0]

        all_folders = mapi.Folders(email).Folders

        for each_folder in all_folders:
            total_messages = 0
            total_attachments = 0

            if (each_folder.name in avoidable_folders or
                ":" in each_folder.name or
                "calendar" in each_folder.name.lower() or
                    "this computer only" in each_folder.name.lower()):
                continue

            path_original_name = f"{path_name}/{sender_name}/{date_today}/{each_folder}"
            messages = each_folder.Items
            print("Got all the messages!")

            try:
                for message in list(messages):
                    total_messages += 1
                    print("Checking read/unread status")
                    if status.lower() == "read":
                        if message.UnRead is True:
                            continue

                    elif status.lower() == 'unread':
                        if message.UnRead is True:
                            pass
                        else:
                            continue
                    else:
                        pass
                    total_messages += 1

                    print("Downloading Attachaments...")
                    try:

                        is_attachment_exist = False
                        for idx, attachment in enumerate(message.Attachments):
                            is_attachment_exist = True
                            if idx == 0:
                                create_folder(path_original_name)
                                print(
                                    f"Folder created for account {sender_name}")
                            filename_new = check_similar_file(
                                attachment.FileName, path_original_name)
                            print("Previous File name: ", attachment.FileName)
                            print("New Filename: ", filename_new)
                            attachment.SaveASFile(os.path.join(
                                path_original_name, filename_new))
                            print(
                                f"attachment {attachment.FileName} from {message.sender} saved")
                            total_attachments += 1
                        if is_attachment_exist:
                            move_message(all_folders, date_and_time, message)

                    except Exception as exception_error:  # pylint: disable=W0703
                        print("Error when saving the attachment:" +
                              str(exception_error))
                        print(path_original_name)

            except Exception as exception_error:  # pylint: disable=W0703
                print("Error when processing emails messages:" +
                      str(exception_error))
            MAIN_LIST.append([date_and_time, each_folder.name,
                             total_messages, total_attachments])

            df_to_excel_main_list(MAIN_LIST, date_and_time)

    return total_messages, total_attachments


def main():
    """Main Function"""

    path_name = input("Please enter download folder path: ")
    status_code = input(
        "Please choose\n1 for Read Emails\2 for Unread Emails\n3 for Read and Unread all Emails\n> "
    )
    if status_code == "1":
        status = "read"
    elif status_code == "2":
        status = "unread"
    else:
        status = "all"
    date_time_today = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
    date_today = f"{date_time_today}".split(" ")[0]
    path_name_modified = modify_path_name(path_name)

    try:
        download_attachments(path_name_modified, date_today,
                             status, date_time_today)
    except Exception as exception_message:  # pylint: disable=W0703
        sys.exit(exception_message)


if __name__ == '__main__':
    main()
