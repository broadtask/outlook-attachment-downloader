"""Import modules"""
import sys
from datetime import datetime
import os
from os import path
from pathlib import Path
import pandas as pd
from win32com import client
# MAIN_LIST = []


def save_csv_or_excel(date_and_time, data_list):
    '''Save Data to CSV'''

    filename = modify_date_time(date_and_time).replace(":", "_")
    try:
        pd.DataFrame([data_list]).to_csv(f'{filename}.csv', sep=',', encoding='utf-8',
                                         doublequote=False, index=False, mode="a", header=False)
    except:  # pylint: disable=W0702
        print(f"Error in saving {data_list}")


def convert_date(time_stamp):
    '''Convert the date and time format'''
    return f'''{datetime.strptime(f"{time_stamp}".split("+")
                             [0], '%Y-%m-%d %H:%M:%S').strftime("%m/%d/%Y %H:%M:%S")}'''


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
    #print("The path name is modified successfully!")
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

    new_filename = filename
    filename_split = filename.split(".")
    file_format = filename_split[-1]
    filename_split.pop()
    filename_without_format = ".".join(filename_split)
    if "(" in filename_without_format and ")" in filename_without_format:
        file_number = filename_without_format.split("(")[1].split(")")[
            0].strip()
        try:
            count_data = int(file_number)
        except:  # pylint: disable=W0702
            count_data = 1
    else:
        count_data = 1
    #print(filename_without_format, count_data)
    while True:
        if path.exists(pathname+'\\' + new_filename):
            # new_filename = f"{filename} ({count_data})"
            if (count_data == 1 and
                "(" not in filename_without_format and
                    ")" not in filename_without_format):
                new_filename = f"{filename_without_format} ({count_data}).{file_format}"
            else:
                count_text = f"({count_data})"
                filename_prev = f'{filename_without_format.replace(count_text,"")}'.strip(
                )

                new_filename = f"{filename_prev} ({count_data+1}).{file_format}"
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
        #print("Called client dispatch")

        mapi = outlook.GetNamespace("MAPI")
        #print("Got the Mapi object")
    except:  # pylint: disable=W0702
        sys.exit("Mapi Object could not be created!")

    # Iterate through all accounts
    for account in mapi.Accounts:
        email = account.DeliveryStore.DisplayName

        all_folders = mapi.Folders(email).Folders

        for each_folder in all_folders:
            first_email = ""
            last_email = ""
            unprocessed_email = 0
            attachment_count = 0
            email_with_invoice = 0

            if (each_folder.name in avoidable_folders or
                ":" in each_folder.name or
                "calendar" in each_folder.name.lower() or
                    "this computer only" in each_folder.name.lower()):
                continue

            path_original_name = f"{path_name}/{each_folder}"
            # path_original_name = f"{path_name}/{sender_name}/{date_today}/{each_folder}"
            messages = each_folder.Items
            try:
                messages.Sort("[ReceivedTime]", True)
            except:  # pylint: disable=W0702
                continue
            #print("Got all the messages!")

            try:
                last_index = len(list(messages))-1
                for indx_msg, message in enumerate(list(messages)):

                    if len(message.Attachments) == 0:
                        unprocessed_email += 1
                    try:
                        if indx_msg == 0:
                            first_email_recieved_date = convert_date(
                                message.ReceivedTime)
                            first_email_subject = message.Subject
                            first_email = f"{first_email_recieved_date} {first_email_subject}"
                    except:  # pylint: disable=W0702
                        continue

                    try:
                        if indx_msg == last_index:
                            last_email_recieved_date = convert_date(
                                message.ReceivedTime)
                            last_email_subject = message.Subject
                            last_email = f"{last_email_recieved_date} {last_email_subject}"
                    except:  # pylint: disable=W0702
                        continue

                    #print("Checking read/unread status")
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

                    #print("Downloading Attachments...")
                    try:

                        is_attachment_exist = False
                        is_pdf_exist = False
                        for idx, attachment in enumerate(message.Attachments):
                            is_attachment_exist = True
                            if 'pdf' in attachment.FileName:
                                attachment_count = attachment_count+1
                                is_pdf_exist = True

                            if idx == 0:

                                create_folder(path_original_name)
                            filename_new = check_similar_file(
                                attachment.FileName, path_original_name)
                            print("Previous File name: ", attachment.FileName)
                            print("New Filename: ", filename_new)
                            attachment.SaveASFile(os.path.join(
                                path_original_name, filename_new))
                            print(
                                f"attachment {attachment.FileName} saved")

                        if is_attachment_exist:
                            move_message(all_folders, date_and_time, message)
                        if is_pdf_exist:
                            email_with_invoice = email_with_invoice+1

                    except Exception as exception_error:  # pylint: disable=W0703
                        print("Error when saving the attachment:" +
                              str(exception_error))
                        print(path_original_name)

            except Exception as exception_error:  # pylint: disable=W0703
                print("Error when processing emails messages:" +
                      str(exception_error))

            if first_email == "" and last_email == "":
                pass
            else:
                data_list = [each_folder.name,
                             path_original_name,
                             email_with_invoice,
                             attachment_count,
                             date_today,
                             first_email,
                             last_email,
                             unprocessed_email]
                save_csv_or_excel(date_and_time, data_list)

            # MAIN_LIST.append([date_and_time, each_folder.name,
            #                  total_messages, total_attachments])

            # df_to_excel_main_list(MAIN_LIST, date_and_time)

    # return None


def main():
    """Main Function"""

    # path_name = input("Please enter download folder path: ")
    # path_name = r"C:\BOT\TAX_Tech_AvinashKaur\OutlookBot_V1\Data\Output"
    path_name = r"E:\Python\brend\job_2\Data"
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
    save_csv_or_excel(date_time_today, [
                      'Folder Name',
                      'Output Folder',
                      'Total Email with Invoices',
                      'Total Downloaded Invoices',
                      'Execution Date',
                      'First Email Details',
                      'Last Email Details',
                      'Unprocessed Emails'])

    try:
        download_attachments(path_name_modified, date_today,
                             status, date_time_today)
    except Exception as exception_message:  # pylint: disable=W0703
        sys.exit(exception_message)


if __name__ == '__main__':
    main()
