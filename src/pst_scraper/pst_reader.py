import os
import pandas as pd
from aspose.email.storage.pst import PersonalStorage, FolderInfo
from pst_scraper.email_reader import parse_mapi_message

def read_folder_emails_internal(pst: PersonalStorage, folder: FolderInfo, output_emails_path: str, output_attachments_path: str, initial_num_emails: int, initial_num_attachments: int) -> tuple[int, int]:
    """
    Reads emails and attachments from a folder and appends them to a csv file.
    
    Args:
        pst: The personal storage object.
        folder: The folder to read emails and attachments from.
        output_emails_path: The path to the csv file to write the emails to.
        output_attachments_path: The path to the csv file to write the attachments to.
        initial_num_emails: The number of emails to start the count at.
        initial_num_attachments: The number of attachments to start the count at.

    Raises:
        RuntimeError: If the message is encrypted.
    """
    num_emails = initial_num_emails
    num_attachments = initial_num_attachments

    for sub_folder in folder.get_sub_folders():
        added_emails, added_attachments = read_folder_emails_internal(sub_folder,  output_emails_path, output_attachments_path, num_emails, num_attachments)
        
        num_emails += added_emails
        num_attachments += added_attachments

    n = folder.content_count
    batch_size = 50

    for i in range(0, n, batch_size):
        batched_emails = []
        batched_attachments = []

        for messageInfo in folder.get_contents(i, batch_size):
            mapi = pst.extract_message(messageInfo)
            email_dict = parse_mapi_message(mapi)

            attachments = email_dict.pop("attachments")
            linked_messages = email_dict.pop("linked_messages")

            email_dict["attachment_ids"] = []
            for attachment in attachments:
                email_dict["attachment_ids"].append(num_attachments)
                batched_attachments.append(attachment)
                num_attachments += 1
            
            email_dict["linked_message_ids"] = []
            for linked_message in linked_messages:
                email_dict["linked_message_ids"].append(num_emails)
                batched_emails.append(linked_message)
                num_emails += 1

            batched_emails.append(email_dict)
            num_emails += 1

        batched_emails_df = pd.DataFrame(batched_emails)
        batched_attachments_df = pd.DataFrame(batched_attachments)

        batched_emails_df.to_csv(output_emails_path, mode="a", index=False, header=not os.path.exists(output_emails_path))
        batched_attachments_df.to_csv(output_attachments_path, mode="a", index=False, header=not os.path.exists(output_attachments_path))
    
    return num_emails, num_attachments

def read_folder_emails(pst: PersonalStorage, folder: FolderInfo, output_emails_path: str, output_attachments_path: str) -> tuple[int, int]:
    """
    Reads emails and attachments from a folder and appends them to a csv file.
    
    Args:
        pst: The personal storage object.
        folder: The folder to read emails and attachments from.
        output_emails_path: The path to the csv file to write the emails to.
        output_attachments_path: The path to the csv file to write the attachments to.

    Returns:
        The number of emails and attachments read.
    
    Raises:
        RuntimeError: If the message is encrypted.
    """
    num_emails = 0
    num_attachments = 0

    return read_folder_emails_internal(pst, folder, output_emails_path, output_attachments_path, num_emails, num_attachments)

def read_psts(pst_file_paths: list[str], output_emails_path: str, output_attachments_path: str) -> tuple[int, int]:
    """
    Reads emails and attachments from a list of pst files and writes them to a csv file.
    
    Args:
        pst_file_paths: The list of pst files to read emails and attachments from.
        output_emails_path: The path to the csv file to write the emails to.
        output_attachments_path: The path to the csv file to write the attachments to.

    Returns:
        The number of emails and attachments read.

    Raises:
        RuntimeError: If any of the messages are encrypted.
    """
    num_emails = 0
    num_attachments = 0
    for pst_file_path in pst_file_paths:
        pst = PersonalStorage.from_file(pst_file_path)
        pst_root = pst.root_folder

        added_emails, added_attachments = read_folder_emails_internal(pst, pst_root, output_emails_path, output_attachments_path, num_emails, num_attachments)
        num_emails += added_emails
        num_attachments += added_attachments

    return num_emails, num_attachments