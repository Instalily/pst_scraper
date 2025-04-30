import os, csv, base64
from aspose.email.storage.pst import PersonalStorage, FolderInfo
from io import BytesIO
from pst_scraper.email_reader import parse_mapi_message
from pst_scraper.email_enums import *
from tqdm import tqdm

def parse_email_dict_internal(email_dict: dict, batched_emails: list[dict], batched_attachments: list[dict], attachments_dir: str, num_emails: int, num_attachments: int) -> tuple[int, int]:
    """
    Parses an email dictionary and appends it to a list of emails and attachments.
    
    Args:
        email_dict: The email dictionary to parse.
        batched_attachments: The list of attachments to append to.
        batched_emails: The list of emails to append to.
        attachments_dir: The directory to save attachments to.
        num_attachments: The number of attachments to start the count at.
        num_emails: The number of emails to start the count at.

    Returns:
        The number of emails and attachments read.
    """
    old_recipients = email_dict.pop("recipients")
    email_dict["to"] = old_recipients[RecipientType.TO]
    email_dict["cc"] = old_recipients[RecipientType.CC]
    email_dict["bcc"] = old_recipients[RecipientType.BCC]

    email_dict["sensitivity"] = email_dict["sensitivity"].name
    email_dict["body_type"] = email_dict["body_type"].name

    attachments = email_dict.pop("attachments")
    linked_messages = email_dict.pop("linked_messages")

    email_dict["linked_message_ids"] = []
    for linked_message in linked_messages:
        num_emails, num_attachments = parse_email_dict_internal(linked_message, batched_emails, batched_attachments, attachments_dir, num_emails, num_attachments)
        email_dict["linked_message_ids"].append(num_emails - 1)  # Use the last email's index

    email_dict["attachment_ids"] = []
    for attachment in attachments:
        attachment_data = attachment.pop("data")

        attachment_path = f"{attachments_dir}/{num_attachments}"
        with open(attachment_path, "wb") as f:
            f.write(attachment_data)
        
        batched_attachments.append(attachment)
        email_dict["attachment_ids"].append(num_attachments)
        num_attachments += 1
    
    batched_emails.append(email_dict)
    num_emails += 1

    return num_emails, num_attachments

    
def read_folder_emails_internal(pst: PersonalStorage, folder: FolderInfo, emails_csv_path: str, attachments_csv_path: str, attachments_dir: str, initial_num_emails: int, initial_num_attachments: int) -> tuple[int, int]:
    """
    Reads emails and attachments from a folder and appends them to a csv file.
    
    Args:
        pst: The personal storage object.
        folder: The folder to read emails and attachments from.
        emails_csv_path: The path to the csv file to write the emails to.
        attachments_csv_path: The path to the csv file to write the attachments to.
        attachments_dir: The directory to save attachments to.
        initial_num_emails: The number of emails to start the count at.
        initial_num_attachments: The number of attachments to start the count at.

    Raises:
        RuntimeError: If the message is encrypted.
    """
    if not os.path.exists(attachments_dir):
        os.makedirs(attachments_dir)

    num_emails = initial_num_emails
    num_attachments = initial_num_attachments

    for sub_folder in folder.get_sub_folders():
        added_emails, added_attachments = read_folder_emails_internal(pst, sub_folder, emails_csv_path, attachments_csv_path, attachments_dir, num_emails, num_attachments)
        
        num_emails += added_emails
        num_attachments += added_attachments

    n = folder.content_count
    batch_size = 50

    for i in tqdm(range(0, n, batch_size), desc=f"Reading emails and attachments from folder {folder.display_name}"):
        batched_emails = []
        batched_attachments = []

        for messageInfo in folder.get_contents(i, batch_size):
            mapi = pst.extract_message(messageInfo)
            email_dict = parse_mapi_message(mapi)

            num_emails, num_attachments = parse_email_dict_internal(email_dict, batched_emails, batched_attachments, attachments_dir, num_emails, num_attachments)

        write_emails_header = not os.path.exists(emails_csv_path)
        with open(emails_csv_path, "w" if write_emails_header else "a") as f:
            fc = csv.DictWriter(f, fieldnames=batched_emails[0].keys())
            
            if write_emails_header:
                fc.writeheader()

            fc.writerows(batched_emails)

        write_attachments_header = not os.path.exists(attachments_csv_path)
        with open(attachments_csv_path, "w" if write_attachments_header else "a") as f:
            fc = csv.DictWriter(f, fieldnames=batched_attachments[0].keys())
            
            if write_attachments_header:
                fc.writeheader()

            fc.writerows(batched_attachments)
    
    return num_emails, num_attachments

def read_folder_emails(pst: PersonalStorage, folder: FolderInfo, emails_csv_path: str, attachments_csv_path: str, attachments_dir: str) -> tuple[int, int]:
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

    return read_folder_emails_internal(pst, folder, emails_csv_path, attachments_csv_path, attachments_dir, num_emails, num_attachments)

def read_psts(pst_files: list[str] | list[BytesIO], emails_csv_path: str, attachments_csv_path: str, attachments_dir: str) -> tuple[int, int]:
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
    if os.path.exists(emails_csv_path):
        num_lines = 0
        with open(emails_csv_path, "r") as f:
            num_lines = sum(1 for _ in f)
        num_emails = num_lines - 1

    num_attachments = 0
    if os.path.exists(attachments_csv_path):
        num_lines = 0
        with open(attachments_csv_path, "r") as f:
            num_lines = sum(1 for _ in f)
        num_attachments = num_lines - 1

    for pst_file in pst_files:
        if isinstance(pst_file, BytesIO):
            pst = PersonalStorage.from_stream(pst_file)
        elif isinstance(pst_file, str):
            pst = PersonalStorage.from_file(pst_file)
        else:
            raise ValueError(f"Invalid pst file type: {type(pst_file)}")

        pst_root = pst.root_folder

        added_emails, added_attachments = read_folder_emails_internal(pst, pst_root, emails_csv_path, attachments_csv_path, attachments_dir, num_emails, num_attachments)
        num_emails += added_emails
        num_attachments += added_attachments

    return num_emails, num_attachments