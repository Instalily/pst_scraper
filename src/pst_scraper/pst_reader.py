import os, csv, base64
from aspose.email.storage.pst import PersonalStorage, FolderInfo
from pst_scraper.email_reader import parse_mapi_message
from pst_scraper.email_enums import *
from io import BytesIO
def parse_email_dict_internal(email_dict: dict, batched_emails: list[dict], batched_attachments: list[dict], accounts: dict[str, str], batched_emails_to_recipients: list[dict], attachments_dir: str, num_emails: int, num_attachments: int, linked_from : int = -1) -> tuple[int, int]:
    """
    Parses an email dictionary and appends it to a list of emails and attachments.
    
    Args:
        email_dict: The email dictionary to parse.
        batched_attachments: The list of attachments to append to.
        batched_emails: The list of emails to append to.
        accounts: The dictionary of accounts to append to.
        batched_emails_to_recipients: The list of emails to recipients to append to.
        attachments_dir: The directory to save attachments to.
        num_attachments: The number of attachments to start the count at.
        num_emails: The number of emails to start the count at.
        linked_from: The id of the email that this email is linked from, or -1 if it is not linked to by any email.

    Returns:
        The number of emails and attachments read.
    """
    email_dict["id"] = num_emails
    num_emails += 1

    email_dict["linked_from"] = linked_from

    email_dict["sender"] = email_dict.pop("sender_email")
    if email_dict["sender"]:
        email_dict["sender"] = email_dict["sender"].lower()
    
    sender_name = email_dict.pop("sender_name")

    if email_dict["sender"] and email_dict["sender"] not in accounts:
        accounts[email_dict["sender"]] = sender_name

    old_recipients = email_dict.pop("recipients")
    for recipient in old_recipients:
        if recipient["email_address"] is None:
            continue

        recipient["email_address"] = recipient["email_address"].lower()
        if recipient["email_address"] not in accounts:
            accounts[recipient["email_address"]] = recipient["display_name"]

        batched_emails_to_recipients.append({
            "email_id": email_dict["id"],
            "account_id": recipient["email_address"],
            "recipient_type": recipient["recipient_type"].name
        })

    email_dict["sensitivity"] = email_dict["sensitivity"].name
    email_dict["body_type"] = email_dict["body_type"].name

    # For CSV compatibility, escape any quotes in the body and wrap in quotes
    # This ensures the body appears as a single field in CSV while remaining somewhat readable
    # Escape quotes by doubling them for CSV compatibility
    email_dict["body"] = email_dict["body"].replace('"', '""')
    # Replace newlines with escaped versions for CSV compatibility
    # To undo this transformation: body = body.replace('\\n', '\n').replace('\\r', '\r')
    email_dict["body"] = email_dict["body"].replace('\n', '\\n').replace('\r', '\\r')
    # Note: This transformation preserves the content but changes the format.
    # It can be reversed using the replace operations in reverse order.
    
    linked_messages = email_dict.pop("linked_messages")
    for linked_message in linked_messages:
        num_emails, num_attachments = parse_email_dict_internal(linked_message, batched_emails, batched_attachments, accounts, batched_emails_to_recipients, attachments_dir, num_emails, num_attachments, linked_from = email_dict["id"])

    attachments = email_dict.pop("attachments")
    for attachment in attachments:
        attachment["id"] = num_attachments
        attachment["email_id"] = email_dict["id"]

        attachment_path = f"{attachments_dir}/{num_attachments}"
        with open(attachment_path, "wb") as f:
            f.write(attachment.pop("data"))

        attachment["path"] = attachment_path
        
        batched_attachments.append(attachment)
        num_attachments += 1
    
    batched_emails.append(email_dict)
    return num_emails, num_attachments

    
def read_folder_emails_internal(pst: PersonalStorage, folder: FolderInfo, emails_csv_path: str, attachments_csv_path: str, accounts: dict[str, str], emails_to_recipients_csv_path: str, attachments_dir: str, initial_num_emails: int, initial_num_attachments: int) -> tuple[int, int]:
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
        num_emails, num_attachments = read_folder_emails_internal(pst, sub_folder, emails_csv_path, attachments_csv_path, accounts, emails_to_recipients_csv_path, attachments_dir, num_emails, num_attachments)

    n = folder.content_count
    batch_size = 50

    for i in range(0, n, batch_size):
        batched_emails = []
        batched_attachments = []
        batched_emails_to_recipients = []

        for messageInfo in folder.get_contents(i, batch_size):
            mapi = pst.extract_message(messageInfo)
            email_dict = parse_mapi_message(mapi)
            num_emails, num_attachments = parse_email_dict_internal(email_dict, batched_emails, batched_attachments, accounts, batched_emails_to_recipients, attachments_dir, num_emails, num_attachments)

        write_emails_header = not os.path.exists(emails_csv_path)
        with open(emails_csv_path, "w" if write_emails_header else "a") as f:
            fc = csv.DictWriter(f, fieldnames=["subject","conversation_topic","client_submit_time","delivery_time","sensitivity","body_type","body","id","linked_from","sender"])
            
            if write_emails_header:
                fc.writeheader()

            fc.writerows(batched_emails)

        write_attachments_header = not os.path.exists(attachments_csv_path)
        with open(attachments_csv_path, "w" if write_attachments_header else "a") as f:
            fc = csv.DictWriter(f, fieldnames=["name","id","email_id","path"])
            
            if write_attachments_header:
                fc.writeheader()

            fc.writerows(batched_attachments)

        write_emails_to_recipients_header = not os.path.exists(emails_to_recipients_csv_path)
        with open(emails_to_recipients_csv_path, "w" if write_emails_to_recipients_header else "a") as f:
            fc = csv.DictWriter(f, fieldnames=["email_id","account_id","recipient_type"])
            
            if write_emails_to_recipients_header:
                fc.writeheader()

            fc.writerows(batched_emails_to_recipients)
    
    return num_emails, num_attachments

def read_folder_emails(pst: PersonalStorage, folder: FolderInfo, emails_csv_path: str, attachments_csv_path: str, accounts: dict[str, str], emails_to_recipients_csv_path: str, attachments_dir: str) -> tuple[int, int]:
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

    return read_folder_emails_internal(pst, folder, emails_csv_path, attachments_csv_path, accounts, emails_to_recipients_csv_path, attachments_dir, num_emails, num_attachments)

def read_psts(pst_files: list[str] | list[BytesIO], emails_csv_path: str, attachments_csv_path: str, accounts_csv_path: str, emails_to_recipients_csv_path: str, attachments_dir: str) -> tuple[int, int]:
    """
    Reads emails and attachments from a list of pst files and writes them to a csv file.
    
    Args:
        pst_files: The list of pst files to read emails and attachments from.
        output_emails_path: The path to the csv file to write the emails to.
        output_attachments_path: The path to the csv file to write the attachments to.
        output_accounts_path: The path to the csv file to write the accounts to.
        output_emails_to_recipients_path: The path to the csv file to write the emails to recipients to.

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

    accounts = {}
    if os.path.exists(accounts_csv_path):
        with open(accounts_csv_path, "r") as f:
            reader = csv.DictReader(f)
            next(reader)
            for row in reader:
                accounts[row["email"]] = row["display_name"]

    for pst_file in pst_files:
        if isinstance(pst_file, str):
            pst = PersonalStorage.from_file(pst_file)
        elif isinstance(pst_file, BytesIO):
            pst = PersonalStorage.from_stream(pst_file)
        else:
            raise ValueError(f"Invalid pst file type: {type(pst_file)}")

        pst_root = pst.root_folder
        num_emails, num_attachments = read_folder_emails_internal(pst, pst_root, emails_csv_path, attachments_csv_path, accounts, emails_to_recipients_csv_path, attachments_dir, num_emails, num_attachments)

    accounts_list = [{"email": email, "display_name": display_name} for email, display_name in accounts.items()]
    with open(accounts_csv_path, "w") as f:
        fc = csv.DictWriter(f, fieldnames=["email","display_name"])
        fc.writeheader()
        fc.writerows(accounts_list)

    return num_emails, num_attachments