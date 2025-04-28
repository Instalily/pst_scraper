import os
from aspose.email.storage.pst import *
from pst_scraper.email_reader import parse_mapi_message

def read_folder_emails(pst: PersonalStorage, folder: , output_emails_path, output_attachments_path) -> tuple[int, int]:
    num_emails = 0
    num_attachments = 0

    for sub_folder in folder.get_sub_folders():
        added_emails, added_attachments = read_folder_emails(sub_folder,  output_emails_file_name, output_attachments_file_name)
        
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

        batched_emails_df.to_csv(output_emails_path, mode="a", index=False, header=not os.path.exists(output_emails_file_name))
        batched_attachments_df.to_csv(output_attachments_path, mode="a", index=False, header=not os.path.exists(output_attachments_file_name))
    
    return num_emails, num_attachments

def read_pst(pst_file_path: str, output_emails_path: str, output_attachments_path: str) -> tuple[int, int]:
    pst = PersonalStorage.from_file(pst_file_path)
    pst_root = pst.root_folder

    return read_folder_emails(pst, pst_root, output_attachments_path, output_attachments_path)

def combine_data(dile_fi)