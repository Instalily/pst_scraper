from aspose.email.mapi import MapiMessage
from pst_scraper.email_enums import *
from base64 import b64encode

def parse_mapi_message(message: MapiMessage):
    """
    Parses a mapi message and returns a dictionary of the message.
    
    Args:
        message: The mapi message to parse.
    """
    if message.is_encrypted:
        raise RuntimeError("We can't work with encrypted files.")

    if message.is_signed:
        message.remove_signature()
    
    subject = message.subject
    conversation_topic = message.conversation_topic

    sender_email = message.sender_email_address
    sender_name = message.sender_name
    
    client_submit_time = message.client_submit_time
    delivery_time = message.delivery_time
    
    sensitivity = Sensitivity(message.sensitivity)

    body_type = BodyType(message.body_type)
    match body_type:
        case BodyType.TEXT:
            body = message.body
        case BodyType.HTML:
            body = message.body_html
        case BodyType.RTF:
            body = message.body_rtf

    body = str(body)
    if (body is None) or (body.lower() == "nan"):
        body = ""
            
    recipient_dict = {RecipientType.TO: [], RecipientType.CC: [], RecipientType.BCC: []}
    for recipient in message.recipients:
        display_name = recipient.display_name
        email_address = recipient.email_address
        recipient_type = RecipientType(recipient.recipient_type)

        recipient_dict[recipient_type].append(f"{display_name} <{email_address}>")
    
    attachments = []
    linked_messages = []
    for attachment in message.attachments:
        display_name = attachment.display_name
        if attachment.object_data:
            object_data = attachment.object_data
            if object_data and object_data.is_outlook_message:
                new_message = object_data.to_mapi_message()
                new_message_dict = parse_mapi_message(new_message)
                linked_messages.append(new_message_dict)
            else:
                raise RuntimeError("Attachment has object data but is not an outlook message, so we don't know how to handle it.")
        elif attachment.binary_data:
            attachment_dict = {"name": display_name, "data": bytes(attachment.binary_data)}
            attachments.append(attachment_dict)
        else:
            attachment_dict = {"name": display_name, "data": None}
            
    email_dict = {
        "subject": subject,
        "conversation_topic": conversation_topic,
        "sender_email": sender_email,
        "sender_name": sender_name,
        "client_submit_time": client_submit_time,
        "delivery_time": delivery_time,
        "sensitivity": sensitivity,
        "body_type": body_type,
        "body": body,
        "recipients": recipient_dict,
        "attachments": attachments,
        "linked_messages": linked_messages,
    }

    return email_dict