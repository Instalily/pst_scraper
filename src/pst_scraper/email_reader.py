import os
import pandas as pd
from aspose.email.mapi import MapiMessage
from magic import from_buffer

from pst_scraper.email_enums import *

def parse_mapi_message(message: MapiMessage):
    """
    Parses a mapi message and returns a dictionary of the message.
    
    Args:
        message: The mapi message to parse.
    """
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

    if body:
        body = body.replace("\n", "\\n")
            
    recipient_dict = {"to": [], "cc": [], "bcc": []}
    for recipient in message.recipients:
        display_name = recipient.display_name
        email_address = recipient.email_address
        recipient_type = RecipientType(recipient.recipient_type)

        recipient_dict[recipient_type.name.lower()].append(f"{display_name} <{email_address}>")
    
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
                raise RuntimeError("Attachment has no binary data and either has no object data or is not an outlook message. Either way, we don't know how to handle it.")
        elif attachment.binary_data:
            binary_data = bytes(attachment.binary_data)
            mime_data = from_buffer(binary_data, mime=True)

            attachment_dict = {"name": display_name, "type": mime_data, "data": binary_data}
            attachments.append(attachment_dict)
        else:
            attachment_dict = {"name": display_name, "type": attachment.mime_tag, "data": None}
            
    email_dict = {
        "subject": subject,
        "conversation_topic": conversation_topic,
        "sender_email": sender_email,
        "sender_name": sender_name,
        "client_submit_time": client_submit_time,
        "delivery_time": delivery_time,
        "sensitivity": sensitivity.name,
        "body_type": body_type.name,
        "body": body,
        "recipients": recipient_dict,
        "attachments": attachments,
        "linked_messages": linked_messages
    }

    return email_dict