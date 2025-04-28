# PST to CSV Converter

A Python utility that converts Microsoft Outlook PST files into structured CSV documents. This tool extracts emails, attachments, and linked messages from PST files and organizes them into separate CSV files for easy analysis and processing.

## Features

- Extracts emails from PST files with full metadata
- Handles email attachments and linked messages
- Outputs data to CSV format for easy importing into databases or analysis tools
- Processes folders recursively, maintaining PST folder structure
- Batched processing for memory efficiency

## Installation

1. Clone this repository
2. Create a Python virtual environment:
   ```bash
   python -m venv pst_scraper_env
   source pst_scraper_env/bin/activate  # On Windows use: pst_scraper_env\Scripts\activate
   ```
3. Install the required dependencies:
   ```bash
   pip install aspose-email pandas python-magic
   ```

## Usage

The tool processes PST files and generates two CSV files:
1. `emails.csv` - Contains all email messages with their metadata
2. `attachments.csv` - Contains information about email attachments

### Basic Usage

```python
from pst_scraper.pst_reader import read_pst

# Convert PST file to CSV
num_emails, num_attachments = read_pst(
    pst_file_path="path/to/your/file.pst",
    output_emails_path="emails.csv",
    output_attachments_path="attachments.csv"
)

print(f"Processed {num_emails} emails and {num_attachments} attachments")
```

### Output Format

#### emails.csv
The emails CSV file contains the following columns:
- `subject`: Email subject
- `conversation_topic`: Conversation topic
- `sender_email`: Sender's email address
- `sender_name`: Sender's display name
- `client_submit_time`: When the email was sent
- `delivery_time`: When the email was delivered
- `sensitivity`: Email sensitivity level (NORMAL, PERSONAL, PRIVATE, CONFIDENTIAL)
- `body_type`: Type of email body (TEXT, HTML, RTF)
- `body`: Email content
- `recipients`: Dictionary of recipients (to, cc, bcc)
- `attachment_ids`: List of attachment IDs
- `linked_message_ids`: List of linked message IDs

#### attachments.csv
The attachments CSV file contains:
- `name`: Attachment filename
- `type`: MIME type of the attachment
- `data`: Binary data of the attachment (if available)

## Notes

- The tool processes emails in batches of 50 for memory efficiency
- Encrypted messages are not supported and will raise an error
- Signed messages will have their signatures removed during processing
- The tool maintains the folder structure of the original PST file