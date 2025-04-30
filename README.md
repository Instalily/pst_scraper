# PST to CSV Converter

A Python utility that converts Microsoft Outlook PST files into structured CSV documents. This tool extracts emails, attachments, and linked messages from PST files and organizes them into separate CSV files for easy analysis and processing.

## Features

- Extracts emails from PST files with full metadata
- Handles email attachments and linked messages
- Outputs data to CSV format for easy importing into databases or analysis tools
- Processes folders recursively, maintaining PST folder structure
- Batched processing for memory efficiency
- Supports processing multiple PST files in a single run
- Saves email attachments to a specified directory with unique numeric identifiers
- Appends to existing output files if they exist
- Supports both file paths and byte streams as input

## Installation

1. Clone this repository
2. Create a Python virtual environment (Python 3.12 or higher required):
   ```bash
   python -m venv pst_scraper_env
   source pst_scraper_env/bin/activate  # On Windows use: pst_scraper_env\Scripts\activate
   ```
3. Install the package and its dependencies:
   ```bash
   pip install -e .
   ```
   For development, install with additional tools:
   ```bash
   pip install -e ".[dev]"
   ```

## Usage

The tool processes PST files and generates:
1. `emails.csv` - Contains all email messages with their metadata
2. `attachments.csv` - Contains information about email attachments
3. An attachments directory - Contains the actual attachment files

### Basic Usage

```python
from pst_scraper.pst_reader import read_psts

# Convert multiple PST files to CSV using file paths
num_emails, num_attachments = read_psts(
    pst_files=["path/to/file1.pst", "path/to/file2.pst"],
    emails_csv_path="emails.csv",
    attachments_csv_path="attachments.csv",
    attachments_dir="attachments"  # Directory where attachment files will be saved
)

# Or using byte streams
with open("path/to/file.pst", "rb") as f:
    pst_data = f.read()
num_emails, num_attachments = read_psts(
    pst_files=[pst_data],  # Can mix file paths and byte streams
    emails_csv_path="emails.csv",
    attachments_csv_path="attachments.csv",
    attachments_dir="attachments"
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
- `to`: List of "To" recipients
- `cc`: List of "CC" recipients
- `bcc`: List of "BCC" recipients
- `attachment_ids`: List of attachment IDs (corresponds to file names in attachments directory)
- `linked_message_ids`: List of linked message IDs

#### attachments.csv
The attachments CSV file contains:
- `name`: Original display name of the attachment in the email
- The file for each attachment is saved in the attachments directory with a name matching its ID number

#### Attachments Directory
- Each attachment is saved as a separate file
- File names are numeric IDs (e.g., "0", "1", "2", etc.)
- The mapping between attachment IDs and their original names is maintained in the attachments.csv file

## Notes

- The tool processes emails in batches of 50 for memory efficiency
- Encrypted messages are not supported and will raise an error
- Signed messages will have their signatures removed during processing
- The tool maintains the folder structure of the original PST file
- When processing multiple PST files, the output CSV files will contain all emails and attachments from all input files
- Attachment files are saved with numeric IDs to avoid filename conflicts
- If the output CSV files already exist, new data will be appended to them
- The tool will automatically create the attachments directory if it doesn't exist
- You can mix file paths and byte streams in the same call to `read_psts`