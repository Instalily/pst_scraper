# PST to CSV Converter

A Python utility that converts Microsoft Outlook PST files into structured CSV documents. This tool extracts emails, attachments, and linked messages from PST files and organizes them into separate CSV files for easy analysis and processing.

## Features

- Extracts emails from PST files with full metadata
- Handles email attachments and linked messages
- Outputs data to multiple CSV tables for easy importing into SQL databases
- Processes folders recursively, maintaining PST folder structure
- Batched processing for memory efficiency
- Supports processing multiple PST files in a single run
- Saves email attachments to a specified directory with unique numeric identifiers

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

The tool processes PST files and generates four CSV files and an attachments directory:

1. `emails.csv` - Contains all email messages with their metadata
2. `attachments.csv` - Contains information about email attachments
3. `accounts.csv` - Contains information about email accounts
4. `emails_to_recipients.csv` - Contains the relationships between emails and their recipients
5. An attachments directory - Contains the actual attachment files

### Basic Usage

```python
from pst_scraper.pst_reader import read_psts

# Convert multiple PST files to CSV
num_emails, num_attachments = read_psts(
    pst_file_paths=["path/to/file1.pst", "path/to/file2.pst"],
    emails_csv_path="emails.csv",
    attachments_csv_path="attachments.csv",
    accounts_csv_path="accounts.csv",
    emails_to_recipients_csv_path="emails_to_recipients.csv",
    attachments_dir="attachments"  # Directory where attachment files will be saved
)

print(f"Processed {num_emails} emails and {num_attachments} attachments")
```

### Output Format

#### emails.csv
The emails CSV file contains the following columns:
- `id`: Unique identifier for the email
- `subject`: Email subject
- `conversation_topic`: Conversation topic
- `sender`: Sender's email address
- `client_submit_time`: When the email was sent
- `delivery_time`: When the email was delivered
- `sensitivity`: Email sensitivity level (NORMAL, PERSONAL, PRIVATE, CONFIDENTIAL)
- `body_type`: Type of email body (TEXT, HTML, RTF)
- `body`: Email content (with special CSV formatting)
- `linked_from`: ID of the email this message is linked from, or -1 if not linked

### Body Text Formatting

The email body text undergoes several transformations to ensure proper CSV formatting:

1. **Quote Escaping**: All double quotes (`"`) in the body are escaped by doubling them (`""`). This is standard CSV escaping for quotes.
   - Example: `"Hello"` becomes `""Hello""`

2. **Newline Handling**: All newlines (`\n`) and carriage returns (`\r`) are replaced with their escaped versions (`\\n` and `\\r` respectively).
   - Example: `Hello\nWorld` becomes `Hello\\nWorld`

To restore the original text format when reading the CSV:
```python
# To restore the original text:
body = body.replace('\\n', '\n').replace('\\r', '\r').replace('""', '"')
```

These transformations ensure that:
- The body appears as a single field in the CSV
- Special characters are properly escaped
- The content remains readable while being CSV-compatible
- The original text can be perfectly restored

#### attachments.csv
The attachments CSV file contains:
- `id`: Unique identifier for the attachment
- `name`: Original display name of the attachment
- `path`: Path to the attachment file in the attachments directory
- `email_id`: ID of the email this attachment belongs to

#### accounts.csv
The accounts CSV file contains:
- `email`: Email address (lowercase)
- `display_name`: Display name of the account

#### emails_to_recipients.csv
The emails to recipients CSV file contains:
- `email_id`: ID of the email
- `account_id`: Email address of the recipient
- `recipient_type`: Type of recipient (TO, CC, BCC)

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
- Email addresses are stored in lowercase to ensure consistent matching
- The body field in emails.csv uses CSV-compatible escaping for quotes and newlines (see Body Text Formatting section)
- All transformations to the body text are reversible using the provided restoration code