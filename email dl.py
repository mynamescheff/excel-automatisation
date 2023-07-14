import win32com.client
from pathlib import Path

# Create output folder
output_dir = Path("directory path")
output_dir.mkdir(parents=True, exist_ok=True)

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Connect to folder
inbox = outlook.Folders("box name").Folders("Inbox")

# Get messages
messages = inbox.Items

# Filter messages by category
category_filter = "category name" 
filtered_messages = []

for message in messages:
    if message.Categories == category_filter:
        filtered_messages.append(message)

# Download non-image attachments from filtered messages
for message in filtered_messages:
    attachments = message.Attachments
    for attachment in attachments:
        file_extension = Path(attachment.FileName).suffix.lower()
        if file_extension not in (".jpg", ".jpeg", ".png", ".gif", ".bmp"):
            attachment.SaveAsFile(str(output_dir / attachment.FileName))
            print(f"Attachment saved: {attachment.FileName}")