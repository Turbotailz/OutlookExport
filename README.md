# OutlookExport
This Python script uses `pywin32` to export emails and attachments from Outlook into organized directories on your computer. It saves emails by received time and includes subject, body, and attachments, ideal for email backup and archiving.

This Python script is designed to automatically export emails and their attachments from Microsoft Outlook into a structured directory system on your local file system. It utilizes the `pywin32` library to interface with Outlook, making it possible to automate interactions with the email client. Here's a detailed breakdown of its functionalities and structure:

### Script Components and Workflow:

1. **Importing Required Modules**:
   - `datetime`: For handling date and time objects.
   - `pathlib.Path`: For filesystem paths operations, creating and managing directory paths.
   - `re`: Regular expression module for cleaning file names.
   - `win32com.client`: To interact with COM objects, in this case, Microsoft Outlook.

2. **Utility Functions**:
   - `create_directory(base_path, folder_name)`: Creates a directory at the specified path and handles errors related to directory creation.
   - `save_attachments(attachments, folder_path)`: Iterates through email attachments and saves them to a specified directory, sanitizing the filenames to remove invalid characters.

3. **Main Functionality - `export_emails(folder, output_dir)`**:
   - Accesses the items in a specified Outlook folder.
   - Sorts these items by "ReceivedTime" if they are email items, avoiding errors for other item types like calendar entries.
   - Processes each item: If it's an email, it creates a directory named after the email's received time, saves the email's body and subject to a text file, and saves any attachments.
   - Skips non-email items and logs any errors or skipped items.

4. **Main Execution Function - `main()`**:
   - Starts by initializing the Outlook application interface and accessing the default namespace.
   - Sets up the base directory for email exports (`EmailExports` within the current working directory).
   - Iterates through each store (email account) and its folders within the Outlook application, processing each folder by exporting its contents using the `export_emails()` function.
   - Collects and prints a summary of all processed emails and any errors or skipped operations.

5. **Error Handling**:
   - Comprehensive throughout the script to handle and log exceptions that may occur when accessing Outlook properties, file IO operations, or COM interactions.

6. **Output**:
   - Emails are organized in directories named after their received timestamp within folders corresponding to their original Outlook folders.
   - Each email's content is saved in a text file named `email_body.txt`, and attachments are saved in their original formats within the same directory.

### Use Case:
This script is particularly useful for backing up emails from Outlook into a local, searchable file system format. It's also beneficial for processing large quantities of emails for data extraction, archiving, or migration purposes.

### Running the Script:
To run this script, you need:
- Python installed on your system.
- The `pywin32` module installed.
- Microsoft Outlook installed and configured with at least one email account.
- Appropriate permissions to interact with Outlook and the filesystem.

This script is an example of how automation can be used to simplify routine data management tasks, providing a robust solution for exporting and archiving email data from Outlook.
