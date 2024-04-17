This script facilitates the automated export of emails and attachments from Microsoft Outlook into a structured directory system on your Windows computer, leveraging the `pywin32` library.

### Script Description
This Python script automates the process of exporting emails and attachments from Microsoft Outlook into organized directories on your Windows system. It utilizes the `pywin32` library to interact with Outlook, managing emails efficiently by sorting, saving, and organizing them into directories named after their received timestamps. Non-mail items are handled appropriately, and errors are robustly managed.

### Key Features and Workflow:
1. **Modules Used**:
   - `datetime` for handling date and time.
   - `pathlib` for file system path operations.
   - `re` for sanitizing filenames using regular expressions.
   - `win32com.client` for interacting with COM objects, specifically Microsoft Outlook.

2. **Utility Functions**:
   - `create_directory(base_path, folder_name)`: Manages directory creation.
   - `save_attachments(attachments, folder_path)`: Saves email attachments to specified directories after cleaning filenames.

3. **Email Processing**:
   - Accesses items in specified Outlook folders, sorting emails by "ReceivedTime".
   - Processes each email, creating directories named by email timestamp, saving subject and body to text files, and saving attachments.

4. **Execution and Error Handling**:
   - Script initializes Outlook application access, processes all folders within each email account, and logs any errors or skips.

5. **Output**:
   - Organizes emails in directories based on timestamps, includes both text files for email content and original formats for attachments.

### System Requirements and Installation Guide:
- **Operating System Requirement**: Designed specifically for Windows OS due to the dependence on `pywin32` and Outlook.

- **Python Installation**:
  1. **Download Python**:
     - Go to the [Python Releases for Windows](https://www.python.org/downloads/windows/) page on Python's official website.
     - Click on "Download Windows installer".
  2. **Install Python**:
     - Run the downloaded installer.
     - Make sure to check "Add Python 3.x to PATH" at the bottom of the installation window to automatically add Python to your environment variables.
     - Click "Install Now".

- **Adding Python and pip to PATH Manually**:
  If you didn’t add Python to your PATH during the installation, you can add it manually:
  1. **Open the Start Search**, type `env`, and select "Edit the system environment variables" or "Edit environment variables for your account".
  2. **Under System Properties**, click on the "Environment Variables…" button.
  3. **Find the 'Path' variable** in the "System variables" section and click "Edit…".
  4. **Add Python path**:
     - Click "New" and add the path to the folder where Python is installed, e.g., `C:\Users\<Username>\AppData\Local\Programs\Python\Python39`.
     - Add another new line for the `Scripts` directory, e.g., `C:\Users\<Username>\AppData\Local\Programs\Python\Python39\Scripts`.
  5. **Click OK** on all dialogs to close them.

- **Install `pywin32`**:
  ```bash
  pip install pywin32
  ```

- **Running the Script**:
  - Ensure Microsoft Outlook is installed and configured with your email account.
  - Run the script with administrative privileges to enable necessary permissions for accessing Outlook and performing file operations.

This setup ensures that the script can be run efficiently on any compatible Windows machine, providing a robust tool for exporting and archiving email data from Microsoft Outlook.
