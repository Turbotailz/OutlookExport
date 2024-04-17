import datetime
from pathlib import Path
import re
import win32com.client
from win32com.client import gencache

def create_directory(base_path, folder_name):
    """ Helper function to create a directory for storing email exports.
    
    Args:
        base_path (Path): The base directory where the folder will be created.
        folder_name (str): The name of the folder to be created.
    
    Returns:
        Path: The path to the newly created directory or None if an error occurred.
    """
    try:
        # Create the directory and any required parent directories
        new_dir = (base_path / folder_name)
        new_dir.mkdir(parents=True, exist_ok=True)
        return new_dir
    except Exception as e:
        # Print error message if the directory creation fails
        print(f"Error creating directory {folder_name}: {str(e)}")
        return None

def save_attachments(attachments, folder_path):
    """ Saves all email attachments to a specified directory.
    
    Args:
        attachments (Attachments): A collection of attachments from an email.
        folder_path (Path): The directory path where attachments should be saved.
    """
    for attachment in attachments:
        try:
            # Sanitize the filename to remove any invalid characters
            attachment_name = re.sub(r'[^\w\s.]+', '', attachment.FileName)
            # Save the attachment to the specified folder
            attachment.SaveAsFile(str(folder_path / attachment_name))
        except Exception as e:
            # Print error message if saving the attachment fails
            print(f"Failed to save attachment {attachment.FileName}: {str(e)}")

def export_emails(folder, output_dir):
    """ Processes and exports emails from a given Outlook folder.
    
    Args:
        folder (Folder): The Outlook folder from which to export emails.
        output_dir (Path): The directory path where emails should be exported.
    
    Returns:
        tuple: A tuple containing the number of processed emails and a list of error messages.
    """
    messages = folder.Items
    # Attempt to sort the messages by ReceivedTime, if applicable
    if folder.DefaultItemType == win32com.client.constants.olMailItem:
        try:
            messages.Sort("[ReceivedTime]", True)
        except Exception as e:
            print(f"Could not sort items in folder {folder.Name}: {str(e)}")

    processed_count = 0
    errors_and_skips = []

    for message in messages:
        email_time = "Unknown"
        try:
            if not hasattr(message, 'ReceivedTime'):
                errors_and_skips.append("Skipped non-mail item with unknown received time.")
                continue
            if message.Class != win32com.client.constants.olMail:
                errors_and_skips.append(f"Skipped non-mail item at {message.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')}")
                continue

            # Format the time stamp for folder naming
            email_time = message.ReceivedTime.strftime("%Y-%m-%d_%H-%M-%S")
            folder_name = f"{email_time}"
            email_folder = create_directory(output_dir, folder_name)

            if email_folder:
                # Extract subject and body, handling missing attributes
                subject = getattr(message, 'Subject', 'No Subject')
                body = getattr(message, 'Body', 'No content available')
                email_content = f"Subject: {subject}\n\n{body}"
                # Write the email content to a text file in the created directory
                with open(email_folder / "email_body.txt", 'w', encoding='utf-8') as file:
                    file.write(email_content)
                # Save attachments if present
                if message.Attachments.Count > 0:
                    save_attachments(message.Attachments, email_folder)
            processed_count += 1
        except Exception as e:
            errors_and_skips.append(f"Error with item at {email_time}: {str(e)}")

    print(f"Successfully processed {processed_count} emails in folder {folder.Name}.")
    return processed_count, errors_and_skips

def main():
    """ Main function to initialize Outlook access and process each folder. """
    print("Starting the script...")
    # Initialize Outlook Application
    outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    base_dir = Path.cwd() / "EmailExports"
    base_dir.mkdir(parents=True, exist_ok=True)
    all_folders = namespace.Folders

    total_processed = 0
    all_errors_and_skips = []

    # Process each store and its folders
    for store in all_folders:
        print(f"Processing store: {store.Name}")
        store_dir = create_directory(base_dir, store.Name)
        folders = store.Folders
        for folder in folders:
            print(f"Processing folder: {folder.Name}")
            folder_dir = create_directory(store_dir, folder.Name)
            processed_count, errors_and_skips = export_emails(folder, folder_dir)
            total_processed += processed_count
            all_errors_and_skips.extend(errors_and_skips)

    # Print summary of all errors and skips
    if all_errors_and_skips:
        print("Summary of errors and skipped emails:")
        for error in all_errors_and_skips:
            print(error)
    else:
        print("No errors or skipped emails.")

    print(f"Total emails processed: {total_processed}")
    print("Email export completed.")

if __name__ == "__main__":
    main()
