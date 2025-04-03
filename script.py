import datetime
import re
import sys
from pathlib import Path
import win32com.client # type: ignore
# Ensure constants are available
try:
    constants = win32com.client.constants
except AttributeError:
    try:
        win32com.client.gencache.EnsureModule('{00062FFF-0000-0000-C000-000000000046}', 0, 9, 6)
        constants = win32com.client.constants
        print("Note: Had to regenerate win32com constants cache.")
    except Exception as e_cache:
        print(f"FATAL: Failed to load win32com constants. Error: {e_cache}")
        sys.exit(1)

def sanitize_filename(filename):
    """Removes invalid characters for Windows filenames and limits length."""
    # Remove characters invalid in Windows filenames
    sanitized = re.sub(r'[<>:"/\\|?*]+', '_', filename)
    # Remove control characters
    sanitized = re.sub(r'[\x00-\x1f\x7f]', '', sanitized)
    # Replace leading/trailing spaces or dots
    sanitized = sanitized.strip('. ')
    # Limit length to avoid issues (e.g., 150 chars)
    max_len = 150
    if len(sanitized) > max_len:
        name_part, dot, extension = sanitized.rpartition('.')
        if dot:
            name_part = name_part[:max_len - len(extension) - 1]
            sanitized = f"{name_part}.{extension}"
        else:
            sanitized = sanitized[:max_len]
    if not sanitized:
        return "Invalid_Name"
    return sanitized

def create_directory(path_obj):
    """Creates a directory if it doesn't exist. Returns Path object or None."""
    try:
        path_obj.mkdir(parents=True, exist_ok=True)
        return path_obj
    except OSError as e:
        print(f"Error creating directory {path_obj}: {str(e)}")
        return None
    except Exception as e:
        print(f"Unexpected error creating directory {path_obj}: {str(e)}")
        return None

def list_folders_recursive(folder, prefix="", all_folders_list=None):
    """Recursively lists folders and appends (display_name, folder_object) to a list."""
    if all_folders_list is None:
        all_folders_list = []

    current_display_name = f"{prefix}{folder.Name}" if prefix else folder.Name
    all_folders_list.append((current_display_name, folder))

    try:
        subfolders = folder.Folders
        if subfolders.Count > 0:
            new_prefix = f"{current_display_name}/"
            for sub_folder in subfolders:
                list_folders_recursive(sub_folder, new_prefix, all_folders_list)
    except Exception as e:
        print(f"Warning: Could not access subfolders of '{folder.Name}'. Error: {e}")

    return all_folders_list


def export_emails_as_msg(folder, output_dir_base):
    """
    Exports emails from a given Outlook folder as .msg files,
    skipping if a file with the same name already exists.
    Returns: (processed_count, errors_list, skipped_non_mail, skipped_duplicates)
    """
    processed_count = 0
    errors = []
    skipped_non_mail = 0
    skipped_duplicates = 0 # <-- New counter

    # Determine output directory based on folder path
    try:
        full_folder_path = folder.FolderPath
        path_parts = full_folder_path.split('\\')
        if len(path_parts) > 3:
             relative_folder_path = Path(*path_parts[3:])
             output_dir = create_directory(output_dir_base / relative_folder_path)
        else:
            # Sanitize top-level folder name just in case
            safe_folder_name = sanitize_filename(folder.Name)
            output_dir = create_directory(output_dir_base / safe_folder_name)

        if not output_dir:
             errors.append(f"Failed to create output directory for folder '{folder.Name}'. Skipping folder.")
             # Return 4 values now
             return 0, errors, 0, 0

    except Exception as e:
        errors.append(f"Error determining output path for folder '{folder.Name}': {e}. Skipping folder.")
        # Return 4 values now
        return 0, errors, 0, 0

    print(f"  Exporting to: {output_dir}")

    try:
        messages = folder.Items
    except Exception as e:
        errors.append(f"Could not retrieve items from folder '{folder.Name}': {e}")
         # Return 4 values now
        return 0, errors, 0, 0

    total_items = len(messages) # Get total count for progress
    print(f"    Found {total_items} items in '{folder.Name}'.")

    for i, message in enumerate(messages):
        # Progress indicator
        if (i + 1) % 100 == 0: # Changed to 100 for less frequent updates
             print(f"    Processed {i+1}/{total_items} items...")

        # Check if it's a mail item
        try:
            if not hasattr(message, 'Class') or message.Class != constants.olMail:
                skipped_non_mail += 1
                continue
        except Exception as e:
             errors.append(f"Error checking item type at index {i}: {e}. Skipping item.")
             skipped_non_mail += 1
             continue

        try:
            # Get timestamp for filename
            received_time_obj = getattr(message, 'ReceivedTime', None)
            if received_time_obj:
                 time_str = received_time_obj.strftime("%Y-%m-%d_%H-%M-%S")
            else:
                time_str = "UnknownTime"

            # Get subject for filename
            subject = getattr(message, 'Subject', 'No Subject')
            safe_subject = sanitize_filename(subject)

            # Construct filename and full path object
            filename = f"{time_str}_{safe_subject}.msg"
            full_path_obj = output_dir / filename

            # --- Check for Duplicates ---
            if full_path_obj.exists():
                skipped_duplicates += 1
                continue # Skip to the next message
            # --- End Check ---

            # Save as .msg file (convert path object to string for SaveAs)
            message.SaveAs(str(full_path_obj), constants.olMSG)
            processed_count += 1

        except Exception as e:
            error_subject = getattr(message, 'Subject', 'Unknown Subject')
            error_time = time_str
            errors.append(f"Error saving item '{error_subject}' (Time: {error_time}): {str(e)}")

    print(f"  Finished folder '{folder.Name}'. Exported: {processed_count}, Skipped Duplicates: {skipped_duplicates}, Errors: {len(errors)}, Skipped non-mail: {skipped_non_mail}")
    # Return 4 values now
    return processed_count, errors, skipped_non_mail, skipped_duplicates

# --- Main Execution ---
def main():
    print("Starting Outlook Email Exporter...")

    # 1. Get Output Path (Code is the same as before)
    while True:
        output_path_str = input("Enter the full path for email exports (e.g., E:\\EmailBackups): ")
        if not output_path_str:
            print("Path cannot be empty. Please try again.")
            continue
        try:
            base_dir = Path(output_path_str)
            if not create_directory(base_dir):
                 raise OSError(f"Failed to create or access base directory: {base_dir}")
            print(f"Using base export directory: {base_dir}")
            break
        except OSError as e:
            print(f"Error: {e}")
            print("Please check the path, ensure the drive/folder exists, and you have write permissions.")
        except Exception as e:
            print(f"An unexpected error occurred with the path: {e}")
        retry = input("Try entering the path again? (y/n): ").lower()
        if retry != 'y':
            print("Exiting script.")
            return

    # 2. Connect to Outlook (Code is the same as before)
    try:
        print("Connecting to Outlook...")
        outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print("Connected to Outlook.")
    except Exception as e:
        print(f"FATAL: Failed to connect to Outlook: {e}")
        return

    # 3. Select Store (Code is the same as before)
    stores = namespace.Folders
    print("\nAvailable Mailboxes/Stores:")
    if not stores or stores.Count == 0:
         print("FATAL: No mailboxes/stores found.")
         return
    for i, store in enumerate(stores):
        print(f"  {i + 1}: {store.Name}")
    selected_store = None
    while selected_store is None:
        try:
            choice = input(f"Select the number of the mailbox to export from (1-{stores.Count}): ")
            store_index = int(choice) - 1
            if 0 <= store_index < stores.Count:
                selected_store = stores[store_index]
                print(f"Selected mailbox: {selected_store.Name}")
            else:
                print("Invalid number selected.")
        except ValueError:
            print("Invalid input. Please enter a number.")
        except Exception as e:
             print(f"An error occurred selecting the store: {e}")

    safe_store_name = sanitize_filename(selected_store.Name)
    store_output_base = create_directory(base_dir / safe_store_name)
    if not store_output_base:
         print(f"FATAL: Could not create directory for store {safe_store_name}. Exiting.")
         return

    # 4. List and Select Folders (Code is the same as before)
    print(f"\nListing folders in '{selected_store.Name}' (this may take a moment)...")
    all_folders = []
    try:
        for top_level_folder in selected_store.Folders:
             list_folders_recursive(top_level_folder, prefix="", all_folders_list=all_folders)
    except Exception as e:
        print(f"Error listing folders: {e}")
        return
    if not all_folders:
        print("No folders found in the selected mailbox.")
        return
    print("\nAvailable Folders:")
    for i, (display_name, _) in enumerate(all_folders):
        print(f"  {i + 1}: {display_name}")
    selected_folders_to_export = []
    while True:
        try:
            choices_str = input("Enter the numbers of the folders to export, separated by commas (e.g., 1, 3, 15), or 'all': ").strip()
            if choices_str.lower() == 'all':
                selected_folders_to_export = [f_obj for _, f_obj in all_folders]
                print("Selected all folders.")
                break
            elif not choices_str:
                 print("No folders selected. Exiting.")
                 return
            indices = [int(x.strip()) - 1 for x in choices_str.split(',')]
            valid_indices = True
            temp_selected_folders = []
            selected_names = []
            for index in indices:
                if 0 <= index < len(all_folders):
                    display_name, folder_obj = all_folders[index]
                    temp_selected_folders.append(folder_obj)
                    selected_names.append(display_name)
                else:
                    print(f"Invalid folder number: {index + 1}")
                    valid_indices = False
                    break
            if valid_indices:
                 selected_folders_to_export = temp_selected_folders
                 print("\nSelected folders:")
                 for name in selected_names:
                      print(f"- {name}")
                 confirm = input("Confirm selection? (y/n): ").lower()
                 if confirm == 'y':
                      break
                 else:
                      print("Selection cancelled. Please re-enter folder numbers.")
                      selected_folders_to_export = []
            else:
                 selected_folders_to_export = []
        except ValueError:
            print("Invalid input. Please enter numbers separated by commas, or 'all'.")
        except Exception as e:
            print(f"An error occurred processing selection: {e}")

    # 5. Process Selected Folders
    print("\n--- Starting Export ---")
    total_processed = 0
    total_errors = []
    total_skipped_non_mail = 0
    total_skipped_duplicates = 0 # <-- New counter

    for folder_obj in selected_folders_to_export:
        # Use display name for clarity if available, otherwise folder name
        folder_display_name = folder_obj.Name # Fallback
        for d_name, f_obj in all_folders:
             if f_obj == folder_obj:
                  folder_display_name = d_name
                  break
        print(f"\nProcessing folder: {folder_display_name}")

        # Update to receive 4 values
        processed, errors, skipped_non_mail, skipped_duplicates = export_emails_as_msg(folder_obj, store_output_base)
        total_processed += processed
        total_errors.extend(errors)
        total_skipped_non_mail += skipped_non_mail
        total_skipped_duplicates += skipped_duplicates # <-- Accumulate count

    # 6. Print Summary (Updated)
    print("\n--- Export Complete ---")
    print(f"Total emails exported as .msg: {total_processed}")
    print(f"Total duplicates skipped (file already exists): {total_skipped_duplicates}") # <-- New summary line
    print(f"Total non-mail items skipped: {total_skipped_non_mail}")
    if total_errors:
        print(f"\nEncountered {len(total_errors)} errors during export:")
        max_errors_to_show = 20
        for i, error in enumerate(total_errors):
            if i >= max_errors_to_show:
                print(f"... (and {len(total_errors) - max_errors_to_show} more errors)")
                break
            print(f"- {error}")
    else:
        print("No errors encountered during export.")

    print(f"\nExport location: {store_output_base}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
    except Exception as e:
        print("\n--- An Unexpected Error Occurred ---")
        import traceback
        print(traceback.format_exc())
        print(f"Error: {e}")
    finally:
        print("\nScript finished.")