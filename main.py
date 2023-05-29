import email
import os
import sys
import re
from datetime import datetime
import logging


def make_dir(folder_name):
    # Get the current date and time
    now = datetime.now()
    # Format the date and time as a string
    date_time = now.strftime("%Y-%m-%d-%H%M%S")
    # Check if the "attachments" directory exists
    if not os.path.isdir(folder_name):
    # Create the "attachments" directory if it does not exist
        os.makedirs(folder_name)
    
    # Create the "attachments_dateTime" directory
    os.chdir(folder_name)
    attachment_folder_name = folder_name+ " - " +date_time
    os.makedirs(attachment_folder_name)
    return attachment_folder_name

def check_extract_mode(sysArgv,prompt):
    """
    Prompts user for the path to the file or  the directory
    """
    while True:
        if (sysArgv is not None):
            path = sysArgv
        else:
            path = input(prompt).strip('"')

        if not path:
            print("Please enter a path.")
        else:
            return path

def sanitize_name(name):
    """
    This function removes invalid characters from the given file name and prints the attachment name
    """
    # Remove any invalid characters from the file name
    pattern = re.compile(r'[\\/*?:"<>|;\t]')
    name = pattern.sub('', name)
    # Remove periods from the beginning of file name
    name = re.sub(r'^\.', '', name)
    # Remove the newline character from the file name
    name = name.replace('\n', '')
    # Prints attachment Name
    print(f"\t\t{name}")
    return name

def process_eml(file_path,attachment_folder_name,folder_path):
    """
    This function processes a single EML file and saves any attachments found
    """
    # Initialize the counter variables
    attachment_count = 0
    attachment_saved_count = 0
    # Get the full file name including extension
    full_file_name= os.path.basename(file_path)
    # Get just the file name
    file_name=full_file_name.split('.')[0]
    # Open the .eml file
    try:
        with open(file_path, 'r') as fp:
            msg = email.message_from_file(fp)
    except PermissionError:
        logging.error(f"Error: {str(e)} - {file_path} Permission denied.\n")
        print(f'Permission Error on {full_file_name}')
    except FileNotFoundError:
        logging.error(f"Error: {str(e)} - {full_file_name} not found in {file_path}\n")
    except Exception as e:
        logging.error(f"Error: {str(e)} - {file_path} , {full_file_name}\n")
        print(f'An error occurred while opening {full_file_name}: {str(e)}')
    # Get the full directory of the .eml file
    eml_dir = os.path.dirname(file_path)
    if folder_path:
        # Creates a corresponding directory that is inside of the folder_path in the "attachments" folder
        save_dir = os.path.join(attachment_folder_name, eml_dir[len(folder_path)+1:])
        if not os.path.isdir(save_dir):
            os.makedirs(save_dir)
        # Create a subdirectory within the corresponding directory for the .eml file
        file_save_dir = os.path.join(save_dir, file_name)
    else:
        file_save_dir = os.path.join(attachment_folder_name,file_name)
    if not os.path.isdir(file_save_dir):
        os.makedirs(file_save_dir)
    print(f'{full_file_name} -\n\tAttachement -')
    # Iterate over all the attachments in the message
    for part in msg.walk():
        # If the attachment is a file, save it to the specified directory
        if part.get_content_maintype() == 'multipart':
            continue
        # Get the file header and Skips the attachment if it does not have a file header
        if part.get('Content-Disposition') is None:
            continue
        attachment_name = part.get_file_name()
        if attachment_name is None:
            continue
        attachment_name=str(attachment_name)
        attachment_count += 1
        attachment_name = sanitize_name(attachment_name)
        file_path = os.path.join(file_save_dir, attachment_name)
        try:
            with open(file_path, 'wb') as f:
                f.write(part.get_payload(decode=True))
            attachment_saved_count += 1
            # Print the number of attachments found and saved in particular folder
        except Exception as e:
            logging.error(f"'{attachment_name}' of '{file_name}' couldn't be saved :\n{e}")
            continue
    print(f'Found : {attachment_count}\nSaved : {attachment_saved_count}\n')
    return (attachment_count,attachment_saved_count)


def process_batch_eml(folder_path, attachment_folder_name):
    total_attachment_count = 0
    total_attachment_saved_count = 0
    try:
        # Iterate over all the subdirectories in the folder tree
        for root,dirs,files in os.walk(folder_path):
            # Iterate over the files in the current directory
            for file_name in files:
                # Check if the file is an .eml file
                if os.path.splitext(file_name)[1] == '.eml':
                    file_path= os.path.join(root,file_name)
                    attachment_count, attachment_saved_count = process_eml(file_path,attachment_folder_name,folder_path)
                    # Increment the total attachment and saved count
                    total_attachment_count += attachment_count
                    total_attachment_saved_count += attachment_saved_count
                else:
                    continue
    except Exception as e:
        logging.error(f"Error: {str(e)} - {file_name}\n")
        print("Error occured.Please check log.")
    return total_attachment_count,total_attachment_saved_count


if __name__=="__main__":
    # configure logging settings
    logging.basicConfig(file_name="attachment.log",level=logging.ERROR, format='%(asctime)s %(levelname)s: %(message)s')

    total_attachment_count = 0
    total_attachment_saved_count = 0

    # Get the folder path from the command line arguments
    if len(sys.argv) > 1:
        extract_mode = sys.argv[1].strip("-").lower()
        # Check extract mode
        if extract_mode in ["s","b"]:
            print("\n\tEML Extractor\n")
            if extract_mode == "s":
                if len(sys.argv) < 3:
                    file_path = check_extract_mode(None, "Please Enter File Path: ")
                else:
                    file_path = check_extract_mode(sys.argv[2].strip('"'), "Please Enter File Path: ")
                while(os.path.isfile(file_path) == False):
                    file_path = input("A directory path was provided instead of a file path. Please enter File Path: ").strip('"')
                attachment_folder_name = make_dir('attachment')
                print("\n")
                total_attachment_count, total_attachment_saved_count = process_eml(file_path, attachment_folder_name,None)
            
            else:
                if len(sys.argv) < 3:
                    folder_path = check_extract_mode(None, "Please Enter Folder Path:")
                else:
                    folder_path = check_extract_mode(sys.argv[2].strip('"'), "Please Enter Folder Path: ")
                while(os.path.isfile(folder_path) == True):
                    folder_path = input("A file path was provided instead of a folder path. Please enter Folder Path: ").strip('"')
                attachment_folder_name = make_dir('attachment')
                print("\n")
                total_attachment_count, total_attachment_saved_count = process_batch_eml(folder_path,attachment_folder_name)
        else:
            print("\nDid you mean -s or -b instead?\n")

    # Prompt the user to enter extract mode
    else:
        print("\n\tEML Extractor\n")
        print("1. Single EML\n2. Bulk EML")
        while True:
            try:
                extract_mode = int(input("\nChoose '1' or '2': "))
                if extract_mode in (1,2):
                    break
            except ValueError:
                pass
            print("Invalid Input.")
    if extract_mode == 1:
        file_path = check_extract_mode(None,"\nFile Path: ")
        while(os.path.isfile(file_path) == False):
            file_path = input("A directory path was provided instead of a file path. Please enter File Path: ").strip('"')
        attachment_folder_name = make_dir('attachment')
        print("\n")
        total_attachment_count, total_attachment_saved_count = process_eml(file_path, attachment_folder_name,None)
    elif extract_mode == 2:
        folder_path = check_extract_mode(None,"\nFolder Path: ")
        while(os.path.isfile(folder_path) == True):
            folder_path = input("A file path was provided instead of a folder path. Please enter Folder Path: ").strip('"')
        attachment_folder_name = make_dir('attachment')
        print("\n")
        total_attachment_count, total_attachment_saved_count = process_batch_eml(folder_path,attachment_folder_name)
    # Print the total number of attachments found and saved
    if total_attachment_saved_count and total_attachment_saved_count: 
        print(f'\nTotal : {total_attachment_count} attachments\nSaved : {total_attachment_saved_count} attachments\n')