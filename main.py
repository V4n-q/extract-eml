import email
import os
import sys
import re
from datetime import datetime
import logging


def makeDir(folderName):
    # Get the current date and time
    now = datetime.now()
    # Format the date and time as a string
    dateTime = now.strftime("%Y-%m-%d-%H%M%S")
    # Check if the "attachments" directory exists
    if not os.path.isdir(folderName):
    # Create the "attachments" directory if it does not exist
        os.makedirs(folderName)
    
    # Create the "attachments_dateTime" directory
    os.chdir(folderName)
    attachFolderName = folderName+ " - " +dateTime
    os.makedirs(attachFolderName)
    return attachFolderName

def checkExtractMode(sysArgv,prompt):
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

def sanitizeName(name):
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

def processEml(filePath,attachFolderName,folderPath):
    """
    This function processes a single EML file and saves any attachments found
    """
    # Initialize the counter variables
    attachmentCount = 0
    savedCount = 0
    # Get the full file name including extension
    fullFileName= os.path.basename(filePath)
    # Get just the file name
    fileName=fullFileName.split('.')[0]
    # Open the .eml file
    try:
        with open(filePath, 'r') as fp:
            msg = email.message_from_file(fp)
    except PermissionError:
        logging.error(f"Error: {str(e)} - {filePath} Permission denied.\n")
        print(f'Permission Error on {fullFileName}')
    except FileNotFoundError:
        logging.error(f"Error: {str(e)} - {fullFileName} not found in {filePath}\n")
    except Exception as e:
        logging.error(f"Error: {str(e)} - {filePath} , {fullFileName}\n")
        print(f'An error occurred while opening {fullFileName}: {str(e)}')
    # Get the full directory of the .eml file
    emlDir = os.path.dirname(filePath)
    if folderPath:
        # Creates a corresponding directory that is inside of the folderPath in the "attachments" folder
        saveDir = os.path.join(attachFolderName, emlDir[len(folderPath)+1:])
        if not os.path.isdir(saveDir):
            os.makedirs(saveDir)
        # Create a subdirectory within the corresponding directory for the .eml file
        fileSaveDir = os.path.join(saveDir, fileName)
    else:
        fileSaveDir = os.path.join(attachFolderName,fileName)
    if not os.path.isdir(fileSaveDir):
        os.makedirs(fileSaveDir)
    print(f'{fullFileName} -\n\tAttachement -')
    # Iterate over all the attachments in the message
    for part in msg.walk():
        # If the attachment is a file, save it to the specified directory
        if part.get_content_maintype() == 'multipart':
            continue
        # Get the file header and Skips the attachment if it does not have a file header
        if part.get('Content-Disposition') is None:
            continue
        attachmentName = part.get_filename()
        if attachmentName is None:
            continue
        attachmentName=str(attachmentName)
        attachmentCount += 1
        attachmentName = sanitizeName(attachmentName)
        filePath = os.path.join(fileSaveDir, attachmentName)
        # print(f"{fileSaveDir}\{attachmentName}")
        try:
            with open(filePath, 'wb') as f:
                f.write(part.get_payload(decode=True))
            savedCount += 1
            # Print the number of attachments found and saved in particular folder
        except Exception as e:
            logging.error(f"'{attachmentName}' of '{fileName}' couldn't be saved :\n{e}")
            continue
    print(f'Found : {attachmentCount}\nSaved : {savedCount}\n')
    return (attachmentCount,savedCount)


def processBatchEml(folderPath, attachFolderName):
    totalAttachmentCount = 0
    totalSavedCount = 0
    try:
        # Iterate over all the subdirectories in the folder tree
        for root,dirs,files in os.walk(folderPath):
            # Iterate over the files in the current directory
            for fileName in files:
                # Check if the file is an .eml file
                if os.path.splitext(fileName)[1] == '.eml':
                    filePath= os.path.join(root,fileName)
                    attachmentCount, savedCount = processEml(filePath,attachFolderName,folderPath)
                    # Increment the total attachment and saved count
                    totalAttachmentCount += attachmentCount
                    totalSavedCount += savedCount
                else:
                    continue
    except Exception as e:
        logging.error(f"Error: {str(e)} - {fileName}\n")
        print("Error occured.Please check log.")
    return totalAttachmentCount,totalSavedCount


if __name__=="__main__":
    # configure logging settings
    logging.basicConfig(filename="attachment.log",level=logging.ERROR, format='%(asctime)s %(levelname)s: %(message)s')

    totalAttachmentCount = 0
    totalSavedCount = 0

    # Get the folder path from the command line arguments
    if len(sys.argv) > 1:
        extractMode = sys.argv[1].strip("-").lower()
        # Check extract mode
        if extractMode in ["s","b"]:
            print("\n\tEML Extractor\n")
            if extractMode == "s":
                if len(sys.argv) < 3:
                    filePath = checkExtractMode(None, "Please Enter File Path: ")
                else:
                    filePath = checkExtractMode(sys.argv[2].strip('"'), "Please Enter File Path: ")
                while(os.path.isfile(filePath) == False):
                    filePath = input("A directory path was provided instead of a file path. Please enter File Path: ").strip('"')
                attachFolderName = makeDir('attachment')
                print("\n")
                totalAttachmentCount, totalSavedCount = processEml(filePath, attachFolderName,None)
            
            else:
                if len(sys.argv) < 3:
                    folderPath = checkExtractMode(None, "Please Enter Folder Path:")
                else:
                    folderPath = checkExtractMode(sys.argv[2].strip('"'), "Please Enter Folder Path: ")
                while(os.path.isfile(folderPath) == True):
                    folderPath = input("A file path was provided instead of a folder path. Please enter Folder Path: ").strip('"')
                attachFolderName = makeDir('attachment')
                print("\n")
                totalAttachmentCount, totalSavedCount = processBatchEml(folderPath,attachFolderName)
        else:
            print("\nDid you mean -s or -b instead?\n")

    # Prompt the user to enter extract mode
    else:
        print("\n\tEML Extractor\n")
        print("1. Single EML\n2. Bulk EML")
        while True:
            try:
                extractMode = int(input("\nChoose '1' or '2': "))
                if extractMode in (1,2):
                    break
            except ValueError:
                pass
            print("Invalid Input.")
    if extractMode == 1:
        filePath = checkExtractMode(None,"\nFile Path: ")
        while(os.path.isfile(filePath) == False):
            filePath = input("A directory path was provided instead of a file path. Please enter File Path: ").strip('"')
        attachFolderName = makeDir('attachment')
        print("\n")
        totalAttachmentCount, totalSavedCount = processEml(filePath, attachFolderName,None)
    elif extractMode == 2:
        folderPath = checkExtractMode(None,"\nFolder Path: ")
        while(os.path.isfile(folderPath) == True):
            folderPath = input("A file path was provided instead of a folder path. Please enter Folder Path: ").strip('"')
        attachFolderName = makeDir('attachment')
        print("\n")
        totalAttachmentCount, totalSavedCount = processBatchEml(folderPath,attachFolderName)
    # Print the total number of attachments found and saved
    if totalSavedCount and totalSavedCount: 
        print(f'\nTotal : {totalAttachmentCount} attachments\nSaved : {totalSavedCount} attachments\n')