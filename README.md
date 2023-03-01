# extractEML

Python Script to extract attachments from an EML File.
Can extract from specific directory and its subdirectories. The extracted attachments are saved inside "attachments - {date-time}" and it maintains directory structure that the EML Files are in, with the name in the same directory as the script.

# Requirements

All of the imported modules are preinstalled.

# Usage

- `py main.py`

  Select Extraction Type when prompted and enter the file path or directory path.

<center>Or,</center>

- `py main.py "-ExtractionType"`


## Extraction Type

- Single (`-s`)
- Bulk (`-b`)

### Single File :

  - `py main.py -s`
    
    Enter file path when prompted.

  Or,

  - `py main.py -s "filePath"`

    "filePath" is location of EML File

    eg: `py main.py -s "D:\EMLFiles\File1\test.eml"`

### Bulk File :

  - `py main.py -b` 

    Enter folder path when prompted.

  Or,

  - `py main.py -b "folderPath"`
  
    "folderPath" is location of where all the EML File is located.

    eg: `py main.py -b "D:\EMLFiles\"`


## Directory Structure

If the folder path provides is `D:\EMLFiles`, the script considered `EMLFiles` as root directory and maintains similar structure and keeps attachments in it.

Original Directory Structure:

```
EMLFiles
    |-> File1
    |     |-> testFile1.eml 
    |
    |-> File2     
          |-> testFile2.eml
```

Attachment Directory Structure:

```
attachment
    |-> attachment - {date-time}
          |-> File1
          |   |-> testFile1
          |         |-> xyz.pdf
          |         |-> abc.txt 
          |
          |-> File2     
              |-> testFile2
                    |-> rfg.xlsx
```
