# Excel2PDF

## `NOTE`: This is a small yet frequently used utility developed for one of our clients, therefore there are certain limits of this code like this will generate the docuements some fixed columns named .xlsx file. It plays an important role in their daily operations, providing a clear and practical example of how streamlined automation solutions are implemented in real-world industry environments.
 

A utility application for generating No Objection Certificate (NOC) documents from Excel data.

## Overview

Excel2PDF is a desktop application that automates the creation of NOC documents for vehicle hypothecation. It reads data from an Excel spreadsheet and populates a template Word document (sample.docx) with the relevant information, saving the customized document in an outputs folder.

## Features

- Automatically generates NOC documents from Excel data
- Populates fields such as file number, customer name, vehicle details, etc.
- Formats text with appropriate styling (bold)
- Saves output files with file number as filename
- Prevents file conflicts with error handling for open documents

## Requirements

- Python 3.6 or higher
- Required Python packages (see `requirements.txt`):
  - python-docx
  - docxedit
  - pandas
  - openpyxl
  - tkinter (included with standard Python)

## Installation

### Option 1: Using the Executable

1. Download the `.exe` file
2. Place it in a folder with the `sample.docx` template
3. Create an `outputs` folder in the same directory
4. Run the application

### Option 2: From Source

1. Clone or download this repository
2. Install required packages:
   ```
   pip install -r requirements.txt
   ```
3. Ensure `sample.docx` is in the project directory
4. Create an `outputs` folder in the project directory
5. Run `main.py`

## Usage

1. Prepare your Excel data with the following columns:
   - File Number
   - Customer Name
   - Vehicle Number
   - Chassis Number
   - Engine Number
   - Date of Closing

2. Launch the application

3. The program will:
   - Read data from your Excel file
   - Generate NOC documents based on the sample.docx template
   - Save completed documents to the outputs folder with filenames based on file numbers
   - Display confirmation messages upon successful generation

## File Structure

```
Excel2PDF/
├── main.py                 # Main application code
├── sample.docx             # Template document
├── requirements.txt        # Required Python packages
├── Excel2PDFIcon.ico       # Application icon
└── outputs/                # Folder for generated documents
```

## Building the Executable

The application can be compiled into a standalone executable using PyInstaller:

```
pyinstaller --onefile --name Excel2PDF --hidden-import=pandas --hidden-import=docxedit --hidden-import=openpyxl --add-data "sample.docx;." --icon=Excel2PDFIcon.ico --noconsole main.py
```

## How It Works

The application:
1. Opens the sample.docx template
2. Updates the date field with the current date
3. Modifies the subject line with account and vehicle numbers
4. Populates the table with:
   - File number
   - Customer name
   - Vehicle number
   - Chassis number
   - Engine number
   - Date of closing
5. Saves the document to the outputs folder with the file number as the filename
6. Handles errors if a file is already open

## Troubleshooting

- If you see an error message about an open NOC file, close the mentioned file and try again
- Ensure the `outputs` directory exists in the same location as the application
- Make sure `sample.docx` is properly formatted and available in the application directory

## Contact

[Siddharth Roy](https://www.linkedin.com/in/siddharth--roy/)

