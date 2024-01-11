# Document Search Tool

This is a simple document search tool that creates a basic graphical user interface (GUI) using the `tkinter` library in Python. The tool allows users to select a document or folder and then recursively search for a specified keyword within the document content or files.

## Features

- Choose a document or folder for searching.
- Recursively search for a specified keyword in document content or files.
- Display search results and support opening files containing search results via a right-click menu.

## Prerequisites

- Python 3.x
- Tkinter library

## Usage

1. After running the program, click the "Select Document" button to choose the document to search or click the "Select Directory" button to choose the folder to search.
2. Enter a keyword in the search box.
3. Click the "Search" button to initiate the search.
4. Search results will be displayed in the text box, and you can open a file by right-clicking on the result.

## Supported File Types

- Currently supports searching in `.docx` (Word documents) and `.xlsx` (Excel documents) files.

## Dependencies

- Make sure to install the required Python libraries:
  ```bash
  pip install python-docx openpyxl
  
## How to Run

- To run the program, use the following command:
  ```bash
  python your_script_name.py
- Replace 'your_script_name.py' with the name of your Python script.

## Notes

- If searching Excel files, ensure the files are not open, as this may cause errors.
- This tool only supports documents in the '.docx' and '.xlsx' formats.

## Contributions
Feel free to raise issues, provide suggestions, or contribute code. Create an issue or submit a pull request.

## License
- This project is open-sourced under the MIT license.
- Remember to replace `your_script_name.py` with the name of your Python script.

  
