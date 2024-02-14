# Email Automation with Python

This repository contains a Python script for automating the process of sending emails through Microsoft Outlook with specific Excel files attached as needed.

## Environment Setup

### Prerequisites

- Microsoft Outlook installed and configured on the computer that will run the script.
- Python installed.
- Python libraries `pandas` and `win32com.client`, which can be installed via pip:

```bash
pip install pandas pywin32
```

## File Structure

- Send Emails.xlsx: A spreadsheet containing the recipient data and the reports to be sent.
- Report Excel files: Files named according to the report areas (example: Financial.xlsx, Logistics.xlsx, etc.).

## Usage

### Initial Steps

1. Clone the repository or download the email.ipynb script.
2. Ensure all Excel report files are in the same folder as the script or update the file paths in the code accordingly.

### Editing the Data File

Edit the Send Emails.xlsx file according to the structure below:

- Manager: The name of the manager or report recipient.
- Email: The email address to which the report will be sent.
- Report: The name of the report area, which must match exactly the name of the Excel file for the report.

### Executing the Script

Run the email.ipynb Jupyter Notebook, and the script will automatically send the emails, following the information provided in the Send Emails.xlsx spreadsheet.

## Additional Notes

- Outlook must be open and functioning on the computer that executes the script.
- The current script is configured to work in a Windows environment, considering the use of Windows file paths and the win32com library.

## Support

For any questions or issues, please open an 'issue' in this repository or contact the maintainer.

Fabricio Colombo  
Software Engineer
