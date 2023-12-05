import win32com.client as win32
import schedule
import time
import tkinter as tk
from tkinter import filedialog

# Create a file dialog to select the Excel file
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(title="Select Excel File")

# Connect to Outlook
outlook = win32.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
inbox = namespace.GetDefaultFolder(6)

# Connect to Excel
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True

# Open the selected workbook
workbook = excel.Workbooks.Open(file_path)
worksheet1 = workbook.Worksheets('Sheet1')
worksheet2 = workbook.Worksheets('Sheet2')

# Define keywords and column mappings
keywords1 = {
    'Location': 'B',
    'Category': 'C',
    'Business Unit': 'D',
    'Cost Centre': 'E',
    'PGS Delivery Manager': 'F',
    'Role': 'G',
    'Skill': 'H',
    'Skill (for report)': 'I',
    'Status': 'J',
    'Resource Name': 'K',
    'Emp ID': 'L',
    'PGS ID': 'M',
    'Billing Rate 2023': 'N',
    'Exp. Bracket': 'O',
    'Request Date': 'P',
    'Selection Date': 'Q',
    'DOJ': 'R',
    'PGS Start Date': 'S',
    'Offboarding Date': 'T',
    'Offboard Reason': 'U',
    'Deferred Reason': 'V',
    'Mail ID': 'W',
    'Remarks / Updates': 'X'
}

keywords2 = {
    'Name': 'B',
    'Source': 'C',
    'Skill': 'D',
    'Profile shared': 'E',
    'Profile Status': 'F',
    'Expected Start': 'G',
    'Interview Date': 'H',
    'Panel': 'I',
    'Interview Status': 'J',
    'Remarks': 'K',
    'BU': 'L'
}

# Define a function to find the last filled row in the Excel sheet
def find_last_filled_row(worksheet):
    row = 2
    while worksheet.Range('A' + str(row)).Value is not None:
        row += 1
    return row - 1

# Define a function to update a worksheet with keyword values
def update_worksheet(keyword_values):
    worksheet1_row = find_last_filled_row(worksheet1) + 1
    worksheet2_row = find_last_filled_row(worksheet2) + 1
    
    for keyword, column in keywords1.items():
        worksheet1.Range(column + str(worksheet1_row)).Value = keyword_values[keyword]
    
    for keyword, column in keywords2.items():
        worksheet2.Range(column + str(worksheet2_row)).Value = keyword_values[keyword]

# Define a function to check if an email has all values already present in the worksheet
def email_has_all_values(email_body):
    for line in email_body.split('\n'):
        if line.startswith(tuple(keywords1.keys())) or line.startswith(tuple(keywords2.keys())):
            keyword, value = line.split(':', maxsplit=1)
            if keyword.strip() in keywords1 and worksheet1.Range(keywords1[keyword.strip()] + ':X' + str(find_last_filled_row(worksheet1))).Value is not None:
                return True
            elif keyword.strip() in keywords2 and worksheet2.Range(keywords2[keyword.strip()] + ':L' + str(find_last_filled_row(worksheet2))).Value is not None:
                return True
    return False

# Define a function to extract keyword values from an email body
def extract_keyword_value(email_body):
    keyword_values = {}
    
    for line in email_body.split('\n'):
        if line.startswith(tuple(keywords1.keys())) or line.startswith(tuple(keywords2.keys())):
            keyword, value = line.split(':', maxsplit=1)
            keyword_values[keyword.strip()] = value.strip()
    
    return keyword_values

# Define a function to process new emails and update the worksheets if necessary
def process_emails():
    # Iterate over inbox messages
    for email in inbox.Items:
        # Check if email has already been processed
        if email.FlagStatus == 1:
            continue
        
        # Extract email body and check if it has all values already present in the worksheet
        email_body = email.Body
        if not email_has_all_values(email_body):
            # Extract keyword values from email body and update worksheets
            keyword_values = extract_keyword_value(email_body)
            update_worksheet(keyword_values)
            
            # Mark email as processed
            email.FlagStatus = 1

# Schedule the process_emails function to run every second
schedule.every(1).seconds.do(process_emails)

# Run the scheduled tasks in a loop
while True:
    schedule.run_pending()
    time.sleep(1)
