"""
This script automates the process of filtering rows with "Processing" status from an Excel file 
named "Orders report.xlsx". It then creates a draft email in Outlook with the filtered data attached.
Feel free to modify the script to suit your needs.
"""

# Imports (you'll need to install openpyxl and pywin32 if you don't have them)
import pandas as pd
import win32com.client as win32
import os
from datetime import date

def create_filtered_report_and_draft_email():
    # Check if the input Excel file exists
    excel_file = "Orders report.xlsx"
    if not os.path.exists(excel_file):
        print(f"Error: The file '{excel_file}' does not exist.")
        return

    # Get today's date
    today = date.today().strftime('%Y-%m-%d')

    # Extract rows with "Processing" as Order Status from Excel
    df = pd.read_excel(excel_file, engine='openpyxl')

    # Filter the DataFrame
    filtered_df = df[df['Order Status'] == 'Processing']

    # Save the filtered rows to a new Excel file, including today's date in the filename
    filtered_file_path = f"Filtered Orders {today}.xlsx"
    filtered_df.to_excel(filtered_file_path, index=False, engine='openpyxl')

    # Create a draft email in Outlook Desktop and attach the new Excel file
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'example@email.com'  # Replace with the email address(es) you want to send the email to
    mail.CC = 'example_cc@email.com'  # Replace with the email address(es) you want to CC
    mail.Subject = f'Orders with Processing Status - {today}' # Replace with the subject you want
    mail.Body = f"Attached are the orders with a status of 'Processing' as of {today}." # Replace with the body you want
    mail.Attachments.Add(os.path.abspath(filtered_file_path))

    # Display the draft email
    mail.Display()

    print(f"Draft email created with '{filtered_file_path}' attached!")

if __name__ == '__main__':
    create_filtered_report_and_draft_email()