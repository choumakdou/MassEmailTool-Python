import openpyxl
import os
import sys

# Note: This script requires 'pywin32' which only works on Windows with Outlook installed.
# Install it using: pip install pywin32 openpyxl

try:
    import win32com.client as win32
except ImportError:
    print("Error: 'pywin32' not found. This script must be run on Windows.")
    print("Install it with: pip install pywin32")
    sys.exit(1)

def send_mass_emails(excel_path):
    if not os.path.exists(excel_path):
        print(f"Error: File {excel_path} not found.")
        return

    # Load Excel workbook
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active

    # Initialize Outlook
    try:
        outlook = win32.Dispatch('outlook.application')
    except Exception as e:
        print(f"Error: Could not connect to Outlook. {e}")
        return

    # Assuming headers are in row 1
    # Columns: A=Name, B=Email, C=Subject, D=Update, E=Path1, F=Path2, G=Status
    for row in range(2, sheet.max_row + 1):
        name = sheet.cell(row=row, column=1).value
        email = sheet.cell(row=row, column=2).value
        subject = sheet.cell(row=row, column=3).value
        update = sheet.cell(row=row, column=4).value
        path1 = sheet.cell(row=row, column=5).value
        path2 = sheet.cell(row=row, column=6).value
        status = sheet.cell(row=row, column=7).value

        if not email or status == "Sent":
            continue

        try:
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = subject
            
            # Construct personalized body
            body = f"Dear {name},\n\n{update}\n\nBest regards,\nMarketing Team"
            mail.Body = body

            # Add attachments if paths exist
            for path in [path1, path2]:
                if path and os.path.exists(str(path)):
                    mail.Attachments.Add(os.path.abspath(str(path)))
                elif path:
                    print(f"Warning: Attachment not found at {path}")

            # Send or Display
            # mail.Send() # Uncomment to send automatically
            mail.Display() # Use Display first to test
            
            # Update status
            sheet.cell(row=row, column=7).value = "Sent"
            print(f"Email prepared for {email}")

        except Exception as e:
            print(f"Failed to send to {email}: {e}")
            sheet.cell(row=row, column=7).value = "Error"

    wb.save(excel_path)
    print("Process complete. Excel updated.")

if __name__ == "__main__":
    # Change this to your actual Excel file path
    FILE_PATH = "marketing_list.xlsx"
    send_mass_emails(FILE_PATH)
