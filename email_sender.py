import smtplib
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


# Email configuration
smtp_server = 'your_smtp_server.com'
smtp_port = 587  # Update the port according to your SMTP server settings
smtp_username = 'your_email@gmail.com'
smtp_password = 'your_app_password'  # Use the App Password you generated
sender_email = 'your_email@gmail.com'
subject = 'Custom Subject'
message_template = 'Hello {name},\n\nThis is a custom message for you.'

# Load Excel file and specify the sheet
excel_file = 'your_excel_file.xlsx'
wb = openpyxl.load_workbook(excel_file)
sheet = wb['Sheet2']  # Change 'Sheet2' to the name of your sheet

#Here name and email are at row 0 and 2 of my excel sheet
for row in sheet.iter_rows(min_row=2, values_only=True):
    name = row[0]
    email = row[2]


     # Create the email content
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = email
    msg['Subject'] = subject
    # Customize the message with the person's name
    custom_message = message_template.format(name=name)
    msg.attach(MIMEText(custom_message, 'plain'))

    try:
    # Connect to the SMTP server and send the email
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(sender_email, email, msg.as_string())
        server.quit()
        print(f"Email sent to {name} at {email}")
    except Exception as e:
        print(f"Failed to send email to {name} at {email}: {str(e)}")


# Close the Excel file
wb.close()
