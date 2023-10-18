import smtplib
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import concurrent.futures

# Email configuration
smtp_server = 'your_smtp_server.com'
smtp_port = 587
smtp_username = 'your_email@gmail.com'
smtp_password = 'your_app_password'
sender_email = 'your_email@gmail.com'
subject = 'Custom Subject'
message_template = 'Hello {name},\n\nThis is a custom message for you.'

# Load Excel file and specify the sheet
excel_file = 'your_excel_file.xlsx'
wb = openpyxl.load_workbook(excel_file)
sheet = wb['Sheet2']  # Change 'Sheet2' to the name of your sheet

#Assuming the name and email are at row[0] and row[1] respectively
# Function to send an email
def send_email(row):
    name, email = row[0], row[1]
    try:
        # Create a new SMTP connection for each thread/task
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = email
        msg['Subject'] = subject
        custom_message = message_template.format(name=name)
        msg.attach(MIMEText(custom_message, 'plain'))

        # Send the email
        server.sendmail(sender_email, email, msg.as_string())
        print(f"Email sent to {name} at {email}")

        # Quit the SMTP server for this thread/task
        server.quit()
    except Exception as e:
        print(f"Failed to send email to {name} at {email}: {str(e)}")

# Create a ThreadPoolExecutor for parallel processing
with concurrent.futures.ThreadPoolExecutor() as executor:
    futures = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        futures.append(executor.submit(send_email, row))

    # Wait for all email sending tasks to complete
    concurrent.futures.wait(futures)

# Close the Excel file
wb.close()
