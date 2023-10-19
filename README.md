# AutomatedEmailSender

A simple Python alternative script to automate sending emails in batches from personal accounts through SMTP. \
[parrallel.py](https://github.com/abhinavpannala/AutomatedEmailSender/blob/main/parallel.py) is the same code with multi-threading support, which helps with faster runtime.

### Caveats:
1. Email providers usually stop connections to their services after a few iterations of the parallel code to avoid suspicious activity, as threads open a distinct connection to the servers concurrently.
2. Email providers close the server connection after a limited amount of emails are sent through a single connection to avoid spam.

### Usecase:
This script will be helpful if the requirement is to email less than 100 contacts from an Excel sheet.
