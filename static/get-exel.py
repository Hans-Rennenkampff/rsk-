import imaplib
import email
import os
import glob

# Set email login credentials
username = 'rsk-rasp@yandex.com'
password = 'zsmjvcufsboisjvk'

# Connect to the Yandex Mail server
mail = imaplib.IMAP4_SSL('imap.yandex.com')
mail.login(username, password)
mail.select('inbox')

# Search for the last email in the inbox
result, data = mail.uid('search', None, 'ALL')
latest_email_uid = data[0].split()[-1]
result, email_data = mail.uid('fetch', latest_email_uid, '(RFC822)')

# Parse the email content
raw_email = email_data[0][1].decode('utf-8')
email_message = email.message_from_string(raw_email)

# Check if the email contains any attached .xlsx files
for part in email_message.walk():
    # далее идет очистка дириктории output от графиков

    files = glob.glob('.\web\output')
    for f in files:
        os.remove(f)

    if part.get_content_type() == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
        # Download the attached .xlsx file
        filename = 'raspisanie.xlsx'
        if os.path.exists('raspisanie.xlsx'):
            os.remove('raspisanie.xlsx')
            print('File deleted successfully.')
        else: print('File does not exist.')
        os.chdir(f'/resourses/xl')
        with open(filename, 'wb') as f:
            f.write(part.get_payload(decode=True))
        print('File downloaded successfully.')
        break


# Delete all emails in the inbox
mail.store('1:*', '+FLAGS', '\\Deleted')
mail.expunge()
print('Inbox cleared.')

# Disconnect from the Yandex Mail server
mail.close()
mail.logout()