#!/usr/bin/python3
import time
import imaplib
import email
import os

# Set email login credentials

username = 'rsk-rasp@yandex.com'
password = 'zsmjvcufsboisjvk'
SAVE_PATH = '/root/rsk-main/resourses/xl'
FILE_NAME = 'raspisanie.xlsx'

def clear_output_folder():
    # Определяем путь к папке "output" относительно местоположения текущего скрипта
    output_folder_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', 'web', 'output'))

    # Проверяем, существует ли папка "output"
    if os.path.exists(output_folder_path):
        # Получаем список файлов в папке
        files = os.listdir(output_folder_path)
        
        # Проходим по каждому файлу и удаляем его
        for file_name in files:
            file_path = os.path.join(output_folder_path, file_name)
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
                elif os.path.isdir(file_path):
                    # Если нужно удалить и подпапки, добавьте эту строку:
                    # os.rmdir(file_path)
                    pass
            except Exception as e:
                print(f"Ошибка при удалении файла {file_path}: {e}")
        print(f"Все файлы в папке {output_folder_path} успешно удалены.")
    else:
        print(f"Папка {output_folder_path} не существует.")       
 
def download_attachment():
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
        if part.get_content_type() == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            if os.path.exists('raspisanie.xlsx'):
                os.remove('raspisanie.xlsx')
                print('File deleted successfully.')
                time.sleep(10)
            file_path = os.path.join(SAVE_PATH, FILE_NAME)
            with open(file_path, 'wb') as f:
                f.write(part.get_payload(decode=True))
            print('Файл успешно сохранен:', file_path)

    # Delete all emails in the inbox
    mail.store('1:*', '+FLAGS', '\\Deleted')
    mail.expunge()
    print('Inbox cleared.')
    # Disconnect from the Yandex Mail server
    mail.close()
    mail.logout()

# Schedule the script to run every Sunday at 18:00

if __name__ == '__main__':
   download_attachment()
