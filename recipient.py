import email
import imaplib
import os

from PyQt5 import uic
from PyQt5.QtCore import QSize
from PyQt5.QtWidgets import *


class RecipientUi(QMainWindow):

    def __init__(self):
        super(RecipientUi, self).__init__()
        uic.loadUi('recipient.ui', self)
        self.show()

        self.gmail_password = os.environ.get('gmail_password')

        self.imap = None

        self.btn_login.clicked.connect(self.login)
        self.btn_get_messages.clicked.connect(self.get_messages)
        self.btn_clear.clicked.connect(self.clear)

    def login(self):
        try:
            self.imap = imaplib.IMAP4_SSL(host=self.options_imap_server.currentText())
            self.imap.login(self.input_email.text(), self.gmail_password)

            create_message_box('Successful authentication')

            self.text_body.setEnabled(True)
            self.input_number_messages.setEnabled(True)
            self.btn_get_messages.setEnabled(True)
            self.btn_clear.setEnabled(True)


        except:
            create_message_box('Error on login')

    def get_messages(self):
        status, messages = self.imap.select('Inbox')
        total_of_messages = int(messages[0])

        for i in range(total_of_messages, total_of_messages - int(self.input_number_messages.text()), -1):
            status, data = self.imap.fetch(str(i), "(RFC822)")
            message = email.message_from_bytes(data[0][1])

            text = ''
            body = None

            text += f"Message number: {i}\n"
            text += f"From: {message.get('From')}\n"
            text += f"To: {message.get('To')}\n"
            text += f"Date: {message.get('Date')}\n"
            text += f"Subject: {message.get('Subject')}\n"

            text += 'Content:'
            for part in message.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                try:
                    body = part.get_payload(decode=True).decode()
                except:
                    pass

                if content_type == "text/plain" and "attachment" not in content_disposition:
                    text += f"Content-Type: {content_type}\n"
                    text += f"Content-Disposition: {content_disposition}\n"
                    if body:
                        text += 'Content:'
                        text += body
                        text += '\n'
                elif "attachment" in content_disposition:
                    download_attachment(part, i)

            text += '=' * 40
            text += '\n\n'

            self.text_body.append(text)

    def clear(self):
        self.imap.close()
        self.input_email.clear()
        self.input_password.clear()
        self.text_body.clear()
        self.input_number_messages.clear()
        self.text_body.setEnabled(False)
        self.input_number_messages.setEnabled(False)
        self.btn_get_messages.setEnabled(False)
        self.btn_clear.setEnabled(False)


def download_attachment(part, folder_name):
    folder_name = f'folder_{folder_name}'
    filename = part.get_filename()

    if filename:
        if not os.path.isdir(folder_name):
            os.mkdir(folder_name)
        filepath = os.path.join(folder_name, filename)
        open(filepath, "wb").write(part.get_payload(decode=True))


def create_message_box(message):
    message_box = QMessageBox()
    message_box.setText(message)
    message_box.setBaseSize(QSize(300, 120))
    message_box.exec()


app = QApplication([])
window = RecipientUi()
app.exec_()
