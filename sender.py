import os
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from PyQt5 import uic
from PyQt5.QtCore import QSize
from PyQt5.QtWidgets import *


class SenderUi(QMainWindow):

    def __init__(self):
        super(SenderUi, self).__init__()
        uic.loadUi('sender.ui', self)
        self.show()

        self.MAX_SIZE_BYTES = 2_097_152
        self.gmail_password = os.environ.get('gmail_password')
        self.outlook_password = os.environ.get('outlook_password')
        self.server = None
        self.size = 0
        self.message = MIMEMultipart()

        self.btn_login.clicked.connect(self.login)
        self.btn_add_attachements.clicked.connect(self.add_attachements)
        self.btn_send.clicked.connect(self.send)

    def login(self):
        try:
            context = ssl.create_default_context()
            if self.options_connection.currentText() == 'SSL':
                self.server = smtplib.SMTP_SSL(self.options_smtp_server.currentText(), self.options_port.currentText(),
                                               context=context)
                self.server.login(self.input_from.text(), self.gmail_password)

            else:
                self.server = smtplib.SMTP(self.options_smtp_server.currentText(), self.options_port.currentText())
                self.server.starttls(context=context)
                self.server.login(self.input_from.text(), self.input_password.text())

            create_message_box('Successful authentication')

            self.input_to.setEnabled(True)
            self.input_subject.setEnabled(True)
            self.text_body.setEnabled(True)
            self.btn_add_attachements.setEnabled(True)

        except smtplib.SMTPAuthenticationError:
            create_message_box('Incorrect Login Info!')
        except:
            create_message_box('General error')

    def add_attachements(self):
        options = QFileDialog.Options()
        filenames, status = QFileDialog.getOpenFileNames(self, 'Open File', filter='All Files (*.*)', options=options)

        if filenames != []:
            for filename in filenames:
                attachement = open(filename, 'rb')
                if self.size + os.path.getsize(filename) > self.MAX_SIZE_BYTES:
                    break
                self.size += os.path.getsize(filename)
                filename = filename[filename.rfind('/') + 1:]
                payload = MIMEBase('application', 'octet-stream')
                payload.set_payload(attachement.read())
                payload.add_header('Content-Disposition', f'attachement; filename={filename}')
                encoders.encode_base64(payload)
                self.message.attach(payload)

                if not self.label_attachements.text().endswith(':'):
                    self.label_attachements.setText(self.label_attachements.text() + ',')
                self.label_attachements.setText(self.label_attachements.text() + ' ' + filename)

    def send(self):
        try:
            plain_text = MIMEText(self.text_body.toPlainText(), 'plain')
            self.message['Subject'] = self.input_subject.text()
            self.message['From'] = self.input_from.text()
            self.message['To'] = self.input_to.text()
            self.message.attach(plain_text)
            self.server.sendmail(self.input_from.text(), [self.input_to.text()], self.message.as_string())

            create_message_box('Email successfully sent')
            self.reset()

        except:
            create_message_box('Error on sending email')

    def reset(self):
        self.message = MIMEMultipart()
        self.input_from.clear()
        self.input_password.clear()
        self.input_to.clear()
        self.input_subject.clear()
        self.text_body.clear()
        self.input_to.setEnabled(False)
        self.input_subject.setEnabled(False)
        self.text_body.setEnabled(False)
        self.btn_add_attachements.setEnabled(False)
        self.label_attachements.clear()
        self.label_attachements.setText('Attachements:')
        self.server.close()


def create_message_box(message):
    message_box = QMessageBox()
    message_box.setText(message)
    message_box.setBaseSize(QSize(300, 120))
    message_box.exec()


app = QApplication([])
window = SenderUi()
app.exec_()
