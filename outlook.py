from appscript import app, k
from mactypes import Alias
from pathlib import Path
from datetime import datetime

def create_message_with_attachment():
    subject = 'This is an important email!'
    body = 'Just kidding its not.'
    to_recip = ['myboss@mycompany.com', 'theguyih8@mycompany.com']

    msg = Message(subject=subject, body=body, to_recip=to_recip)

    # attach file
    p = Path('path/to/myfile.pdf')
    msg.add_attachment(p)

    msg.show()

def create_dsr(outlook):
    daily_status_report_file = open('daily_status_report.html', 'r')
    daily_status_report_html = daily_status_report_file.read()
    current_date = datetime.today().strftime('%d %B')
    title = 'Sarthak Negi - Status Report - ' + current_date
    message = Message(parent=outlook, subject=title, body=daily_status_report_html, to_recip=[''])
    message.show()

def create_csr(outlook):
    client_status_report_file = open('client_status_report.html', 'r')
    client_status_report_html = client_status_report_file.read()
    current_date = datetime.today().strftime('%d %B')
    title = 'Sarthak Negi - Status Report - ' + current_date
    message = Message(parent=outlook, subject=title, body=client_status_report_html, to_recip=[''])
    message.show()

class Outlook(object):
    def __init__(self):
        self.client = app('Microsoft Outlook')
        print('Opened outlook')


class Message(object):
    def __init__(self, parent=None, subject='', body='', to_recip=[], cc_recip=[], show_=True):
        if parent is None:
            parent = Outlook()
        client = parent.client

        print('Message init')
        print(client)
        self.msg = client.make(
            new=k.outgoing_message,
            with_properties={k.subject: subject, k.content: body})

        # self.add_recipients(emails=to_recip, type_='to')
        # self.add_recipients(emails=cc_recip, type_='cc')

        print('message set up')
        if show_:
            print('come to foreground')
            self.show()
            print('foreground success')

    def show(self):
        self.msg.open()
        self.msg.activate()

    def add_attachment(self, p):
        # p is a Path() obj, could also pass string

        p = Alias(str(p))  # convert string/path obj to POSIX/mactypes path

        attach = self.msg.make(new=k.attachment, with_properties={k.file: p})

    def add_recipients(self, emails, type_='to'):
        if not isinstance(emails, list):
            emails = [emails]
        for email in emails:
            self.add_recipient(email=email, type_=type_)

    def add_recipient(self, email, type_='to'):
        msg = self.msg

        if type_ == 'to':
            recipient = k.to_recipient
        elif type_ == 'cc':
            recipient = k.cc_recipient

        msg.make(new=recipient, with_properties={
                 k.email_address: {k.address: email}})

