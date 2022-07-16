# modules
import imaplib
import email
from email.header import decode_header
# from typing_extensions import Self
import webbrowser
import os
import pandas as pd
import string


class EmailConnector():

    def __init__(
        self,
        useremail: str,
        password: str,
        imapserver: str,
        conn: str = '',
    ) -> None:
        "Connection parameters"
        self.useremail = useremail
        self.password = password
        self.imapserver = imapserver
        self.conn = None

    def create_conn(self):
        # function for creatingconnection
        # create an IMAP4 class with SSL
        imap = imaplib.IMAP4_SSL(self.imapserver)
        try:
            imap.login(self.useremail, self.password)  # authenticate
        except:
            raise
        self.conn = imap


# imap_server = 'imap.gmail.com'  # for gmail
# imap_server = "outlook.office365.com" # for outlook
