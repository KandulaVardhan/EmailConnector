# modules
from base import EmailConnector
import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import pandas as pd
import string


class SourceEmailConnector(EmailConnector):
    def __init__(self, **kwargs) -> None:
        super().__init__("emailid",
                         "app password", "imap.gmail.com")
        self.create_conn()   # creating connection
        self.pull_data()
        self.__exit__()

    def cleanText(self, text) -> string:
        # cleaning text for creating a folder
        return "".join(c if c.isalnum() else "_" for c in text)

    def test_connector(self) -> bool:
        """
        Tests if the input configuration can be used to successfully connect to the integration.
        Tests if the enough read priviledges are available for credentails provided
        :return: Boolean indicating a Success or Failure
        """

    def pull_data(self) -> pd.DataFrame:
        imap = self.conn
        # fetching emails from INBOX folder
        status, messages = imap.select("INBOX")
        # We can use the imap.list() method to see the available mailboxes.
        # print(imap.list())
        # number of top emails to fetch
        N = 4
        # total number of emails
        messages = int(messages[0])  # converting to integer
        # print(messages)  # shows the total number of messages in that specific folder
        # print(status)  # shows the status of our connection whether we have received our message successfully or not.
        count = 0
        dic = {count: ["Subject", "From", "Date", "To", "Email body"]}
        count = count+1
        dic[count] = []
        count = count+1
        # oldest email will have id as 1 and recent one will have higher value, so looping from higher-lower
        for i in range(messages, messages-N, -1):
            # fetch the email message by ID
            res, msg = imap.fetch(str(i), "(RFC822)")
            for response in msg:
                lst = []
                if isinstance(response, tuple):
                    # parse a bytes email into a message object
                    msg = email.message_from_bytes(response[1])
                    # decode the email subject to human-readable Unicode
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        # if it's a bytes, decode to str
                        subject = subject.decode(encoding)
                    # decode email sender to human-readable Unicode
                    From, encoding = decode_header(msg.get("From"))[0]
                    if isinstance(From, bytes):
                        # if it's a bytes, decode to str
                        From = From.decode(encoding)
                    Date, encoding = decode_header(msg.get("Date"))[0]
                    if isinstance(Date, bytes):
                        # if it's a bytes, decode to str
                        Date = Date.decode(encoding)
                    To, encoding = decode_header(msg.get("To"))[0]
                    if isinstance(To, bytes):
                        # if it's a bytes, decode to str
                        To = To.decode(encoding)
                    lst.append(subject)
                    lst.append(From)
                    lst.append(Date)
                    lst.append(To)
                    # if the email message is multipart (text+HTML)
                    if msg.is_multipart():
                        # iterate over email parts
                        for part in msg.walk():
                            # extract content type of email
                            content_type = part.get_content_type()
                            # detecting file attachments with Content-Disposition header
                            content_disposition = str(
                                part.get("Content-Disposition"))
                            # get the email body
                            try:
                                body = part.get_payload(
                                    decode=True).decode()
                            except:
                                pass
                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                # print text/plain emails and skip attachments
                                lst.append(body)
                            elif content_type == "text/html" and "attachment" not in content_disposition:
                                try:
                                    # print html part to retrive table data
                                    df_list = pd.read_html(body)
                                    # print(df_list)
                                    # no. of tables in our HTML part
                                    N = len(df_list)
                                    # A,B,C,D...
                                    groups_names = list(
                                        string.ascii_uppercase[0:N])
                                    # grouping all the extracted tables to one table
                                    df = pd.DataFrame()
                                    for i in range(0, N):
                                        group_col = [
                                            groups_names[i]] * len(df_list[i])
                                        df_list[i]['Group'] = group_col
                                        df = df.append(df_list[i])
                                        #df = pd.concat(df, df_list[i])
                                    df.to_csv('Tables.csv')
                                except:
                                    pass
                            elif "attachment" in content_disposition:
                                # download attachment
                                filename = part.get_filename()
                                if filename:
                                    folder_name = self.cleanText(subject)
                                    if not os.path.isdir(folder_name):
                                        # make a folder for this email (named after the subject)
                                        os.mkdir(folder_name)
                                    filepath = os.path.join(
                                        folder_name, filename)
                                    # download attachment and save it
                                    open(filepath, "wb").write(
                                        part.get_payload(decode=True))
                    else:
                        # extract content type of email
                        content_type = msg.get_content_type()
                        # get the email body
                        try:
                            body = msg.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain":
                            # print only text email parts
                            lst.append(body)
                # print(lst)
                if count % 2 == 0:
                    dic[count//2] = lst
                count = count+1
        # print(dic)
        ans = pd.DataFrame.from_dict(dic)
        return ans

    def __exit__(self):
        # close the connection and logout
        imap = self.conn
        imap.close()
        imap.logout()

    def pull_meta_data(self, **kwargs) -> pd.DataFrame:
        pass


SourceEmailConnector()
