import win32com.client

class EmailSender:
    def __init__(self) -> None:
        self.ol = win32com.client.Dispatch('outlook.application')
    def send_email(self, subject, body, recipient, cc_emails = [], display_or_send='d') -> None:
        """
        Sends an email from Outlook account

        subject: subject of email
        body: body of email
        recipient: email address of recipient
        display_or_send: 'd' if display (and send manually), 's' if send without displaying
        """
        olmailitem = 0x0 #size of new email
        newmail = self.ol.CreateItem(olmailitem)
        newmail.Subject = subject
        newmail.To = recipient
        newmail.Body = body
        newmail.CC = ';'.join(cc_emails)
        if display_or_send == 's':
            newmail.Send()
        else:
            newmail.Display()
