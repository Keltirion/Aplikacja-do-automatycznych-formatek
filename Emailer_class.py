# Libraries to import
import win32com.client as win32

# Emailer class with creation method.
class Emailer:
    def __init__(self, recipient, subject, body,):
        self.recipient = recipient
        self.subject = subject
        self.body = body
    # Creates an email within outlook. Three parametrs mus be given.
    # Subject, recipient and body.
    def create(self):
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.To = self.recipient
        email.Subject = self.subject
        email.htmlBody = self.body
        email.Display(True)
