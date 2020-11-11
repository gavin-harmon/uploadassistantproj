"""
Opens an email on a button request
==================================

This is a script that executes a command tied to the "Ask for Assistance" button push in the Upload Assistant. The action opens a "Request for Assistance" email.

==================================

"""

# Examples of how to use this are avaliable here: https://stackoverflow.com/
# questions/20956424/how-do-i-generate-and-open-an-outlook-email-with-python-but-do-not-send

#libraries and modules used

import win32com.client as win32

def emailer(text, subject, recipient):
    """Opens a prepopulated Outlook email if Outlook is open.

    Examples of how to use this are available here:

    https://stackoverflow.com/questions/20956424/how-do-i-generate-and-open-an-outlook-email-with-python-but-do-not-send

    Args:
        text (str): Body of email.

        subject (str): Subject of email.

        recipient (str): List of recipients example "<person.1@company.com>; <person.2@company.com>"
   """
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    mail.Display(True)

if __name__ == '__main__':
    emailer("", "Upload Application Assistance Required", " <Dana.Mark@allianz.com>;"
                                                          " <angela.chenxx@allianz.com>; <Federico.Guerreschi@allianz.com>; <gavin.harmon@allianz.com>")
    """execution gaurd"""
