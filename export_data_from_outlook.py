from pathlib import Path
import datetime
from pytz import timezone
import win32com.client
from output_file import ContactsFile
import os

contacts_file = ContactsFile()


class OutLook:
    def __init__(self):
        self.inbox = None
        self.outlook = None
        self.list_of_contacts = []
        today = datetime.datetime.now()
        self.end_date = datetime.datetime(year=today.year, month=today.month, day=today.day, tzinfo=timezone('UTC'))
        self.start_date = datetime.datetime(year=2023, month=9, day=10, tzinfo=timezone('UTC'))


    def check_output_folder(self):
        if not os.path.exists("Output"):
            folder_name = Path('Output')
            folder_name.mkdir(parents=True)

    def connect_to_outlook(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            # self.inbox = self.outlook.Folders("talor.keren1@gmail.com").Folders("Inbox")
            self.inbox = self.outlook.GetDefaultFolder(5)
        except:
            print("try again")

    def get_messages(self):
        messages = self.inbox.Items
        messages_counter = messages.Count
        print(messages_counter)
        for message in messages:
            # subject = message.Subject
            # body = message.body

            sent_date = message.ReceivedTime
            if self.start_date < sent_date < self.end_date:
                for recip in message.recipients:
                    if "/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT" not in recip.Address:
                        if recip.Address.lower() not in self.list_of_contacts:
                            self.list_of_contacts.append(recip.Address.lower())

    def main(self):
        self.check_output_folder()
        self.connect_to_outlook()
        self.get_messages()
        contacts_file.append_contacts(self.list_of_contacts)


run_script = OutLook()
run_script.main()
