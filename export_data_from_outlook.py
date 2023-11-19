from pathlib import Path
import datetime
from pytz import timezone
import win32com.client
from output_file import ContactsFile
import os

contacts_file = ContactsFile()


class OutLook:
    def __init__(self):
        self.meetings = None
        self.inbox = None
        self.outlook = None
        self.list_of_contacts = []
        today = datetime.datetime.now()
        self.end_date = datetime.datetime(year=today.year, month=today.month, day=today.day, tzinfo=timezone('UTC'))
        self.start_date = datetime.datetime(year=2023, month=3, day=10, tzinfo=timezone('UTC'))


    def check_output_folder(self):
        if not os.path.exists("Output"):
            folder_name = Path('Output')
            folder_name.mkdir(parents=True)

    def connect_to_outlook(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            # self.inbox = self.outlook.Folders("talor.keren1@gmail.com").Folders("Inbox")
            self.meetings = self.outlook.GetDefaultFolder(9)
            # print(self.inbox.Items)
            # for x in self.inbox.Items:
            #     # print(x)
            #     subject = x.Subject
            #     for y in x.recipients:
            #         print(y.Adress)

                # print(subject)

        except:
            print("try again")

    def get_messages(self):
        meets = self.meetings.Items
        meets_counter = meets.Count
        print(meets_counter)
        for meet in meets:
            # meet_subject = meet.Subject
            # print(meet_subject)
            # body = message.body

            # meet_date = meet.ReceivedTime
            # if self.start_date < meet_date < self.end_date:
            for recip in meet.recipients:
                if "/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT" not in recip.Address and "cropx.com" not in recip.Address:
                    if recip.Address.lower() not in self.list_of_contacts:
                        self.list_of_contacts.append(recip.Address.lower())

    def main(self):
        self.check_output_folder()
        self.connect_to_outlook()
        self.get_messages()
        contacts_file.append_contacts(self.list_of_contacts)


run_script = OutLook()
run_script.main()
