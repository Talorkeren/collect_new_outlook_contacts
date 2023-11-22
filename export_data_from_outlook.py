from pathlib import Path
import datetime
from pytz import timezone
import win32com.client
from output_file import ContactsFile
import os

contacts_file = ContactsFile()


class OutLook:
    def __init__(self):
        self.meets = None
        self.list_of_contacts = []
        today = datetime.datetime.now()
        self.today = datetime.datetime(year=today.year, month=today.month, day=today.day, tzinfo=timezone('UTC'))
        self.start_date = self.today - datetime.timedelta(30)

    def check_output_folder(self):
        if not os.path.exists("Output"):
            folder_name = Path('Output')
            folder_name.mkdir(parents=True)

    def connect_to_outlook(self):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            # self.calendar = self.outlook.GetDefaultFolder(9)
            shared_calendar = outlook.CreateRecipient("Shahar Dadon Zusman")
            self.meets = outlook.GetSharedDefaultFolder(shared_calendar, 9).Items

        except:
            print("check what happened")

    def get_messages(self):
        for meet in self.meets:
            meet_date = meet.Start
            if self.start_date <= meet_date <= self.today:
                for recip in meet.recipients:
                    if "/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT" not in recip.Address and "cropx.com" not in recip.Address:
                        if recip.Address.lower() not in self.list_of_contacts:
                            print(f'{str(meet_date)[:10]} - {recip.Address.lower()}')
                            self.list_of_contacts.append(recip.Address.lower())

    def main(self):
        self.check_output_folder()
        self.connect_to_outlook()
        self.get_messages()
        contacts_file.append_contacts(self.list_of_contacts)


run_script = OutLook()
run_script.main()
