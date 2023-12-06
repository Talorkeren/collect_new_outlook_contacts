import datetime
from pytz import timezone
import win32com.client


class OutLook:
    def __init__(self):
        self.start_date = None
        self.list_of_contacts = None
        self.list_of_names = []
        self.list_of_address = []
        self.meets = None
        today = datetime.datetime.now()
        self.today = datetime.datetime(year=today.year, month=today.month, day=today.day, tzinfo=timezone('UTC'))

    def connect_to_outlook(self):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            ## Use for DefaultFolder
            # self.meets = outlook.GetDefaultFolder(9).Items

            ## Use for Shared Default Folder
            shared_calendar = outlook.CreateRecipient("Tomer Tzach")
            self.meets = outlook.GetSharedDefaultFolder(shared_calendar, 9).Items
            return self.meets

        except:
            print("check what happened")

    def get_messages(self, meets, days_back):
        self.meets = meets
        self.start_date = self.today - datetime.timedelta(int(days_back))
        for meet in self.meets:
            meet_date = meet.Start
            if self.start_date <= meet_date <= self.today:
                for recip in meet.recipients:
                    if "/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT" not in recip.Address and "cropx.com" not in recip.Address:
                        if recip.Address.lower() not in self.list_of_address:
                            self.list_of_address.append(recip.Address.lower())
                            self.list_of_names.append(recip.name)

        self.list_of_contacts = dict(zip(self.list_of_address, self.list_of_names))
        return self.list_of_contacts





