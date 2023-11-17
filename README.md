# collect_new_outlook_contacts

The script will collect new contacts from your Outlook sent emails and will save them in folder Output.

The collection will start from 10-09-2023, in order to change it go to line 18 in file export_data_from_outlook.py
        self.start_date = datetime.datetime(year=2023, month=9, day=10, tzinfo=timezone('UTC'))
