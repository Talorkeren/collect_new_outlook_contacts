import csv
import time
import pandas
import shutil


class ContactsFile:
    def __init__(self):
        self.new_filename_csv = None
        self.count_addr = []
        self.var = None
        self.list_addr_file = []
        self.addr = None
        self.write_new_row = None
        self.data_contacts_from_file = None
        self.contacts_from_outlook = None

    def read_contacts_file(self, filepath):
        self.list_addr_file = []
        self.data_contacts_from_file = pandas.read_csv(filepath)
        self.data_contacts_from_file = self.data_contacts_from_file.to_dict(orient="records")
        for key in self.data_contacts_from_file:
            self.list_addr_file.append(key["Email_Address"])
        return len(self.list_addr_file)

    def copy_file_and_rename(self, filepath):
        time_str = time.strftime("%Y%m%d-%H%M%S")
        self.new_filename_csv = f"Newsletter - Mailing list updated {time_str}.csv"
        shutil.copy(filepath, self.new_filename_csv)

    def append_contacts(self, list_of_contacts):
        self.contacts_from_outlook = list_of_contacts
        print(self.new_filename_csv)
        self.count_addr = []
        for key, value in self.contacts_from_outlook.items():
            self.addr = key.lower()
            self.var = {"Email_Address": key.lower(),
                        # For Shahar
                        "First_Name": value}
                        ## For Matan
                        # "Full_Name": value}

            if self.addr not in self.list_addr_file:
                self.count_addr.append(self.addr)
                with open(self.new_filename_csv, 'a', newline='') as f:
                    self.write_new_row = csv.DictWriter(f, fieldnames=self.var.keys())
                    try:
                        self.write_new_row.writerow(self.var)

                    except:
                        self.var = {"Email_Address": key.lower()}
                        self.write_new_row.writerow(self.var)
        print(len(self.count_addr))
        return len(self.count_addr)
