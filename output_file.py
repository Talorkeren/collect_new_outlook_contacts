import os


class ContactsFile:
    def __init__(self):
        self.file_a = None
        self.data = None
        self.file = None
        self.data_contacts_from_file = None
        self.contacts_from_outlook = None
        self.read_contacts_file()

    def read_contacts_file(self):
        if os.path.exists("Output/contacts_file.txt"):
            self.file = open('Output/contacts_file.txt', "r")
            self.data = self.file.read()
            self.data_contacts_from_file = self.data.split("\n")
            print(self.data_contacts_from_file)

    def append_contacts(self, list_of_contacts):
        self.contacts_from_outlook = list_of_contacts
        for name_from_outlook in self.contacts_from_outlook:
            if not self.data_contacts_from_file:
                self.file_a = open('Output/contacts_file.txt', "+a")
                print(name_from_outlook)
                self.file_a.write(f'{name_from_outlook.lower()}\n')
            else:
                if name_from_outlook.lower() not in self.data_contacts_from_file:
                    self.file_a = open('Output/contacts_file.txt', "+a")
                    print(name_from_outlook)
                    self.file_a.write(f'{name_from_outlook.lower()}\n')
