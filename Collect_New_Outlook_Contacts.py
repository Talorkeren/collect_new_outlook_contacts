import os
import threading
from tkinter.filedialog import askopenfile
import customtkinter
from export_data_from_outlook import OutLook
from output_file import ContactsFile

outlook = OutLook()
contacts_file = ContactsFile()


class Gui(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.days_back_get = None
        self.file_contacts_from_csv_file = None
        self.thread_run = None
        self.iconbitmap("icon.ico")
        self.title("Collect_New_Outlook_Contacts")
        self.geometry(f"{360}x{520}")
        customtkinter.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
        customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.input_frame = customtkinter.CTkFrame(self, width=130, corner_radius=0)
        self.input_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.input_frame.grid_rowconfigure(7, weight=1)

        self.logo_label = customtkinter.CTkLabel(self.input_frame, text="Cropx", text_color='#3474eb', font=customtkinter.CTkFont(size=60, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=10, columnspan=2)
        self.title_label = customtkinter.CTkLabel(self.input_frame, text="", text_color='#5e8a71', font=customtkinter.CTkFont(size=20, weight="bold"))
        self.title_label.grid(row=1, column=0, padx=20, pady=20, columnspan=2)

        self.start_button = customtkinter.CTkButton(self.input_frame, text="Start", command=self.start_button_def)
        self.start_button.grid(row=2, column=0, padx=10, pady=10, columnspan=2)

        self.days_back_label = customtkinter.CTkLabel(self.input_frame, text="Write how many days were passed from the last run", text_color='white', font=customtkinter.CTkFont(size=12))
        self.days_back_label.grid(row=3, column=0, padx=20, pady=(20, 1), columnspan=2)
        self.days_back = customtkinter.CTkEntry(self.input_frame, width=320, justify="center")
        self.days_back.grid(row=4, column=0, padx=20, pady=(1, 30), columnspan=2)
        self.days_back.insert(0, "60")

        self.Total_Contacts = customtkinter.CTkLabel(self.input_frame, text="Total Contacts In Old File: 0", text_color='#5e8a71', font=customtkinter.CTkFont(size=18, weight="bold"))
        self.Total_Contacts.grid(row=5, column=0, padx=20, pady=(50, 2), columnspan=2)
        self.new_contacts = customtkinter.CTkLabel(self.input_frame, text="New Contacts: 0", text_color='#5e8a71', font=customtkinter.CTkFont(size=18, weight="bold"))
        self.new_contacts.grid(row=6, column=0, padx=20, pady=(2, 5), columnspan=2)
        self.total_new_contacts = customtkinter.CTkLabel(self.input_frame, text="", text_color='#5e8a71', font=customtkinter.CTkFont(size=18, weight="bold"))
        self.total_new_contacts.grid(row=7, column=0, padx=20, pady=(5, 30), columnspan=2)

        self.version_label = customtkinter.CTkLabel(self.input_frame, text_color='#3474eb', text="Ver: 0.11    (28.11.2023)", anchor="ne")
        self.version_label.grid(row=8, column=0, padx=10, pady=(3, 3), columnspan=2)

    def start_button_def(self):
        self.days_back_get = self.days_back.get()
        self.start_button.configure(state="disabled")
        self.new_contacts.configure(text=f'New Contacts: 0')
        self.days_back.configure(state="disabled")
        self.run_thread()

    def thread_print(self):
        self.file_contacts_from_csv_file = None
        self.file_contacts_from_csv_file = askopenfile(mode='r', filetypes=[('CSV File', '*csv')])
        filepath = os.path.abspath(self.file_contacts_from_csv_file.name)
        self.title_label.configure(text="Read CSV file")
        self.Total_Contacts.configure(text=f'Total Contacts In Old File: {contacts_file.read_contacts_file(filepath)}')
        self.title_label.configure(text="Copy and rename the input file")
        contacts_file.copy_file_and_rename(filepath)
        self.title_label.configure(text="Read all meetings from Outlook")
        meets = outlook.connect_to_outlook()
        self.title_label.configure(text="Filtering the meetings")
        list_of_contacts = outlook.get_messages(meets, self.days_back_get)
        self.title_label.configure(text="Append new contacts")
        count_addr = contacts_file.append_contacts(list_of_contacts)
        self.new_contacts.configure(text=f'New Contacts: {count_addr}')
        self.total_new_contacts.configure(text=f'Total Contacts In New File: {count_addr + contacts_file.read_contacts_file(filepath)}')
        self.start_button.configure(state="normal")
        self.days_back.configure(state="normal")
        self.title_label.configure(text="Done, report is ready")

    def run_thread(self):
        self.thread_run = threading.Thread(target=run, args=(self,))
        self.thread_run.start()
        # thread_id = threading.get_ident()
        return self.thread_run


gui = Gui()


def run(gui: Gui):
    # print("run_def " + str(threading.get_ident()))
    gui.thread_print()


gui.mainloop()

