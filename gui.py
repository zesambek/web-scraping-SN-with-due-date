import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import subprocess
import os
import pandas as pd

from HTCinventoryWithSN  import write_due_date
from loginPage import login_to_ethiopian_airlines


class Application(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.username = None
        self.password = None
        self.selected_file_path = None
        self.selected_sheet = None
        self.create_widgets()

    def create_widgets(self):
        # Create a label for the username entry field.
        self.username_label = ttk.Label(self, text="Username:",anchor='e')
        self.username_label.grid(row=0, column=0)

        # Create an entry field for the username.
        self.username_entry = ttk.Entry(self,justify = 'left')
        self.username_entry.grid(row=0, column=1)

        # Create a label for the password entry field.
        self.password_label = ttk.Label(self, text="Password:",anchor='e')
        self.password_label.grid(row=1, column=0)

        # Create an entry field for the password.
        self.password_entry = ttk.Entry(self, show="*",justify = 'left')
        self.password_entry.grid(row=1, column=1)

        # Create a button to log in.
        self.login_button = ttk.Button(self, text="Extract Due Date", command=self.login)
        self.login_button.grid(row=10, column=0, columnspan=2,sticky = "W")

        # Create a label to display error messages.
        self.error_text = tk.Text(self, height=2, state=tk.DISABLED)
        self.error_text.grid(row=12, column=0, columnspan=2)

        # Create a button to select an Excel file.
        self.select_file_button = ttk.Button(self, text="Select Excel File", command=self.select_excel_file)
        self.select_file_button.grid(row=4, column=0, columnspan=2,sticky = "W")

        # Create a label for the sheet dropdown.
        self.sheet_label = ttk.Label(self, text="Select a Sheet:")
        self.sheet_label.grid(row=6, column=0,columnspan=2)

        # Create a dropdown to select a sheet.
        self.selected_sheet = tk.StringVar()
        self.sheet_dropdown = ttk.Combobox(self, textvariable=self.selected_sheet, state="readonly")
        self.sheet_dropdown.grid(row=8, column=0,columnspan=2)



    def login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()
        
        #assign a excel file to a variable
        workbook = self.selected_file_path

        #assign the sheet name to variable ,default is None
        selected_sheet_name = None

        # disable the login button
        self.login_button.config(state=tk.DISABLED)

        #calling the login method   and assign to the  variable 
        browser = login_to_ethiopian_airlines(username, password, self.username_entry, self.password_entry, self.error_text, self.login_button)

        # Enable the login button
        self.login_button.config(state=tk.NORMAL)
        
        # Check if the login was successful before calling the write_due_date method.
        if browser is not None:
            try:

                # Check if the selected file path is not empty.
                if self.selected_file_path:

                    # Call the write_due_date function with the selected file path and sheet name.
                    selected_sheet_name = None if self.selected_sheet.get() == "All Sheets" else self.selected_sheet.get()
                    write_due_date(browser, self.selected_file_path, selected_sheet_name)
                else:
                    # Display an error message if no file is selected.
                    self.error_text.config(state=tk.NORMAL)
                    self.error_text.delete(1.0, tk.END)
                    self.error_text.insert(tk.END, "Please select an Excel file.")
                    self.error_text.config(state=tk.DISABLED)
            except:

                # Display an error message if the sending email failed.
                self.error_text.delete("1.0", tk.END)
                self.error_text.insert(tk.END, "due_date extraction  failed. Please try again.")

                # Close the browser.
                browser.quit()
        else:
            # Display an error message if the login failed.
            self.error_text.delete("1.0", tk.END)
            self.error_text.insert(tk.END, "Login failed. Please check your username and password.")
        # Clear the password field.
        self.password_entry.delete(0, tk.END)
    def select_excel_file(self):
        self.selected_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("Excel files (legacy)", "*.xls")])
        if self.selected_file_path:

            # Update the error_text widget with the selected file name
            self.error_text.config(state=tk.NORMAL)
            self.error_text.delete(1.0, tk.END)
            self.error_text.insert(tk.END, "Selected: " + self.selected_file_path.split("/")[-1])
            self.error_text.config(state=tk.DISABLED)
            # Read the Excel file and get the sheet names
            excel_file = pd.ExcelFile(self.selected_file_path)
            sheet_names = excel_file.sheet_names

            # Update the sheet dropdown with the sheet names and an "All Sheets" option
            self.sheet_dropdown["values"] = ["All Sheets"] + sheet_names
            self.selected_sheet.set("All Sheets")
        else:

            self.select_file_button.config(text="Select Excel File")

root = tk.Tk()

# Set the title of the window
root.title("HTC inventory With SN Generator")
app = Application(master=root)
app.mainloop()

