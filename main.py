# Project Owner: Vasoya Hiren
# Contact: vasoyahiren01@gmail.com

from tkinter import*
import os
from tkinter import messagebox
import pyodbc
import threading
import requests
from tkcalendar import Calendar
from datetime import datetime
import tkinter as tk
from tkinter import ttk
# ============main============================

class Sync_App:
    def __init__(self, root):
        self.root = root
        self.root.geometry("1150x800+0+0")
        self.root.title("Superworks Sync")
        bg_color = "#f5f6fa"
        title = Label(self.root, text="Superworks Sync", font=('times new roman', 30, 'bold'), pady=2, bd=12, bg="white", fg="Black", relief=GROOVE)
        title.pack(fill=X)
        self.api_endpoint = "http://localhost:3130/destination"
    
    # ==============Customer==========================
    
        self.c_key_id = StringVar()
        self.c_host = StringVar()
        self.c_port = StringVar()
        self.c_db_name = StringVar()
        self.c_user_name = StringVar()
        self.c_password = StringVar()
        self.c_table = StringVar()
        self.c_punch_id = StringVar()
        self.c_punch_date = StringVar()
        self.c_index = StringVar()
        self.cursor = None
        self.conn = None  # Initialize the connection attribute
        self.interval_timer = None  # Initialize the interval timer
        self.data_array = []  # Initialize the array to store data
        self.connection_details = {}
        # =============connection retail details======================
    
        F1 = LabelFrame(self.root, text="Connection Details", font=('times new roman', 14, 'bold'), bd=10, fg="Black", bg="#f5f6fa")
        F1.place(x=0, y=80, relwidth=1)

        c_key_id_lbl = Label(F1, text="Api Key:", bg=bg_color, font=('times new roman', 10, 'bold'))
        c_key_id_lbl.grid(row=0, column=0, padx=20, pady=5)
        ckey_id_txt = Entry(F1, width=50, textvariable=self.c_key_id, font='arial 10', bd=7, relief=GROOVE)
        ckey_id_txt.grid(row=0, column=1, pady=5, padx=10)

        chost_lbl = Label(F1, text="Host:", bg="#f5f6fa", font=('times new roman', 10, 'bold'))
        chost_lbl.grid(row=1, column=0, padx=20, pady=5)
        chost_txt = Entry(F1, width=50, textvariable=self.c_host, font='arial 10', bd=7, relief=GROOVE)
        chost_txt.grid(row=1, column=1, pady=5, padx=10)

        cport_lbl = Label(F1, text="Port:", bg="#f5f6fa", font=('times new roman', 10, 'bold'))
        cport_lbl.grid(row=2, column=0, padx=20, pady=5)
        cport_txt = Entry(F1, width=50, textvariable=self.c_port, font='arial 10', bd=7, relief=GROOVE)
        cport_txt.grid(row=2, column=1, pady=5, padx=10)

        cdb_name_lbl = Label(F1, text="Database:", bg="#f5f6fa", font=('times new roman', 10, 'bold'))
        cdb_name_lbl.grid(row=3, column=0, padx=20, pady=5)
        cdb_name_txt = Entry(F1, width=50, textvariable=self.c_db_name, font='arial 10', bd=7, relief=GROOVE)
        cdb_name_txt.grid(row=3, column=1, pady=5, padx=10)

        cuser_name_lbl = Label(F1, text="Username:", bg="#f5f6fa", font=('times new roman', 10, 'bold'))
        cuser_name_lbl.grid(row=4, column=0, padx=20, pady=5)
        cuser_name_txt = Entry(F1, width=50, textvariable=self.c_user_name, font='arial 10', bd=7, relief=GROOVE)
        cuser_name_txt.grid(row=4, column=1, pady=5, padx=10)

        cpassword_lbl = Label(F1, text="Password:", bg="#f5f6fa", font=('times new roman', 10, 'bold'))
        cpassword_lbl.grid(row=5, column=0, padx=20, pady=5)
        cpassword_txt = Entry(F1, width=50, textvariable=self.c_password, font='arial 10', bd=7, relief=GROOVE)
        cpassword_txt.grid(row=5, column=1, pady=5, padx=10)

        ctable_lbl = Label(F1, text="Table:", bg="#f5f6fa", font=('times new roman', 10, 'bold'))
        ctable_lbl.grid(row=6, column=0, padx=20, pady=5)
        ctable_txt = Entry(F1, width=50, textvariable=self.c_table, font='arial 10', bd=7, relief=GROOVE)
        ctable_txt.grid(row=6, column=1, pady=5, padx=10)

        cpunch_id_lbl = Label(F1, text="Punch Id:", bg="#f5f6fa", font=('times new roman', 10, 'bold'))
        cpunch_id_lbl.grid(row=7, column=0, padx=20, pady=5)
        cpunch_id_txt = Entry(F1, width=50, textvariable=self.c_punch_id, font='arial 10', bd=7, relief=GROOVE)
        cpunch_id_txt.grid(row=7, column=1, pady=5, padx=10)

        cpunch_date_lbl = Label(F1, text="Punch Date:", bg="#f5f6fa", font=('times new roman', 10, 'bold'))
        cpunch_date_lbl.grid(row=8, column=0, padx=20, pady=5)
        cpunch_date_txt = Entry(F1, width=50, textvariable=self.c_punch_date, font='arial 10', bd=7, relief=GROOVE)
        cpunch_date_txt.grid(row=8, column=1, pady=5, padx=10)

        cindex_lbl = Label(F1, text="Index No:", bg="#f5f6fa", font=('times new roman', 10, 'bold'))
        cindex_lbl.grid(row=9, column=0, padx=20, pady=5)
        cindex_txt = Entry(F1, width=50, textvariable=self.c_index, font='arial 10', bd=7, relief=GROOVE)
        cindex_txt.grid(row=9, column=1, pady=5, padx=10)

        submit_btn = Button(F1, text="Submit", command=self.update_credentials, width=10, bd=7, font=('arial', 12, 'bold'), relief=GROOVE)
        submit_btn.grid(row=10, column=1, pady=5, padx=10)

        # =================BillArea======================
    
        F2 = Frame(self.root, bd=10, relief=GROOVE)
        F2.place(x=610, y=120, width=350, height=380)

        bill_title = Label(F2, text="Connection details", font='arial 10 bold', bd=7, relief=GROOVE)
        bill_title.pack(fill=X)
        scroll_y = Scrollbar(F2, orient=VERTICAL)
        self.txtarea = Text(F2, yscrollcommand=scroll_y.set)
        scroll_y.pack(side=RIGHT, fill=Y)
        scroll_y.config(command=self.txtarea.yview)
        self.txtarea.pack(fill=BOTH, expand=1)
        
         # =======================ButtonFrame=============
    
        F3 = LabelFrame(self.root, text="Action Area", font=('times new roman', 14, 'bold'), bd=10, fg="Black", bg="#f5f6fa")
        F3.place(x=0, y=600, height=180, relwidth=1)

        # =======Buttons-======================================
    
        btn_f = Frame(F3, bd=7, relief=GROOVE)
        btn_f.place(x=160, width=400, height=105)

        self.stop_interval_btn = Button(F3, command=self.stop_interval, text="Stop Interval", width=15, bd=7, font=('arial', 10, 'bold'), relief=GROOVE)
        self.stop_interval_btn.grid(row=0, column=0, padx=2, pady=5)
        self.stop_interval_btn.grid_forget()

        self.start_interval_btn = Button(F3, command=self.start_interval, text="Start Interval", width=15, bd=7, font=('arial', 10, 'bold'), relief=GROOVE)
        self.start_interval_btn.grid(row=1, column=0, padx=2, pady=5)
        self.start_interval_btn.grid_forget()

        self.clear_btn = Button(F3, command=self.clear_data, text="Delete Credentials", width=15, bd=7, font=('arial', 10, 'bold'), relief=GROOVE)
        self.clear_btn.grid(row=2, column=0, padx=2, pady=5)

        exit_btn = Button(F3, command=self.exit_app, text="Exit", width=15, bd=7, font=('arial', 10, 'bold'), relief=GROOVE)
        exit_btn.grid(row=3, column=0, padx=2, pady=5)

        # Create a label for the loader
        self.loader_label = Label(btn_f, text="", font=('times new roman', 15, 'bold'))
        self.loader_label.grid(row=0, column=0, pady=5, padx=2)

        self.data_label = Label(btn_f, text="", font=("times new roman", 10, 'bold'))
        self.data_label.grid(row=1, column=0, pady=5, padx=5)


        btn_f2 = Frame(F3, bd=7, relief=GROOVE)
        btn_f2.place(x=580, width=500, height=105)
        
        # Create button to open calendar popup
        btn_from_calendar = ttk.Button(btn_f2, text="Start Date", command=self.open_calendar_popup_from)
        btn_from_calendar.grid(row=0, column=0, padx=0, pady=0)

        self.from_date = Label(btn_f2, text="", font=('times new roman', 12, 'bold'))
        self.from_date.grid(row=0, column=1, padx=0, pady=0)

        btn_to_calendar = ttk.Button(btn_f2, text="End Date", command=self.open_calendar_popup_to)
        btn_to_calendar.grid(row=1, column=0, padx=0, pady=0)

        self.to_date = Label(btn_f2, text="", font=('times new roman', 12, 'bold'))
        self.to_date.grid(row=1, column=1, padx=0, pady=0)

        self.resync_btn = Button(btn_f2, command=self.manually_fetch_sql_data, text="Resync", width=15, font=('arial', 12, 'bold'))
        self.resync_btn.grid(row=2, column=1, padx=0, pady=0)

        self.find_credentials()

    def update_credentials(self):
            self.txtarea.delete('1.0', END)
            self.txtarea.insert(END, "\tWelcome to my sync app")
            self.txtarea.insert(END, f"\nKeyId:{self.c_key_id.get()}")
            self.txtarea.insert(END, f"\nHost:{self.c_host.get()}")
            self.txtarea.insert(END, f"\nPort:{self.c_port.get()}")
            self.txtarea.insert(END, f"\nDatabase:{self.c_db_name.get()}")
            self.txtarea.insert(END, f"\nUserName:{self.c_user_name.get()}")
            self.txtarea.insert(END, f"\nPassword:{self.c_password.get()}")
            self.txtarea.insert(END, f"\nTableName:{self.c_table.get()}")
            self.txtarea.insert(END, f"\nPunchId:{self.c_punch_id.get()}")
            self.txtarea.insert(END, f"\nPunchDate:{self.c_punch_date.get()}")
            self.txtarea.insert(END, f"\nIndex:{self.c_index.get()}")
            self.txtarea.insert(END, f"\n================================")
            self.txtarea.insert(END, f"\nConnection String:")
            self.save_credentials()

    def save_credentials(self):
        op = messagebox.askyesno("Save Details", "Do you want to save the credentials?")
        if op > 0:
            self.sync_details = self.txtarea.get('1.0', END)
            f1 = open("credentials/config.txt", "w")
            f1.write(self.sync_details)
            f1.close()
            messagebox.showinfo("Saved", "Credentials Saved Successfully")

            self.c_key_id.set('')
            self.c_host.set('')
            self.c_port.set('')
            self.c_db_name.set('')
            self.c_user_name.set('')
            self.c_password.set('')
            self.c_table.set('')
            self.c_punch_id.set('')
            self.c_punch_date.set('')
            self.c_index.set('')
            
            # remove old connection
            self.connection_details = {}
            self.cursor = None
            if(self.conn != None):
                self.conn.close()
            if(self.interval_timer != None):
                self.stop_interval()

            self.connect_to_database()
        else:
           return
        
    def find_credentials(self):
        present = "no"
        credentials_dir = "credentials"
        if not os.path.exists(credentials_dir):
            os.mkdir(credentials_dir) # Show a message box indicating that the directory was created
            # messagebox.showinfo("Info", "Credentials directory created")  

        for i in os.listdir("credentials/"):
            if i.split('.')[0] == 'config':
                with open(f"credentials/{i}", "r") as f1:
                    self.txtarea.delete("1.0", END)
                    for d in f1:
                        self.txtarea.insert(END, d)
                    present = "yes"
                break  # No need to continue looping once the config file is found
        if present == "yes":
            self.auto_connect = True
            self.connect_to_database()
        else:
            self.clear_btn.grid_forget()
            
    def connect_to_database(self):
        try:
            with open('credentials/config.txt', 'r') as file:
                for line in file:
                    # Split each line by ':'
                    parts = line.strip().split(':')
                    if len(parts) == 2:
                        key = parts[0].strip()
                        value = parts[1].strip()
                        self.connection_details[key] = value

            # Replace the connection string with your actual connection details
            self.conn = pyodbc.connect(f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.connection_details['Host']};DATABASE={self.connection_details['Database']};UID={self.connection_details['UserName']};PWD={self.connection_details['Password']}")
            # messagebox.showinfo("Success", "Connected to database successfully!")
            # Execute a SELECT query
            self.cursor = self.conn.cursor()
            #   Show hide button
            self.clear_btn.grid(row=2, column=0, padx=5, pady=5)
            # Schedule the function to run at regular intervals
            self.schedule_interval()

            self.start_interval_btn.grid_forget()
            self.stop_interval_btn.grid(row=0, column=0, padx=5, pady=5)

            # Update the loader label periodically
            self.update_loader()

            # conn.close()
        except pyodbc.Error as e:
            messagebox.showinfo("Success", f"Failed to connect to database: {e}")

    def schedule_interval(self):
        # Define the function to be executed at regular intervals
        def interval_function():
            # Call the function you want to execute at intervals
            self.fetch_sql_data()
            # Schedule the next call after a delay (60 seconds for 1 minute interval)
            self.interval_timer = threading.Timer(60, interval_function)
            self.interval_timer.start()

        # Start the first call of the function
        interval_function()

    def fetch_sql_data(self):
        try:
            last_index = False
            rows = []
            with open("credentials/last_index.txt", "r") as text_file:
                last_index = int(text_file.read())

            if(last_index != False):
                self.cursor.execute(f"SELECT {self.connection_details['PunchDate']}, {self.connection_details['Index']}, {self.connection_details['PunchId']} FROM {self.connection_details['TableName']} WHERE {self.connection_details['Index']} > {last_index} ORDER BY {self.connection_details['Index']} DESC")
                rows = self.cursor.fetchall()
            else:
                self.cursor.execute(f"SELECT {self.connection_details['PunchDate']}, {self.connection_details['Index']}, {self.connection_details['PunchId']} FROM {self.connection_details['TableName']} WHERE {self.connection_details['PunchDate']} >= DATEADD(day, DATEDIFF(day, 0, GETDATE()), 0) ORDER BY {self.connection_details['PunchDate']} DESC")
                rows = self.cursor.fetchall()

            if(len(rows)):
                # Process the fetched rows
                formatted_rows = []
                for row in rows:
                    formatted_row = {
                        "IndexNo": row[1],  # Assuming the IndexNo is in the second column
                        "fingerId": row[2],  # Assuming the fingerId (PunchId) is in the third column
                        "time": row[0].strftime("%Y-%m-%d %H:%M:%S")  # Assuming the time (PunchDate) is in the first column
                    }
                    formatted_rows.append(formatted_row)

                # Create JSON object
                json_data = {
                    "auth": self.connection_details['KeyId'],
                    "rows": formatted_rows
                }

                # Send the data using POST request to the API endpoint
                response = requests.post(self.api_endpoint, json=json_data)

                # Check the response status
                if response.status_code == 200:
                    print("Data sent successfully.")
                    # Fetch the last index from formatted_rows
                    last_index = formatted_rows[-1]["IndexNo"]
                    print(last_index)

                    # Write the last index to the text file
                    with open("credentials/last_index.txt", "w") as text_file:
                        text_file.write(str(last_index))

                    self.data_array.append(len(rows))
                    # If the length of the array exceeds 10, remove the first element
                    if len(self.data_array) > 10:
                        self.data_array.pop(0)
                    self.data_label.config(text="Last Send: " + ", ".join(map(str, self.data_array)))

                else:
                    print(f"Failed to send data. Status code: {response.status_code}")
                
        except pyodbc.Error as e:
            print(e)
            # ===========exit=======================
    
    def exit_app(self):
        op = messagebox.askyesno("Exit", "Do you really want to exit?")
        if op > 0:
            self.stop_interval()
            self.root.destroy()

    def stop_interval(self):
        print('interval stoped successfully!')
        # Check if the interval timer is running
        if self.interval_timer and self.interval_timer.is_alive():
            # Cancel the interval timer
            self.interval_timer.cancel()
            self.interval_timer = None
            
            self.stop_interval_btn.grid_forget()
            self.start_interval_btn.grid(row=1, column=0, padx=5, pady=5)
            self.loader_label.grid_forget()

    def start_interval(self):
        print('interval started successfully!')
        # Check if the interval timer is not running
        if not (self.interval_timer and self.interval_timer.is_alive()):
            # Start the interval
            self.schedule_interval()
            # Show the stop button and hide the start button
            self.start_interval_btn.grid_forget()
            self.stop_interval_btn.grid(row=0, column=0, padx=5, pady=5)
            self.loader_label.grid(row=0, column=0, pady=5, padx=5)
            
    def clear_data(self):
        op = messagebox.askyesno("Clear", "Do you really want to Clear?")
        if op > 0:
            self.c_key_id.set('')
            self.c_host.set('')
            self.c_port.set('')
            self.c_db_name.set('')
            self.c_user_name.set('')
            self.c_password.set('')
            self.c_table.set('')
            self.c_punch_id.set('')
            self.c_punch_date.set('')
            self.c_index.set('')
            self.connection_details = {}
            self.cursor = None
            if(self.conn != None):
                self.conn.close()
            if(self.interval_timer != None):
                self.stop_interval()

            self.conn = None
            self.txtarea.delete("1.0", END)

            # Hide the clear button
            self.clear_btn.grid_forget()
            self.start_interval_btn.grid_forget()

            # Check if the config file exists
            if os.path.exists("credentials/config.txt"):
                # Delete the config file
                os.remove("credentials/config.txt")
                messagebox.showinfo("Success", "credentials delete successfully!")
            else:
                messagebox.showinfo("Success", "credentials does not exists")

    def update_loader(self):
        # Update the loader label text
        loader_text = self.loader_label.cget("text")
        if len(loader_text) < 10:
            loader_text += "."
        else:
            loader_text = "sync"
        self.loader_label.config(text=loader_text)

        # Schedule the next update after 200 milliseconds
        self.loader_label.after(200, self.update_loader)

    def open_calendar_popup_from(self):
        self.open_calendar_popup(True)

    def open_calendar_popup_to(self):
        self.open_calendar_popup(False)

    def open_calendar_popup(self, is_from_date=True):
        global popup
        popup = tk.Toplevel(root)
        popup.title("Select Date and Time")
        print(is_from_date)

        cal_frame = ttk.Frame(popup)
        cal_frame.pack(padx=10, pady=10)

        date_label = ttk.Label(cal_frame, text="Select Date:")
        date_label.grid(row=0, column=0, padx=5, pady=5)

        self.cal = Calendar(cal_frame)
        self.cal.grid(row=0, column=1, padx=5, pady=5)

        time_label = ttk.Label(cal_frame, text="Select Time:")
        time_label.grid(row=1, column=0, padx=5, pady=5)

        self.time_entry = ttk.Entry(cal_frame)
        self.time_entry.insert(0, datetime.now().strftime("%H:%M:%S"))
        self.time_entry.grid(row=1, column=1, padx=5, pady=5)

        btn_select_datetime = ttk.Button(popup, text="Select Date and Time")
        if is_from_date:
            btn_select_datetime.config(command=self.select_datetime_from)
        else:
            btn_select_datetime.config(command=self.select_datetime_to)
        
        btn_select_datetime.pack(pady=10)

    def select_datetime_from(self):
        self.select_datetime(True)

    def select_datetime_to(self):
        self.select_datetime(False)
        
    def select_datetime(self, is_from_date=True):
        selected_date = self.cal.get_date()
        selected_time = self.time_entry.get()
        # Convert the string to a datetime object
        date_obj = datetime.strptime(f"{selected_date} {selected_time}", "%m/%d/%y %H:%M:%S")
        # Format the datetime object as per your requirement
        formatted_date = date_obj.strftime("%Y-%m-%d %H:%M:%S")

        if(is_from_date): 
            self.from_date.config(text=formatted_date)
        else:
            self.to_date.config(text=formatted_date)

        print("Selected date and time:",is_from_date, formatted_date)
        popup.destroy()
    
    def manually_fetch_sql_data(self):
        try:
            print(self.from_date.cget("text"))
            print(self.to_date.cget("text"))

            self.cursor.execute(f"SELECT {self.connection_details['PunchDate']}, {self.connection_details['Index']}, {self.connection_details['PunchId']} FROM {self.connection_details['TableName']} WHERE {self.connection_details['PunchDate']} BETWEEN '{self.from_date.cget('text')}' AND '{self.to_date.cget('text')}' ORDER BY {self.connection_details['PunchDate']} DESC")

            rows = self.cursor.fetchall()

            self.to_date.config(text="")
            self.from_date.config(text="")

            if(len(rows)):
                # Process the fetched rows
                formatted_rows = []
                for row in rows:
                    formatted_row = {
                        "IndexNo": row[1],  # Assuming the IndexNo is in the second column
                        "fingerId": row[2],  # Assuming the fingerId (PunchId) is in the third column
                        "time": row[0].strftime("%Y-%m-%d %H:%M:%S")  # Assuming the time (PunchDate) is in the first column
                    }
                    formatted_rows.append(formatted_row)

                # Create JSON object
                json_data = {
                    "auth": self.connection_details['KeyId'],
                    "rows": formatted_rows
                }

                # Send the data using POST request to the API endpoint
                response = requests.post(self.api_endpoint, json=json_data)

                # Check the response status
                if response.status_code == 200:
                    messagebox.showinfo("Success", "Data sent successfully")
                    print("Manually Data sent successfully.")
                else:
                    messagebox.showerror("Error", f"Customer Failed to send data")
                    print(f"Manually Failed to send data. Status code: {response.status_code}")
            else: 
                messagebox.showinfo("Success", "Data Not Exists")
                
        except pyodbc.Error as e:
            print(e)
            # ===========exit=======================
    
root = Tk()
obj = Sync_App(root)
root.mainloop()