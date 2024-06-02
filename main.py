
# Project Owner: Vasoya Hiren
# Contact: vasoyahiren01@gmail.com

from tkinter import *
import os
from tkinter import messagebox
import pyodbc
import pymssql
import threading
import requests
from tkcalendar import Calendar
from datetime import datetime
import tkinter as tk
from tkinter import ttk
import pystray
from PIL import Image, ImageDraw
import json
from decimal import Decimal

class DecimalEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, Decimal):
            return float(obj)
        return super(DecimalEncoder, self).default(obj)

class Sync_App:
    def __init__(self, root):
        self.root = root
        self.root.geometry("1150x800+0+0")
        self.root.title("Superworks Sync")

        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)
        self.setup_tray()

        title = Label(self.root, text="Superworks Sync", font=('times new roman', 30, 'bold'), pady=2, bd=12, bg="white", fg="Black", relief=GROOVE)
        title.pack(fill=X)

        self.api_endpoint = "http://localhost:3130/bio/pythonAppData"

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
        self.conn = None
        self.interval_timer = None
        self.data_array = []
        self.connection_details = {}

        F1 = LabelFrame(self.root, text="Connection Details", font=('times new roman', 14, 'bold'), bd=10, fg="Black", bg="#f5f6fa")
        F1.place(x=0, y=80, relwidth=1)

        self.create_label_entry(F1, "Api Key:", self.c_key_id, 0)
        self.create_label_entry(F1, "Host:", self.c_host, 1)
        self.create_label_entry(F1, "Port:", self.c_port, 2)
        self.create_label_entry(F1, "Database:", self.c_db_name, 3)
        self.create_label_entry(F1, "Username:", self.c_user_name, 4)
        self.create_label_entry(F1, "Password:", self.c_password, 5)
        self.create_label_entry(F1, "Table:", self.c_table, 6)
        self.create_label_entry(F1, "Punch Id:", self.c_punch_id, 7)
        self.create_label_entry(F1, "Punch Date:", self.c_punch_date, 8)
        self.create_label_entry(F1, "Index No:", self.c_index, 9)

        submit_btn = Button(F1, text="Submit", command=self.update_credentials, width=10, bd=7, font=('arial', 12, 'bold'), relief=GROOVE)
        submit_btn.grid(row=10, column=1, pady=5, padx=10, sticky=E)

        F2 = Frame(self.root, bd=10, relief=GROOVE)
        F2.place(x=610, y=120, width=350, height=380)

        bill_title = Label(F2, text="Connection details", font='arial 10 bold', bd=7, relief=GROOVE)
        bill_title.pack(fill=X)
        scroll_y = Scrollbar(F2, orient=VERTICAL)
        self.txtarea = Text(F2, yscrollcommand=scroll_y.set)
        scroll_y.pack(side=RIGHT, fill=Y)
        scroll_y.config(command=self.txtarea.yview)
        self.txtarea.pack(fill=BOTH, expand=1)

        F3 = LabelFrame(self.root, text="Action Area", font=('times new roman', 14, 'bold'), bd=10, fg="Black", bg="#f5f6fa")
        F3.place(x=0, y=600, height=180, relwidth=1)

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

        self.loader_label = Label(btn_f, text="", font=('times new roman', 15, 'bold'))
        self.loader_label.grid(row=0, column=0, pady=5, padx=2)

        self.data_label = Label(btn_f, text="", font=("times new roman", 10, 'bold'))
        self.data_label.grid(row=1, column=0, pady=5, padx=5)

        btn_f2 = Frame(F3, bd=7, relief=GROOVE)
        btn_f2.place(x=580, width=500, height=105)

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

    def create_label_entry(self, frame, text, variable, row):
        lbl = Label(frame, text=text, bg="#f5f6fa", font=('times new roman', 10, 'bold'))
        lbl.grid(row=row, column=0, padx=20, pady=5, sticky=W)
        txt = Entry(frame, width=50, textvariable=variable, font='arial 10', bd=7, relief=GROOVE)
        txt.grid(row=row, column=1, pady=5, padx=10, sticky=W+E)

    # Remaining methods...

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
        # self.txtarea.insert(END, f"\n================================")
        # self.txtarea.insert(END, f"\nConnection String:")
        self.save_credentials()

    def save_credentials(self):
        op = messagebox.askyesno("Save Details", "Do you want to save the credentials?")
        if op > 0:
            self.sync_details = self.txtarea.get('1.0', END)
            with open("credentials/config.txt", "w") as f1:
                f1.write(self.sync_details)
            messagebox.showinfo("Saved", "Credentials Saved Successfully")
            self.clear_fields()
            self.connection_details = {}
            self.cursor = None
            if self.conn is not None:
                self.conn.close()
            if self.interval_timer is not None:
                self.stop_interval()
            self.connect_to_database()
        else:
            return

    def clear_fields(self):
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

    def find_credentials(self):
        present = "no"
        credentials_dir = "credentials"
        if not os.path.exists(credentials_dir):
            os.mkdir(credentials_dir)
        for i in os.listdir("credentials/"):
            if i.split('.')[0] == 'config':
                with open(f"credentials/{i}", "r") as f1:
                    self.txtarea.delete("1.0", END)
                    for d in f1:
                        self.txtarea.insert(END, d)
                    present = "yes"
                break
        if present == "yes":
            self.auto_connect = True
            self.connect_to_database()
        else:
            self.clear_btn.grid_forget()

    def connect_to_database(self):
        try:
            with open('credentials/config.txt', 'r') as file:
                for line in file:
                    parts = line.strip().split(':')
                    if len(parts) == 2:
                        key = parts[0].strip()
                        value = parts[1].strip()
                        self.connection_details[key] = value

            isODBC = True
            if isODBC:
                self.conn = pyodbc.connect(f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.connection_details['Host']};DATABASE={self.connection_details['Database']};UID={self.connection_details['UserName']};PWD={self.connection_details['Password']}")
                self.cursor = self.conn.cursor()
            else:
                self.conn = pymssql.connect(server=self.connection_details['Host'], user=self.connection_details['UserName'], password=self.connection_details['Password'], database=self.connection_details['Database'], port=self.connection_details['Port'])
                self.cursor = self.conn.cursor()
                self.cursor.execute("SELECT @@version;")
                row = self.cursor.fetchone()
                print(f"SQL Server version: {row[0]}")

            self.clear_btn.grid(row=2, column=0, padx=5, pady=5)
            self.schedule_interval()
            self.start_interval_btn.grid_forget()
            self.stop_interval_btn.grid(row=0, column=0, padx=5, pady=5)
            self.update_loader()
        except pyodbc.Error as e:
            messagebox.showinfo("Success", f"Failed to connect to database: {e}")

    def schedule_interval(self):
        def interval_function():
            self.fetch_sql_data()
            self.interval_timer = threading.Timer(60, interval_function)
            self.interval_timer.start()
        interval_function()

    def fetch_sql_data(self):
        try:
            last_index = False
            rows = []
            if os.path.exists("credentials/last_index.txt"):
                with open("credentials/last_index.txt", "r") as text_file:
                    last_index = int(text_file.read())

            if last_index:
                self.cursor.execute(f"SELECT {self.connection_details['PunchDate']}, {self.connection_details['Index']}, {self.connection_details['PunchId']} FROM {self.connection_details['TableName']} WHERE {self.connection_details['Index']} > {last_index} ORDER BY {self.connection_details['Index']} DESC")
                rows = self.cursor.fetchall()
            else:
                self.cursor.execute(f"SELECT {self.connection_details['PunchDate']}, {self.connection_details['Index']}, {self.connection_details['PunchId']} FROM {self.connection_details['TableName']} WHERE {self.connection_details['PunchDate']} >= DATEADD(day, DATEDIFF(day, 0, GETDATE()), 0) ORDER BY {self.connection_details['PunchDate']} DESC")
                rows = self.cursor.fetchall()

            formatted_rows = []
            for row in rows:
                formatted_row = {
                    "IndexNo": row[1],
                    "fingerId": row[2],
                    "time": row[0].strftime("%Y-%m-%d %H:%M:%S")
                }
                formatted_rows.append(formatted_row)

            json_data = {
                "auth": self.connection_details['KeyId'],
                "rows": formatted_rows
            }

            if(len(formatted_rows)):
                json_str = json.dumps(json_data, cls=DecimalEncoder)
                response = requests.post(self.api_endpoint, data=json_str, headers={'Content-Type': 'application/json'})

                if response.status_code == 200:
                    print("Data sent successfully.")
                    if len(formatted_rows) > 0:
                        last_index = formatted_rows[-1]["IndexNo"]
                        with open("credentials/last_index.txt", "w") as text_file:
                            text_file.write(str(last_index))

                        self.data_array.append(len(rows))
                        if len(self.data_array) > 10:
                            self.data_array.pop(0)
                        self.data_label.config(text="Last Send: " + ", ".join(map(str, self.data_array)))
                else:
                    error_message = response.json().get('data', {}).get('message', 'No message found')
                    print(f"Failed to send data. Status code: {error_message}")
        except Exception as e:
            print(f"Connection error occurred: {e}")

    def exit_app(self):
        op = messagebox.askyesno("Exit", "Do you really want to exit?")
        if op > 0:
            self.stop_interval()
            self.root.destroy()
            self.quit_application()

    def stop_interval(self):
        if self.interval_timer and self.interval_timer.is_alive():
            self.interval_timer.cancel()
            self.interval_timer = None
            self.stop_interval_btn.grid_forget()
            self.start_interval_btn.grid(row=1, column=0, padx=5, pady=5)
            self.loader_label.grid_forget()

    def start_interval(self):
        if not (self.interval_timer and self.interval_timer.is_alive()):
            self.schedule_interval()
            self.start_interval_btn.grid_forget()
            self.stop_interval_btn.grid(row=0, column=0, padx=5, pady=5)
            self.loader_label.grid(row=0, column=0, pady=5, padx=5)

    def clear_data(self):
        op = messagebox.askyesno("Clear", "Do you really want to Clear?")
        if op > 0:
            self.clear_fields()
            self.connection_details = {}
            self.cursor = None
            if self.conn:
                self.conn.close()
            if self.interval_timer:
                self.stop_interval()
            self.conn = None
            self.txtarea.delete("1.0", END)
            self.clear_btn.grid_forget()
            self.start_interval_btn.grid_forget()
            if os.path.exists("credentials/config.txt"):
                os.remove("credentials/config.txt")
                messagebox.showinfo("Success", "Credentials deleted successfully!")
            else:
                messagebox.showinfo("Success", "Credentials do not exist")

    def update_loader(self):
        loader_text = self.loader_label.cget("text")
        if len(loader_text) < 10:
            loader_text += "."
        else:
            loader_text = "sync"
        self.loader_label.config(text=loader_text)
        self.loader_label.after(200, self.update_loader)

    def open_calendar_popup_from(self):
        self.open_calendar_popup(True)

    def open_calendar_popup_to(self):
        self.open_calendar_popup(False)

    def open_calendar_popup(self, is_from_date=True):
        global popup
        popup = tk.Toplevel(self.root)
        popup.title("Select Date and Time")

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
        date_obj = datetime.strptime(f"{selected_date} {selected_time}", "%m/%d/%y %H:%M:%S")
        formatted_date = date_obj.strftime("%Y-%m-%d %H:%M:%S")
        if is_from_date:
            self.from_date.config(text=formatted_date)
        else:
            self.to_date.config(text=formatted_date)
        popup.destroy()

    def manually_fetch_sql_data(self):
        try:
            print(self.from_date.cget("text"))
            print(self.to_date.cget("text"))
            self.cursor.execute(f"SELECT {self.connection_details['PunchDate']}, {self.connection_details['Index']}, {self.connection_details['PunchId']} FROM {self.connection_details['TableName']} WHERE {self.connection_details['PunchDate']} BETWEEN '{self.from_date.cget('text')}' AND '{self.to_date.cget('text')}' ORDER BY {self.connection_details['PunchDate']} DESC")
            rows = self.cursor.fetchall()
            self.to_date.config(text="")
            self.from_date.config(text="")
            if rows:
                formatted_rows = []
                for row in rows:
                    formatted_row = {
                        "IndexNo": row[1],
                        "fingerId": row[2],
                        "time": row[0].strftime("%Y-%m-%d %H:%M:%S")
                    }
                    formatted_rows.append(formatted_row)
                json_data = {
                    "auth": self.connection_details['KeyId'],
                    "rows": formatted_rows
                }
                response = requests.post(self.api_endpoint, json=json_data)
                if response.status_code == 200:
                    messagebox.showinfo("Success", "Data sent successfully")
                    print("Manually Data sent successfully.")
                else:
                    messagebox.showerror("Error", "Failed to send data")
                    print(f"Manually Failed to send data. Status code: {response.status_code}")
            else:
                messagebox.showinfo("Success", "Data Not Exists")
        except Exception as e:
            print(e)

    def hide_window(self):
        self.root.withdraw()
        self.tray_icon.visible = True

    def show_window(self, icon, item):
        self.root.deiconify()
        self.tray_icon.visible = False

    def quit_application(self):
        self.tray_icon.stop()
        self.root.quit()

    def setup_tray(self):
        image = Image.new('RGB', (64, 64), color=(73, 109, 137))
        draw = ImageDraw.Draw(image)
        draw.rectangle((10, 10, 54, 54), fill=(255, 255, 255))
        menu = pystray.Menu(
            pystray.MenuItem('Show', self.show_window),
            pystray.MenuItem('Quit', self.quit_application)
        )
        self.tray_icon = pystray.Icon("tkinter_icon", image, "Superworks Sync", menu)
        threading.Thread(target=self.tray_icon.run).start()

def main():
    root = Tk()
    app = Sync_App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
