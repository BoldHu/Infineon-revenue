import tkinter as tk
from tkinter import ttk
import calendar
from datetime import datetime

class CalendarApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Date Picker")
        self.current_year = datetime.now().year
        self.current_month = datetime.now().month

        # Create a combobox to select month
        self.month_cb = ttk.Combobox(root, values=[calendar.month_name[i] for i in range(1, 13)], state='readonly')
        self.month_cb.grid(row=0, column=0, padx=10, pady=10)
        self.month_cb.current(self.current_month - 1)
        self.month_cb.bind("<<ComboboxSelected>>", self.update_calendar)

        # Create a combobox to select year
        self.year_var = tk.IntVar()
        self.year_cb = ttk.Combobox(root, textvariable=self.year_var, values=[year for year in range(1900, 2101)], state='readonly')
        self.year_cb.grid(row=0, column=1, padx=10, pady=10)
        self.year_cb.set(self.current_year)
        self.year_cb.bind("<<ComboboxSelected>>", self.update_calendar)

        # Create a frame for the calendar
        self.calendar_frame = tk.Frame(root)
        self.calendar_frame.grid(row=1, column=0, columnspan=2)
        
        # Create a label to show selected date
        self.selected_date_label = tk.Label(root, text="", font=('Arial', 12))
        self.selected_date_label.grid(row=2, column=0, columnspan=2, pady=10)

        # Initialize calendar
        self.update_calendar()

    def update_calendar(self, event=None):
        # Remove old calendar
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()

        # Get selected year
        year = self.year_var.get()

        # Get the name of the selected month from the combobox
        selected_month_name = self.month_cb.get()

        # Find the index of the selected month
        # Since calendar.month_name is an array starting with an empty string, we start indexing from 1
        month_index = [calendar.month_name[i] for i in range(1, 13)].index(selected_month_name) + 1

        # Create a new calendar for the selected year and month
        self.cal = calendar.monthcalendar(year, month_index)
        self.create_calendar_widgets(year, month_index)

    def create_calendar_widgets(self, year, month):
        # Create headers
        headers = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        for i, header in enumerate(headers):
            label = tk.Label(self.calendar_frame, text=header, font=('Arial', 10, 'bold'))
            label.grid(row=0, column=i, padx=5, pady=5)

        # Create day buttons
        for row, week in enumerate(self.cal, start=1):
            for col, day in enumerate(week):
                if day != 0:
                    btn = tk.Button(self.calendar_frame, text=str(day), command=lambda d=day: self.select_date(year, month, d))
                    btn.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")

    def select_date(self, year, month, day):
        # Update selected date label
        self.selected_date_label.config(text=f"Selected Date: {day}-{month}-{year}")

# Create the main window
root = tk.Tk()
app = CalendarApp(root)
root.mainloop()
