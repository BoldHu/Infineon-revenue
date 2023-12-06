import tkinter as tk
from tkinter import ttk
import calendar
from datetime import datetime

class CalendarApp:
    def __init__(self, root):
        self.root = root
        self.current_year = datetime.now().year
        self.current_month = datetime.now().month
        self.selection_mode = None  # Tracks what the user is selecting
        self.selected_last_working_day = []  # Tracks the last working day selected by the user
        self.selected_holiday = []  # Tracks the holiday selected by the user

        # Create a combobox to select month
        self.month_cb = ttk.Combobox(root, values=[calendar.month_name[i] for i in range(1, 13)], state='readonly')
        self.month_cb.grid(row=6, column=6, padx=(5, 10), pady=(40, 10))
        self.month_cb.current(self.current_month - 1)
        self.month_cb.bind("<<ComboboxSelected>>", self.update_calendar)

        # Create a combobox to select year
        self.year_var = tk.IntVar()
        self.year_cb = ttk.Combobox(root, textvariable=self.year_var, values=[year for year in range(1900, 2101)], state='readonly')
        self.year_cb.grid(row=6, column=5, padx=(5, 10), pady=(40, 10))
        self.year_cb.set(self.current_year)
        self.year_cb.bind("<<ComboboxSelected>>", self.update_calendar)
        
        # Create a label and button for Last Working Day
        self.lwd_label = tk.Label(root, text="Last Working Day: None", font=('Arial', 12))
        self.lwd_label.grid(row=8, column=6, padx=(5, 10), pady=(5, 10))
        self.lwd_button = tk.Button(root, text="Select Last Working Day", command=lambda: self.set_selection_mode("LWD"))
        self.lwd_button.grid(row=8, column=5, padx=(5, 10), pady=(5, 10))

        # Create a label and button for Holiday
        self.holiday_label = tk.Label(root, text="Selected Holiday: None", font=('Arial', 12))
        self.holiday_label.grid(row=9, column=6, padx=(5, 10), pady=(5, 10))
        self.holiday_button = tk.Button(root, text="Select Holiday", command=lambda: self.set_selection_mode("Holiday"))
        self.holiday_button.grid(row=9, column=5, padx=(5, 10), pady=(5, 10))
        
        # Create a frame for the calendar
        self.calendar_frame = tk.Frame(root)
        self.calendar_frame.grid(row=7, column=5, columnspan=2, padx=(5, 10), pady=(40, 10))

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
        
    def set_selection_mode(self, mode):
        self.selection_mode = mode

    def select_date(self, year, month, day):
        selected_date = f"{day}-{month}-{year}"

        if self.selection_mode == "LWD":
            self.selected_last_working_day.append(selected_date)
            self.lwd_label.config(text=f"Last Working Day: {self.selected_last_working_day}")
        elif self.selection_mode == "Holiday":
            self.selected_holiday.append(selected_date)
            self.holiday_label.config(text=f"Selected Holiday: {self.selected_holiday}")