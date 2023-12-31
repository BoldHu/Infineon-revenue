import tkinter as tk
from tkinter import filedialog
from calendar_create import CalendarApp
from operate import operator

class GUI:
    def __init__(self):
        # Create main window
        self.root = tk.Tk()
        self.root.title("Infineon Revenue Report Generator")
        # full screen
        self.root.state('zoomed')
        
        # Initialize file path variables
        self.revord_path = tk.StringVar(value="")
        self.ship_to_path = tk.StringVar(value="")
        self.sold_to_path = tk.StringVar(value="")
        self.allocation_path = tk.StringVar(value="")
        self.dn_path = tk.StringVar(value="")
        self.stock_path = tk.StringVar(value="")
        self.customer_priority_path = tk.StringVar(value="")
        self.exception_path = tk.StringVar(value="")
        self.save_path = tk.StringVar(value="")

        # Maintain a list of file paths for easy access
        self.file_paths = [self.revord_path, self.ship_to_path, self.sold_to_path, self.allocation_path, self.dn_path, self.stock_path, self.customer_priority_path, self.exception_path, self.save_path]
        self.calendar = CalendarApp(self.root)
        
    def run(self):
        # Run the GUI
        self.create_button_label()
        self.root.mainloop()

    def upload_file(self, path, label):
        # Upload a file and set the path
        label.config(text="Uploading...")
        path_selected = filedialog.askopenfilename()
        path.set(path_selected)
        label.config(text=path_selected)
        
    def select_save_path(self):
        # Select the path to save the report
        self.save_path.set("Selecting...")
        path_selected = filedialog.askdirectory()
        self.save_path.set(path_selected)
        self.save_path_label.config(text=path_selected)
        
    def confirm_action(self):
        # initialize operator
        self.operator = operator(self)
        # Placeholder for confirm action
        print("Begin to generate report...")
        self.operator.sold_to_check()
        print("Sold-to check completed!")
        self.operator.ship_to_check()
        print("Ship-to check completed!")
        self.operator.allocation_check()
        print("Allocation check completed!")
        self.operator.add_dn_infro()
        print("DN information added!")
        self.operator.dn_check()
        print("DN check completed!")
        if self.exception_path.get() != "":
            self.operator.exception_check()
            print("Exception check completed!")
        self.operator.add_stock()
        print("Stock added!")
        self.operator.cal_proposed_day()
        print("Proposed PGI added!")
        self.operator.arrange_stock()
        print("Stock arranged!")
        self.operator.save()
        print("Report saved!")
        
        print("Report generated successfully!")

    def clear_action(self):
        # Clear all file paths
        for path in self.file_paths:
            path.set("")
        
        # Clear the selected_date and reset the date label
        self.calendar.selected_holiday = []
        self.calendar.holiday_label.config(text='')
        self.calendar.selected_last_working_day = []
        self.calendar.lwd_label.config(text='')
        
    def create_button_label(self):
        # Create labels and buttons for file uploads
        # First is revord Excel
        self.revord_path_label = tk.Label(self.root, textvariable=self.revord_path)
        upload_revord = lambda: self.upload_file(self.revord_path, self.revord_path_label)
        self.revord_upload_button = tk.Button(self.root, text="Upload Revord Excel", command=upload_revord)
        self.revord_upload_button.grid(row=1, column=0, padx=(5, 10), pady=(40, 10))
        self.revord_path_label.grid(row=1, column=1, padx=(5, 10), pady=(40, 10))
        
        # second is ship-to excel
        self.ship_to_path_label = tk.Label(self.root, textvariable=self.ship_to_path)
        upload_ship_to = lambda: self.upload_file(self.ship_to_path, self.ship_to_path_label)
        self.ship_to_upload_button = tk.Button(self.root, text="Upload Ship-to Excel", command=upload_ship_to)
        self.ship_to_upload_button.grid(row=3, column=0, padx=(5, 10), pady=(40, 10))
        self.ship_to_path_label.grid(row=3, column=1, padx=(5, 10), pady=(40, 10))
        
        # third is sold-to excel
        self.sold_to_path_label = tk.Label(self.root, textvariable=self.sold_to_path)
        upload_sold_to = lambda: self.upload_file(self.sold_to_path, self.sold_to_path_label)
        self.sold_to_upload_button = tk.Button(self.root, text="Upload Sold-to Excel", command=upload_sold_to)
        self.sold_to_upload_button.grid(row=2, column=0, padx=(5, 10), pady=(40, 10))
        self.sold_to_path_label.grid(row=2, column=1, padx=(5, 10), pady=(40, 10))
        
        # fourth is allocation excel
        self.allocation_path_label = tk.Label(self.root, textvariable=self.allocation_path)
        upload_allocation = lambda: self.upload_file(self.allocation_path, self.allocation_path_label)
        self.allocation_upload_button = tk.Button(self.root, text="Upload Allocation Excel", command=upload_allocation)
        self.allocation_upload_button.grid(row=4, column=0, padx=(5, 10), pady=(40, 10))
        self.allocation_path_label.grid(row=4, column=1, padx=(5, 10), pady=(40, 10))
        
        # fifth is DN excel
        self.dn_path_label = tk.Label(self.root, textvariable=self.dn_path)
        upload_dn = lambda: self.upload_file(self.dn_path, self.dn_path_label)
        self.dn_upload_button = tk.Button(self.root, text="Upload DN Excel", command=upload_dn)
        self.dn_upload_button.grid(row=5, column=0, padx=(5, 10), pady=(40, 10))
        self.dn_path_label.grid(row=5, column=1, padx=(5, 10), pady=(40, 10))

        # sixth is stock excel
        self.stock_path_label = tk.Label(self.root, textvariable=self.stock_path)
        upload_stock = lambda: self.upload_file(self.stock_path, self.stock_path_label)
        self.stock_upload_button = tk.Button(self.root, text="Upload stock Excel", command=upload_stock)
        self.stock_upload_button.grid(row=6, column=0, padx=(5, 10), pady=(40, 10))
        self.stock_path_label.grid(row=6, column=1, padx=(5, 10), pady=(40, 10))
        
        # seventh is customer priority excel
        self.customer_priority_path_label = tk.Label(self.root, textvariable=self.customer_priority_path)
        upload_customer_priority = lambda: self.upload_file(self.customer_priority_path, self.customer_priority_path_label)
        self.customer_priority_upload_button = tk.Button(self.root, text="Upload Customer Priority Excel", command=upload_customer_priority)
        self.customer_priority_upload_button.grid(row=2, column=2, padx=(5, 10), pady=(40, 10))
        self.customer_priority_path_label.grid(row=2, column=3, padx=(5, 10), pady=(40, 10))
        
        # eighth is exception excel
        self.exception_path_label = tk.Label(self.root, textvariable=self.exception_path)
        upload_exception = lambda: self.upload_file(self.exception_path, self.exception_path_label)
        self.exception_upload_button = tk.Button(self.root, text="Upload Exception Excel", command=upload_exception)
        self.exception_upload_button.grid(row=3, column=2, padx=(5, 10), pady=(40, 10))
        self.exception_path_label.grid(row=3, column=3, padx=(5, 10), pady=(40, 10))

        # Create Confirm, save and Clear buttons
        self.confirm_button = tk.Button(self.root, text="Confirm", command=self.confirm_action)
        self.clear_button = tk.Button(self.root, text="Clear", command=self.clear_action)
        self.save_button = tk.Button(self.root, text="Save", command=self.select_save_path)
        self.save_path_label = tk.Label(self.root, textvariable=self.save_path)
        self.save_button.grid(row=1, column=2, padx=(5, 10), pady=(40, 10))
        self.save_path_label.grid(row=1, column=3, padx=(5, 10), pady=(40, 10))
        self.confirm_button.grid(row=7, column=0, padx=(5, 10), pady=(40, 10))
        self.clear_button.grid(row=7, column=1, padx=(5, 10), pady=(40, 10))

   
# Create and run the GUI
gui = GUI()
gui.run()