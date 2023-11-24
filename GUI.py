import tkinter as tk
from tkinter import filedialog
from calendar_create import CalendarApp
from operate import operator

class GUI:
    def __init__(self):
        # Create main window
        self.root = tk.Tk()
        self.root.title("Infineon Revenue Report Generator")
        self.root.geometry("1000x800")
        
        # Initialize file path variables
        self.revord_path = tk.StringVar(value="")
        self.ship_to_path = tk.StringVar(value="")
        self.sold_to_path = tk.StringVar(value="")
        self.allocation_path = tk.StringVar(value="")
        self.dn_path = tk.StringVar(value="")
        self.zm_path = tk.StringVar(value="")
        self.save_path = tk.StringVar(value="")

        # Maintain a list of file paths for easy access
        self.file_paths = [self.revord_path, self.ship_to_path, self.sold_to_path, self.allocation_path, self.dn_path, self.zm_path]
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
        label.set(path_selected)
        
    def select_save_path(self):
        # Select the path to save the report
        self.save_path.set("Selecting...")
        path_selected = filedialog.askdirectory()
        self.save_path.set(path_selected)
        self.save_path_label.set(path_selected)
        
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
        
        print("Report generated successfully!")

    def clear_action(self):
        # Clear all file paths
        for path in self.file_paths:
            path.set("")
        
        # Clear the selected_date and reset the date label
        self.calendar.selected_date = []
        self.calendar.selected_dates_label.config(text='')
        
    
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
        self.ship_to_upload_button.grid(row=2, column=0, padx=(5, 10), pady=(40, 10))
        self.ship_to_path_label.grid(row=2, column=1, padx=(5, 10), pady=(40, 10))
        
        # third is sold-to excel
        self.sold_to_path_label = tk.Label(self.root, textvariable=self.sold_to_path)
        upload_sold_to = lambda: self.upload_file(self.sold_to_path, self.sold_to_path_label)
        self.sold_to_upload_button = tk.Button(self.root, text="Upload Sold-to Excel", command=upload_sold_to)
        self.sold_to_upload_button.grid(row=3, column=0, padx=(5, 10), pady=(40, 10))
        self.sold_to_path_label.grid(row=3, column=1, padx=(5, 10), pady=(40, 10))
        
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

        # sixth is ZM excel
        self.zm_path_label = tk.Label(self.root, textvariable=self.zm_path)
        upload_zm = lambda: self.upload_file(self.zm_path, self.zm_path_label)
        self.zm_upload_button = tk.Button(self.root, text="Upload ZM Excel", command=upload_zm)
        self.zm_upload_button.grid(row=6, column=0, padx=(5, 10), pady=(40, 10))
        self.zm_path_label.grid(row=6, column=1, padx=(5, 10), pady=(40, 10))

        # Create Confirm, save and Clear buttons
        self.confirm_button = tk.Button(self.root, text="Confirm", command=self.confirm_action)
        self.clear_button = tk.Button(self.root, text="Clear", command=self.clear_action)
        self.save_button = tk.Button(self.root, text="Save", command=self.select_save_path)
        self.save_path_label = tk.Label(self.root, textvariable=self.select_save_path)
        self.save_button.grid(row=8, column=0, padx=(5, 10), pady=(40, 10))
        self.save_path_label.grid(row=8, column=1, padx=(5, 10), pady=(40, 10))
        self.confirm_button.grid(row=7, column=0, padx=(5, 10), pady=(40, 10))
        self.clear_button.grid(row=7, column=1, padx=(5, 10), pady=(40, 10))

   
# Create and run the GUI
gui = GUI()
gui.run()