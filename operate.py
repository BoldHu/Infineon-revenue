import pandas as pd
import re

class operator(object):
    def __init__(self, GUI):
        self.revord_path = GUI.revord_path.get()
        self.ship_to_path = GUI.ship_to_path.get()
        self.sold_to_path = GUI.sold_to_path.get()
        self.allocation_path = GUI.allocation_path.get()
        self.dn_path = GUI.dn_path.get()
        self.zm_path = GUI.zm_path.get()
        self.save_path = GUI.save_path.get()
        
        self.revord_df = pd.read_excel(self.revord_path)
        self.ship_to_df = pd.read_excel(self.ship_to_path)
        self.sold_to_df = pd.read_excel(self.sold_to_path)
        self.allocation_df = pd.read_excel(self.allocation_path)
        self.dn_df = pd.read_excel(self.dn_path)
        self.zm_df = self.repair_zm(self)
        
    def repair_zm(self):
        try:
            df = pd.read_csv(self.zm_path, delim_whitespace=True, error_bad_lines=False, warn_bad_lines=True)
            df.to_excel('read_text_to_dataframe.xlsx', index=False)
            return df
        except Exception as e:
            return f'Error reading file: {e}'
    
    def sold_to_check(self):
        pass
    
    def ship_to_check(self):
        pass
    
    def allocation_check(self):
        pass
    
        