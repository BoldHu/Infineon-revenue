import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import os
from datetime import datetime, timedelta

class operator(object):
    def __init__(self, GUI):
        self.revord_path = GUI.revord_path.get()
        self.ship_to_path = GUI.ship_to_path.get()
        self.sold_to_path = GUI.sold_to_path.get()
        self.allocation_path = GUI.allocation_path.get()
        self.dn_path = GUI.dn_path.get()
        self.stock_path = GUI.stock_path.get()
        self.save_path = GUI.save_path.get()
        
        # memory the last working day and holiday
        self.last_working_day = GUI.calendar.selected_last_working_day
        self.holiday = GUI.calendar.selected_holiday
        
        # this excel file is '.xlsx' format, so we can use pandas to read it directly
        self.revord_df = pd.read_excel(self.revord_path)
        self.dn_df = pd.read_excel(self.dn_path)
        
        # this excel file is '.xls' format, so we need to read it by pandas
        self.sold_to_df = pd.read_excel(self.sold_to_path)
        self.ship_to_df = pd.read_excel(self.ship_to_path)
        self.allocation_df = pd.read_excel(self.allocation_path)
        self.CPN_df = pd.read_excel(self.allocation_path, sheet_name=1)
    
    def sold_to_check(self):
        # Convert 'Sold To No.' column in revord_df to string and remove leading zeros
        self.revord_df['Sold To No.'] = self.revord_df['Sold To No.'].astype(str).str.lstrip('0')

        # Add a new column 'shipping point', 'CPN', 'DDL block' with default value '' and 'EETT', 'ETT' with 0
        self.revord_df['shipping point'] = ''
        self.revord_df['CPN'] = ''
        self.revord_df['EETT'] = 0
        self.revord_df['ETT'] = 0
        self.revord_df['DDL block'] = ''
        self.revord_df['Stock'] = 0
        self.revord_df['Proposed PGI'] = ''
        self.revord_df['Remark'] = ''
        self.revord_df['Arrange stock'] = ''

        # Convert 'SoldTo' column in sold_to_df to string for consistent comparison
        self.sold_to_df['SoldTo'] = self.sold_to_df['SoldTo'].astype(str)

        # Iterate through each row in revord_df
        for index, row in self.revord_df.iterrows():
            # Check if modified 'Sold To No.' is in 'SoldTo' column of sold_to_df
            if row['Sold To No.'] in self.sold_to_df['SoldTo'].values:
                # Update 'DDL block' column
                self.revord_df.at[index, 'DDL block'] = 'sold to block'
    
    def ship_to_check(self):
        # Convert 'Ship To No.' column in revord_df to string and remove leading zeros
        self.revord_df['Ship To No.'] = self.revord_df['Ship To No.'].astype(str).str.lstrip('0')

        # Convert 'ShipTo' column in ship_to_df to string for consistent comparison
        self.ship_to_df['ShipTo'] = self.ship_to_df['ShipTo'].astype(str)

        # Iterate through each row in revord_df
        for index, row in self.revord_df.iterrows():
            # Check if modified 'Ship To No.' is in 'ShipTo' column of ship_to_df
            if row['Ship To No.'] in self.ship_to_df['ShipTo'].values:
                # Check if 'DDL block' column already has a value
                if pd.notna(self.revord_df.at[index, 'DDL block']) and self.revord_df.at[index, 'DDL block'] != '':
                    # Concatenate new comment with existing comment
                    self.revord_df.at[index, 'DDL block'] += '; ship to block'
                else:
                    # Update 'DDL block' column with new comment
                    self.revord_df.at[index, 'DDL block'] = 'ship to block'
    
    def allocation_check(self):
        # Iterate through each row in revord_df
        for index, row in self.revord_df.iterrows():
            # Prepare the comment to be added
            comment_to_add = f"{row['Material entered']} {row['Plant']} block"

            # Check if there is a matching row in allocation_df
            if any((self.allocation_df['SP'] == row['Material entered']) & (self.allocation_df['Plant'] == row['Plant'])):
                # Check if 'DDL block' column already has a value
                if pd.notna(self.revord_df.at[index, 'DDL block']) and self.revord_df.at[index, 'DDL block'] != '':
                    # Concatenate new comment with existing comment
                    self.revord_df.at[index, 'DDL block'] += '; ' + comment_to_add
                else:
                    # Update 'DDL block' column with new comment
                    self.revord_df.at[index, 'DDL block'] = comment_to_add
    
    def add_dn_infro(self):
        # Iterate through each row in revord_df
        for index, revord_row in self.revord_df.iterrows():
            # Find matching rows in dn_df
            matching_rows = self.dn_df[
                (self.dn_df['Sales Doc.'] == revord_row['Sales Document']) &
                (self.dn_df['Item'] == revord_row['Sales Document Item'])
            ]

            # If there's a match, update the relevant columns in revord_df
            if not matching_rows.empty:
                # Assuming the first matching row is the relevant one
                matching_row = matching_rows.iloc[0]

                # Update 'shipping point', 'EETT', 'ETT', 'CPN' in revord_df with values from dn_df
                self.revord_df.at[index, 'shipping point'] = matching_row['ShPt']
                # check nan, padding with 0
                if pd.notna(matching_row['EETT']):
                    self.revord_df.at[index, 'EETT'] = matching_row['EETT']
                if pd.notna(matching_row['ETT']):
                    self.revord_df.at[index, 'ETT'] = matching_row['ETT']
                self.revord_df.at[index, 'CPN'] = matching_row['Customer Material Number']
    
    def dn_check(self):
        # Iterate through each row in revord_df
        for index, row in self.revord_df.iterrows():
            # Prepare the comment to be added
            comment_to_add = 'CPN+Plant block'

            # Get the first and second column names of CPN_df
            first_column = self.CPN_df.columns[0]
            second_column = self.CPN_df.columns[1]

            # Check if there is a matching row in self.CPN_df
            if any((self.CPN_df[first_column] == row['CPN']) & (self.CPN_df[second_column] == row['Plant'])):
                # Check if 'DDL block' column already has a value
                if pd.notna(self.revord_df.at[index, 'DDL block']) and self.revord_df.at[index, 'DDL block'] != '':
                    # Concatenate new comment with existing comment
                    self.revord_df.at[index, 'DDL block'] += '; ' + comment_to_add
                else:
                    # Update 'DDL block' column with new comment
                    self.revord_df.at[index, 'DDL block'] = comment_to_add
    
    def add_stock(self):
        # Ensure that 'Stock' column exists in revord_df
        if 'Stock' not in self.revord_df.columns:
            self.revord_df['Stock'] = 0

        # Iterate through each row in revord_df
        for index, revord_row in self.revord_df.iterrows():
            # Find matching rows in stock_df
            matching_rows = self.stock_df[
                (self.stock_df['SP'] == revord_row['Material entered']) &
                (self.stock_df['Plnt'] == revord_row['Plant']) &
                (self.stock_df['Material'] == revord_row['Material'])
            ]

            # If there's a match, update the 'Stock' value in revord_df
            if not matching_rows.empty:
                # Assuming the first matching row is the relevant one
                matching_row = matching_rows.iloc[0]

                # Update 'Stock' in revord_df with 'Sum of stock' from stock_df
                self.revord_df.at[index, 'Stock'] = matching_row['FREEQTY']
    
    def cal_proposed_day(self):
        def cal_date(self, date, dw):
            # Convert the string date to a datetime object for comparison
            start_date = datetime.strptime(date, '%Y-%m-%d').date()
            # Assuming self.holiday is a list of holidays in 'YYYY-MM-DD' format
            holidays = [datetime.strptime(day, '%Y-%m-%d').date() for day in self.holiday]
            days_to_subtract = dw
            while days_to_subtract > 0:
                start_date -= timedelta(days=1)
                # Skip weekends and holidays
                if start_date.weekday() >= 5 or start_date in holidays:
                    continue
                else:
                    days_to_subtract -= 1
            return start_date
        
        # Convert the string date to a datetime object for comparison
        last_working_day = datetime.strptime(self.last_working_day[0], '%Y-%m-%d').date()
        # Iterate through each row in revord_df
        for index, row in self.revord_df.iterrows():
            # Convert 'Goods Issue Date' to datetime.date for comparison
            goods_issue_date = datetime.strptime(row['Goods Issue Date'], '%Y-%m-%d').date()
            dw = int(row['Del Windows Minus'])
            crd = datetime.strptime(row['Customer requested date'], '%Y-%m-%d').date()
            # the EETT and ETT is like '1,00' and '2,00', convert it to int 1, 2...
            eett = int(row['EETT'].split(',')[0])
            ett = int(row['ETT'].split(',')[0])

            # Check the conditions and update 'Remark' and 'Proposed PGI' accordingly
            if goods_issue_date <= last_working_day:
                self.revord_df.at[index, 'Remark'] = 'Open AT'
                self.revord_df.at[index, 'Proposed PGI'] = goods_issue_date
            else:
                if cal_date(goods_issue_date, dw) <= last_working_day:
                    if crd-ett-eett <= last_working_day:
                        self.revord_df.at[index, 'Remark'] = 'Open AT'
                        self.revord_df.at[index, 'Proposed PGI'] = cal_date(last_working_day, dw) # last_working_day - dw
                    else:
                        self.revord_df.at[index, 'Remark'] = 'DW potential'
                        self.revord_df.at[index, 'Proposed PGI'] = cal_date(goods_issue_date, dw) # goods_issue_date - dw
                else:
                    if crd - ett - eett <= last_working_day:
                        self.revord_df.at[index, 'Remark'] = 'Due CRD with late GI'
                        self.revord_df.at[index, 'Proposed PGI'] = cal_date(date=crd, dw=ett+eett) # crd - ett - eett
                    else:
                        if cal_date(date=crd, dw=ett+eett) <= last_working_day:
                            self.revord_df.at[index, 'Remark'] = 'CRD potential with late GI'
                            self.revord_df.at[index, 'Proposed PGI'] = cal_date(date=crd, dw=ett+eett) # crd - ett - eett
                        else:
                            self.revord_df.at[index, 'Remark'] = 'No potential'
                            self.revord_df.at[index, 'Proposed PGI'] = None
    
    def arrange_stock(self):
        pass

    def save(self):
        # modify the self.revord_df the format of 'Customer requested date' and 'Goods Issue Date' and 'Delivery Date' to yyyy/mm/dd
        self.revord_df['Customer requested date'] = self.revord_df['Customer requested date'].dt.strftime('%Y/%m/%d')
        self.revord_df['Goods Issue Date'] = self.revord_df['Goods Issue Date'].dt.strftime('%Y/%m/%d')
        self.revord_df['Delivery Date'] = self.revord_df['Delivery Date'].dt.strftime('%Y/%m/%d')
        # write the self.revord_df to excel by specific path and format
        # create a new excel file
        wb = Workbook()
        # create a new sheet
        ws = wb.active
        # set the sheet name
        ws.title = 'RevOrd'
        # set the font
        font = Font(name='Arial', size=10)
        # set the writer
        writer = pd.ExcelWriter(os.path.join(self.save_path, 'infineon revenue'), engine='openpyxl')
        writer.book = wb
        # write the self.revord_df to excel
        self.revord_df.to_excel(writer, sheet_name='RevOrd', index=False)
        # set the column width
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column].width = adjusted_width
            
        # set the header style grey and center and bold
        for cell in ws[1]:
            cell.fill = PatternFill(start_color='808080', end_color='808080', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(bold=True)
            
        # save the excel file
        writer.save()
        # close the excel file
        writer.close()