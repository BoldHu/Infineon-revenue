import pandas as pd

class operator(object):
    def __init__(self):
        self.revord_path = 'Rev activity/ZSD_REVORD 20231115.XLSX'
        self.ship_to_path = 'Rev activity/DDue_List_ShipTo_Data 20231115.xls'
        self.sold_to_path = 'Rev activity/DDue_List_SoldTo_Data 20231115.xls'
        self.allocation_path = 'Rev activity/DDue_List_Allocation_Data 20231115.xls'
        self.dn_path = 'Rev activity/ZSD_CHIPBIZ_DN 2023115.XLSX'
        self.zm_path = 'Rev activity/ZM.xlsx'
        self.save_path = 'Rev activity/'
        
        # this excel file is '.xlsx' format, so we can use pandas to read it directly
        self.revord_df = pd.read_excel(self.revord_path)
        # self.dn_df = pd.read_excel(self.dn_path)
        
        # this excel file is '.xls' format, so we need to read it by pandas
        self.sold_to_df = pd.read_excel(self.sold_to_path)
        self.ship_to_df = pd.read_excel(self.ship_to_path)
        self.allocation_df = pd.read_excel(self.allocation_path)
        
        # self.zm_df = self.repair_zm(self)
        
    def repair_zm(self):
        try:
            df = pd.read_csv(self.zm_path, delim_whitespace=True, error_bad_lines=False, warn_bad_lines=True)
            df.to_excel('read_text_to_dataframe.xlsx', index=False)
            return df
        except Exception as e:
            return f'Error reading file: {e}'
    
    def sold_to_check(self):
        # Convert 'Sold To No.' column in revord_df to string and remove leading zeros
        self.revord_df['Sold To No.'] = self.revord_df['Sold To No.'].astype(str).str.lstrip('0')

        # Add a new column 'DDL block' with default value ''
        self.revord_df['DDL block'] = ''

        # Convert 'SoldTo' column in sold_to_df to string for consistent comparison
        self.sold_to_df['SoldTo'] = self.sold_to_df['SoldTo'].astype(str)

        # Iterate through each row in revord_df
        for index, row in self.revord_df.iterrows():
            # Check if modified 'Sold To No.' is in 'SoldTo' column of sold_to_df
            if row['Sold To No.'] in self.sold_to_df['SoldTo'].values:
                # Update 'DDL block' column
                self.revord_df.at[index, 'DDL block'] = 'sold to block'

        # save the result to excel
        self.revord_df.to_excel('Rev activity/sold_to_check.xlsx', index=False)
        print("Sold-to check completed!")

    
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
        # save the result to excel
        self.revord_df.to_excel('Rev activity/ship_to_check.xlsx', index=False)
        print("Ship-to check completed!")
    
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
                    
        # save the result to excel
        self.revord_df.to_excel('Rev activity/allocation_check.xlsx', index=False)
        print("Allocation check completed!")
        
    
    def add_dn_infro(self):
        pass
    
    def dn_check(self):
        pass
    
if  __name__ == '__main__':
    operator = operator()
    operator.sold_to_check()
    print("Sold-to check completed!")
    operator.ship_to_check()
    print("Ship-to check completed!")
    operator.allocation_check()
    print("Allocation check completed!")
    # operator.add_dn_infro()
    # print("DN information added!")
    # operator.dn_check()
    # print("DN check completed!")
    # operator.repair_zm()
    # print("ZM repaired!")
    
    print("Report generated successfully!")
        