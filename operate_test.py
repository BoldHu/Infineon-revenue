import pandas as pd

class operator(object):
    def __init__(self):
        self.revord_path = 'Rev activity/ZSD_REVORD 20231115.XLSX'
        self.ship_to_path = 'Rev activity/DDue_List_ShipTo_Data 20231115.xls'
        self.sold_to_path = 'Rev activity/DDue_List_SoldTo_Data 20231115.xls'
        self.allocation_path = 'Rev activity/DDue_List_Allocation_Data 20231115.xls'
        self.dn_path = 'Rev activity/ZSD_CHIPBIZ_DN 2023115.XLSX'
        self.zm_path = 'Rev activity/1.txt'
        self.save_path = 'Rev activity/'
        
        # this excel file is '.xlsx' format, so we can use pandas to read it directly
        self.revord_df = pd.read_excel(self.revord_path)
        self.dn_df = pd.read_excel(self.dn_path)
        
        # this excel file is '.xls' format, so we need to read it by pandas
        self.sold_to_df = pd.read_excel(self.sold_to_path)
        self.ship_to_df = pd.read_excel(self.ship_to_path)
        self.allocation_df = pd.read_excel(self.allocation_path)
        self.CPN_df = pd.read_excel(self.allocation_path, sheet_name=1)
        
        # dict to store the relationship of Sloc and whorehouse
        self.sloc_to_whr = {
            'LA00': 'WA00', 
            'LA00': 'WA01',
            'LF00': 'WF00',
            'LF00': 'WF01',
            'LS00': 'WAD0',
            'LS00': 'WAD1',
            'LH00': 'WH00',
            'LH00': 'WH01',
            'USA1': 'USA1',
            'LT00': 'WT00',
            'LT00': 'WT01',
            'LI00': 'WI00',
            'LI00': 'WI01',
            'WHC2': 'US01',
            'LF00': 'WF02',
            'LA00': 'WA02',
            'LH00': 'WH02',
            'LAU0': 'WUA0',
            'LAU0': 'WUA1',
            'LAU0': 'WUA2',
        }

    def repair_zm(self):
        try:
            # Read the .txt file
            with open(self.zm_path, 'r') as file:
                lines = file.readlines()
            
            # Process each line to get the last 8 values and create a list of lists
            data = [line.strip().split()[-8:] for line in lines]

            # Create a DataFrame from the data
            df = pd.DataFrame(data)

            # save the column 1,4,5
            df = df.iloc[:, [1, 4, 5]]

            # Drop the first line
            df = df.iloc[1:, :]

            # Rename the last three columns
            df.columns = [*df.columns[:-3], 'Stock', 'Sloc', 'Sales Product']

            # Convert 'Stock' to numeric for summation
            df['Stock'] = pd.to_numeric(df['Stock'], errors='coerce')

            # Map 'Sloc' to 'Whorehouse' using self.sloc_to_whr
            df['Whorehouse'] = df['Sloc'].map(self.sloc_to_whr)

            # Group by 'Whorehouse' and 'Sales Product' and calculate the sum of 'Stock'
            self.zm_df =  df.groupby(['Sales Product', 'Whorehouse'])['Stock'].sum().reset_index(name='Sum of stock')
            
            # save the result
            self.zm_df.to_excel('Rev activity/zm_df.xlsx')
            print('zm_df successfully')
            
        except Exception as e:
            return f'Error reading file: {e}'
    
    def sold_to_check(self):
        # Convert 'Sold To No.' column in revord_df to string and remove leading zeros
        self.revord_df['Sold To No.'] = self.revord_df['Sold To No.'].astype(str).str.lstrip('0')

        # Add a new column 'shipping point', 'CPN', 'DDL block' with default value '' and 'EETT', 'ETT' with 0
        self.revord_df['shipping point'] = ''
        self.revord_df['CPN'] = ''
        self.revord_df['EETT'] = 0
        self.revord_df['ETT'] = 0
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
        # save result
        self.revord_df.to_excel('Rev activity/add_dn_infro.xlsx', index=False)
        print('DN information added!')
    
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
        # save the result
        self.revord_df.to_excel('Rev activity/dn_check.xlsx')
        print('dn check')
    
    def cal_stock(self):
        pass
    
    def cal_proposed_day(self):
        pass
    
    def remark(self):
        pass
    
if  __name__ == '__main__':
    operator = operator()
    # operator.sold_to_check()
    # print("Sold-to check completed!")
    # operator.ship_to_check()
    # print("Ship-to check completed!")
    # operator.allocation_check()
    # print("Allocation check completed!")
    # operator.add_dn_infro()
    # print("DN information added!")
    # operator.dn_check()
    # print("DN check completed!")
    print(operator.repair_zm())
    print("ZM repaired!")
    
    print("Report generated successfully!")
        