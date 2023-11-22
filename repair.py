import pandas as pd
import os

def repair_excel(file_path, delimiter=' '):
    # Try reading the Excel file with the specified engine
    try:
        df = pd.read_excel(file_path, header=None, engine='openpyxl')
    except Exception as e:
        print('Error reading file: {e}')
        return f'Error reading file: {e}'

    # Split the data in each cell into separate columns
    split_data = df[0].str.split(delimiter, expand=True)

    # Assuming the first row contains headers
    headers = split_data.iloc[0]
    repaired_df = pd.DataFrame(split_data.values[1:], columns=headers)

    # Save the repaired DataFrame to a new Excel file
    repaired_df.to_excel('repair_excel.xlsx', index=False)
    print('Repaired Excel saved as "repaired_excel.xlsx"')

    return 'Repaired Excel saved as "repaired_excel.xlsx"'

# Usage
# repair_excel('path_to_your_damaged_excel_file.xlsx')

# Use the function like this
# repair_excel('path_to_your_damaged_excel_file.xlsx')
repair_excel('Rev activity/ZM.xlsx')