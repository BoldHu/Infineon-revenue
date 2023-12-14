import pandas as pd

# read the infineon revenue.xlsx file in Rev activity
df = pd.read_excel('Rev activity/infineon revenue.xlsx')
# check if the column is in df
columns = ['Sales Office', 'Description', 'Sold To No.', 'Ship To No.', 'Sales Document Type', 'CPN', 'leaf seller', 'Material entered', 'Material', 'shipping point', 'Sales Document', 'Sales Document Item', 'Open Quantity', 'Customer requested date', 'Transit', 'Allocation policy', 'Open Quantity', 'Proposed PGI']
for column in columns:
    if column not in df.columns:
        print(column)