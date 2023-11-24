import pandas as pd

def read_text_to_dataframe(file_path):
    try:
        df = pd.read_csv(file_path, delim_whitespace=True, error_bad_lines=False, warn_bad_lines=True)
        df.to_excel('dataframe.xlsx', index=False)
        return df
    except Exception as e:
        return f'Error reading file: {e}'

# Usage
# df = read_text_to_dataframe('path_to_your_txt_file.txt')
# print(df)

# Usage
df = read_text_to_dataframe('Rev activity/1.txt')
