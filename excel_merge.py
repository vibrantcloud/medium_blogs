import pandas as pd
from pathlib import Path
from xlrd import XLRDError
""" 
A simple script to find excel files in a target 
location and merge them into a single file.

You need Python installed along with Pandas.

pathlib is available in Python 3.4 + 

error handling added.

"""


excel_path = r'C:\Users\ibn_k\Documents\Excels' # add your excel path r string added for Windows users.

sheet_name = 'Sheet1' # add in your sheet name 

target_path = r'C:\Users\ibn_k\Documents\Excels'

# create empty list for excel files.

excel_files = []

for file in Path(excel_path).glob('*.xlsx'):
    excel_files.append(file)


# if you want to check your files un comment the following code

# for excel in excel_files:
#     print(excel.name)
    
# print(len(excel_files)) # returns the number of files in your list.


# create empty list to store each individual dataframe.
excel_dataframe = [] 

# loop through our file to read each file and append it to our list.

for file in excel_files:
    try:
        df = pd.read_excel(file,sheet_name=sheet_name)
        df.columns = df.columns.str.lower() # lowercase all columns
        df.columns = df.columns.str.strip() # remove any trailing or leading white space.
        excel_dataframe.append(df)
    except XLRDError as err:
        print(f"{err} in {file.name}, skipping")
        
try:
    final_dataframe = pd.concat(excel_dataframe)
    final_dataframe.to_excel(target_path + '\master_file.xlsx',index=False)
    
    print(f"File Saved to {target_path}")

except ValueError as err_2:
    print(f"No Sheets Matched in any of your excel files, are you sure {sheet_name} is correct?")

