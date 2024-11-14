import sys
import os
import pandas as pd

# Check if file argument is provided
if len(sys.argv) <= 1:
    print('Enter file name!')
    exit(1)

# File path from command-line arguments
file_path = sys.argv[1]

# Check if the file exists
if not os.path.exists(file_path):
    print("File does not exist")
    exit(1)

# Extract file name and extension
file_name, file_extension = os.path.splitext(file_path)

# Define source and destination alphabets for transliteration
if len(sys.argv) == 2:
    source = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
    destination = 'АБСДЕФГХИЖКЛМНОПКРСТУВВЙИЗабсдефгхижклмнопкрстуввйиз'
else:
    source = sys.argv[2]
    destination = sys.argv[3]

if len(source) != len(destination):
    print('Wrong parameters!')
    exit(1)

# Define the change function for transliteration
def change(param):
    for i in range(len(param)):
        index = source.find(param[i])
        if index != -1:
            param = param[:i] + destination[index] + param[i+1:]
    return param

# Read the Excel file
df = pd.read_excel(file_path, header=None)

# Apply the change function to each string cell
for index, row in df.iterrows():
    for column in df.columns:
        if isinstance(row[column], str):
            df.at[index, column] = change(row[column])

# Save the modified DataFrame
df.to_excel(f"{file_name}_new{file_extension}", sheet_name="Sheet1", index=False)
