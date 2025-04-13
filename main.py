"""
This is the main file for the project. This file will be used to run the project.
The main() function is the entry point for the project.
"""

# Importing the required libraries for the project
import os
import sys

# Adding the code directory to the path so that the code can be imported
sys.path.append('./code')
from Extract_Data_from_Gmail import get_data_from_Gmail
from Find_Duplicated_Files import find_duplicates
from Write_Data_from_Gmail_to_Excel import write_data_to_excel


# Check if the directory exists or not
def check_for_dir(path):
    if not os.path.exists(path):
        # os.makedirs(path)

        # print(f'Directory {path} created')

        print(f'Directory {path} is not eixst.')

        exit(1)
    else:
        print(f'Directory {path} already exists.')

# Importing the required libraries for the project
def main():
    # # get_data_from_Gmail()

    # # Get the messages from the Gmail account using the get_data_from_Gmail function from the Extract_Data_from_Gmail.py file
    # mail_count = int(input("Enter the number of emails to fetch: "))
    # msgs = get_data_from_Gmail(mail_count)

    # # Write the data to the Excel file using the write_data_to_excel function from the Write_Data_from_Gmail_to_Excel.py file
    # write_data_to_excel(msgs)

    # Get the path to the directory from the user
    path = input('Enter the path to the directory: ')

    # Check if the directory exists or not
    check_for_dir(path)

    # Find the duplicate files in the directory using the find_duplicates function from the Find_Duplicate_Files.py file
    find_duplicates(path)

# The following is the standard boilerplate that calls the main() function. 
if __name__ == "__main__":
    main()