"""
This is the main file for the project. This file will be used to run the project.
The main() function is the entry point for the project.
"""

# Importing the required libraries for the project
import sys

# Adding the code directory to the path so that the code can be imported
sys.path.append('./code')
from Extract_Data_from_Gmail import get_data_from_Gmail
from Write_Data_from_Gmail_to_Excel import write_data_to_excel


# Importing the required libraries for the project
def main():
    # get_data_from_Gmail()

    # Get the messages from the Gmail account using the get_data_from_Gmail function from the Extract_Data_from_Gmail.py file
    mail_count=100
    msgs = get_data_from_Gmail(mail_count)

    # Write the data to the Excel file using the write_data_to_excel function from the Write_Data_from_Gmail_to_Excel.py file
    write_data_to_excel(msgs)

# The following is the standard boilerplate that calls the main() function. 
if __name__ == "__main__":
    main()