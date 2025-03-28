# Importing the required libraries for the project
import sys

# Adding the code directory to the path so that the code can be imported
sys.path.append('./code')
import email
import email.utils
import imaplib

import openpyxl as xl  # Importing the openpyxl library to work with Excel files
import yaml
from Extract_Data_from_Gmail import get_data_from_Gmail
from openpyxl import \
    Workbook  # Importing the Workbook class from the openpyxl library
from openpyxl.utils import \
    get_column_letter  # Importing the get_column_letter function from the openpyxl library to get the column letter from the column number


def write_data_to_excel(data):
    msgs = get_data_from_Gmail() # Get the messages from the Gmail account using the get_data_from_Gmail function from the Extract_Data_from_Gmail.py file
    wb = Workbook() # Create a new Workbook object
    ws = wb.active # Get the active worksheet from the Workbook object
    ws.title = "Gmail Data" # Set the title of the worksheet to "Gmail Data"
    
    # #Now we have all messages, but with a lot of details
    # #Let us extract the right text  and write to the Excel file
    # # #In a multipart e-mail, email.message.Message.get_payload() returns a
    # # list with one item for each part. The easiest way is to walk the message
    # # and get the payload on each part:
    # # https://stackoverflow.com/questions/1463074/how-can-i-get-an-email-messages-text-content-using-python
    # # NOTE that a Message object consists of headers and payloads.

    for msg in msgs[::-1]:
        for response_part in msg:
            if type(response_part) is tuple:
                my_msg=email.message_from_bytes((response_part[1]))

                # Get the subject of the email
                subject = my_msg['subject']

                # Get the sender of the email
                sender = my_msg['from']

                # Get the date of the email
                date = email.utils.parsedate_to_datetime(my_msg['date'])
                date = date.strftime("%Y-%m-%d %H:%M:%S")

                # Get the body of the email
                for part in my_msg.walk():  
                    #print(part.get_content_type())
                    if part.get_content_type() == 'text/plain':
                        # print (part.get_payload())
                        body = part.get_payload()
                        
                # Write the data to the Excel file
                ws.append([subject, date, sender, body])

    # Save the Excel file
    wb.save("gmail_data.xlsx")