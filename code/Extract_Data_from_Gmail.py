"""
This code extracts data from Gmail and prints it on the screen.
The code reads the credentials from the yaml file and connects to the Gmail account.
It searches for emails from a specific email address and extracts the email subject and body.
The code uses the imaplib library to connect to the Gmail account.
"""

#Import the necessary libraries
import email
import email.utils
import imaplib

import yaml


def get_credentials():
    #Read the credentials from the yaml file
    with open("./code/credentials.yml") as f:
        content = f.read()

    #Load the credentials from the yaml file
    my_credentials = yaml.load(content, Loader=yaml.FullLoader)

    #Get the user and password from the credentials file
    user, password, key, value = (
        my_credentials['user'], 
        my_credentials['password'], 
        my_credentials['key'], 
        my_credentials['value']
    )
    return user, password , key, value

def get_data_from_Gmail(mail_count):
    #Get the user and password from the credentials file
    user, password, key, value = get_credentials()

    #Define the imap_url for Gmail account
    imap_url = 'imap.gmail.com'

    # Connect to the Gmail account using the imaplib library
    my_mail = imaplib.IMAP4_SSL(imap_url)

    # Login to the Gmail account using the user and password from the credentials file
    my_mail.login(user, password)

    # Select the Inbox folder in the Gmail account to search for emails
    my_mail.select('Inbox')

    # Search for emails from a specific email address
    #For other keys (criteria): https://gist.github.com/martinrusev/6121028#file-imap-search
    _, data = my_mail.search(None, key, value)

    # data is a list of mail ids
    mail_id_list = data[0].split()

    # Create an empty list to store the messages
    msgs = []

    # Loop through the mail_id_list and fetch the messages from the Gmail account
    count = 0
    for num in mail_id_list:
        count+= 1
        if count > mail_count:
            break

        typ, data = my_mail.fetch(num, '(RFC822)') #RFC822 returns whole message (BODY fetches just body)
        msgs.append(data)

    return msgs #Return the messages to the main function for further processing

    # #Now we have all messages, but with a lot of details
    # #Let us extract the right text and print on the screen

    # #In a multipart e-mail, email.message.Message.get_payload() returns a 
    # # list with one item for each part. The easiest way is to walk the message 
    # # and get the payload on each part:
    # # https://stackoverflow.com/questions/1463074/how-can-i-get-an-email-messages-text-content-using-python

    # # NOTE that a Message object consists of headers and payloads.

    # for msg in msgs[::-1]:
    #     for response_part in msg:
    #         if type(response_part) is tuple:
    #             my_msg=email.message_from_bytes((response_part[1]))

    #             print("_________________________________________")
    #             print ("subj:", my_msg['subject'])
    #             print ("from:", my_msg['from'])
    #             print ("date received:", email.utils.parsedate_to_datetime(my_msg['date']))
    #             print ("body:")

    #             for part in my_msg.walk():  
    #                 #print(part.get_content_type())
    #                 if part.get_content_type() == 'text/plain':
    #                     print (part.get_payload())
    #         break
        