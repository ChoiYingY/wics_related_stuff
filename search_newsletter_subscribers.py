import csv
import sys
import os
import pandas as pd

'''
    Helper script to search all members who say yes when asked to be added on mail list.

    Created by Choi Ying Yau, 2023
'''

__copyright__  = 'Copyright 2023, Women in Computer Science(WiCS) @ SBU'


# Read response & save a contact list of students who wants to subscribe WiCS newsletter
def addSubscribers(event_response_csv):
    df = pd.read_csv(event_response_csv)

    # save email & name of all members who agreed to subscribe
    if 'Do you want to be added to the WiCS mailing list?' in df:
        df = df[df['Do you want to be added to the WiCS mailing list?'] == 'Yes']
    else:
        df = df[df['Would you like to be added to the WiCS mailing list?'] == 'Yes']
    
    df.rename(columns={'Email': 'Email Address'}, inplace=True)
    df.rename(columns={'First name': 'First Name'}, inplace=True)
    df.rename(columns={'Last name': 'Last Name'}, inplace=True)
    df = df[['Email Address', 'First Name', 'Last Name']]

    return df

# Save contact of all subscribers into a separate csv file
def saveSubscriberContact(subscriber_list, file_name):
    output_directory = 'subscribers'

    # Create the output directory if it doesn't exist
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    output_csv_file_path = os.path.join(output_directory, f'{file_name}_subscriber.csv')

    subscriber_list.to_csv(output_csv_file_path, index=False)
    print('All future subscriber contacts have been saved.')

# Obtain array of all csv files under given directory name
def getCSVs(csv_dir):
    # Get all csv files from the dir
    files = [
            f for f in os.listdir(csv_dir)
            if os.path.isfile(os.path.join(csv_dir, f)) and
            f.endswith('.csv')
        ]
    return files
'''
    Main function to save subscribers' contact based on response csv file
'''
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python3 search_newsletter_subscribers.py <event_response_csv>')
        exit(1)
    else:
        # Get directory consisting all event responses from arg
        csv_dir = sys.argv[1].strip()

        # Get all CSVs from there
        files = getCSVs(csv_dir)

        # read & search from each of them
        for event_response_csv in files:
            file_path = event_response_csv.split('.')[0].strip()
            file_name = file_path.split('/')[-1].strip()
            subscriber_list = addSubscribers(os.path.join(csv_dir, event_response_csv))
            print(f'\nnew_subscriber_list:\n{subscriber_list}\n')
            saveSubscriberContact(subscriber_list, file_name)