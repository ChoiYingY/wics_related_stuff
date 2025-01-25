import pandas as pd
import re
import os
import sys
import xlsxwriter

'''
    Helper script to compile attendance for WiCS events.

    Created by Mendy Wu, 2018-2019
    Edited by Cynthia Lee, Christina Low 2019
    Edited by Arpita Abrol 2020
    Edited by Sammi Lin 2021
    Edited by Choi Ying Yau 2023-2024
'''

__copyright__  = 'Copyright 2023-24, Women in Computer Science(WiCS) @ SBU'

ATTENDANCE_DIR = ''         # Set as command line arg[1]
OUTPUT = ''                 # Set as command line arg[2]
emails = []                 # Set of unique emails collected
VOTING_RESULTS = None       # Final dataframe to write to OUTPUT file

'''
    Validate existences for required columns from a given excel file (.xlsx)
    Required columns includes
        1) 'First Name'
        2) 'Last Name'
        3) 'Year'
        4) 'Email'
'''
def validate_file_columns(filename, dataframe):
    required_columns = ['First Name', 'Last Name', 'Year', 'Email']

    for column in required_columns:
        if column not in dataframe:
            raise Exception(f"'{column}' column not found from file named '{filename}'.\n-> Possible reasons: the column doesn't exist / might have extra space(s) / capitalization problems in name, or it simply has a different name (e.g. Email Address, Academic Year).")

'''
    Compile names from a given excel file (.xlsx)
    Assume all sheets have columns names 'First Name', 'Last Name', 'Year', and 'Email'

    Note:
    - If the required column(s) is not found from excel sheet, please check all column names from current excel sheet when you receive such error!
'''
def get_data_from(file):
    # Retrieve name/year information, which will be updated later
    names = VOTING_RESULTS['Name'].values
    years = VOTING_RESULTS['Year'].values

    # Read excel sheet as dataframe & validate its required columns' existence
    df = pd.read_excel(file)
    validate_file_columns(file, df)
    
    for i in df.index:
        # Obtain student first name, last name to create capitalized full name
        first_name = df['First Name'][i].strip().title()
        last_name = df['Last Name'][i].strip().title()
        full_name = first_name + ' ' + last_name

        # Obtain student email & use it to find its corresponding position in result array
        email_address = df['Email'][i].strip().lower()
        index = emails.index(email_address.lower())

        # Only update name / year of a student if such information hasn't set up for them
        names[index] = full_name if not names[index] else names[index]
        years[index] = df['Year'][i].strip() if not years[index] else years[index]

    # Store & update information
    VOTING_RESULTS['Name'] = names
    VOTING_RESULTS['Year'] = years

'''
    Compile emails from a given excel file (.xlsx)
    Assume all sheets have column name 'Email'

    Note:
    - If 'Email' column is not found from excel sheet, please check column name for emails from current excel sheet when you receive such error!
'''
def get_emails_from(file):
    global emails

    # Read excel sheet as dataframe & validate if column 'Email' exists
    df = pd.read_excel(file)
    if 'Email' not in df:
        raise Exception(f"'Email' column not found from file named '{file}'.\n-> Possible reasons: the column doesn't exist / might have extra space(s) / capitalization problems in name, or it simply has a different name (e.g. Email Address).")
 
    # Obtain & store each unique student email
    for i in df.index:
        email_address = df['Email'][i].strip().lower()
        if email_address == None:
            break
        if email_address not in emails:
            emails.append(email_address)

'''
    Populate attendance values for event in given excel file (.xlsx)
    Values: [1 = attended, 0 = no attendance]

    Assume all filenames are named by one of the following formats, ignoring case:
        1) event_name.xlsx
        2) Copy of event_name.xlsx
        3) Copy of event_name (Responses).xlsx
        4) event_name (Responses).xlsx

    Also assume all ':' in file names have converted to '_' when downloading xlsx file from drive
'''
def populate_attendance_for(file):
    # Make all values 0 by default
    attendance = [0] * len(emails)

    # Obtain file name (excluding directory, ex. dir/file.xlsx -> file.xlsx)
    file_name = file.split('/')[-1]

    # Make sure the file ends with '.xlsx'
    if not file_name.endswith('.xlsx'):
        raise Exception(f"File '{file_name}' does not have an '.xlsx' extension.")

    # Set up regular expression pattern to parse event name (<- <event_name>.xlsx)
    pattern = r'^(?:copy of\s*)?(.*?)(?:\s*\(responses\))?\s*\.xlsx$'

    # Obtain the desired group '(.*?)' & set it as event_name based on regex pattern
    match = re.search(pattern, file_name, re.IGNORECASE)
    if match:
        event_name = match.group(1).strip()
    else:
        raise Exception(f"failed to parse event name by regex. Please check name format for current file '{file}'.")

    # Replace _ as : because 
    event_name = event_name.replace('_', ':')

    print(f'Processing event: {event_name}...')

    # Read excel sheet as dataframe & validate its required columns' existence
    df = pd.read_excel(file)
    validate_file_columns(file, df)

    # Set column to collected values
    for i in df.index:
        email_address = df['Email'][i].strip().lower()
        attendance[emails.index(email_address)] = 1

    # Set column to collected values
    VOTING_RESULTS[event_name] = attendance

# Reads and validates cmd args into global variables
def read_cmd_args():
    try:
        global ATTENDANCE_DIR
        global OUTPUT

        # Attempt to read cmd args
        ATTENDANCE_DIR = sys.argv[1]
        OUTPUT = sys.argv[2] + '.xlsx'

        # Create empty spreadsheet if it doesnt already exist
        workbook = xlsxwriter.Workbook(OUTPUT)
        if workbook:
            workbook.add_worksheet(name='Active Roster')
            workbook.close()

        # Resulting voting member spreadsheet
        global VOTING_RESULTS
        VOTING_RESULTS = pd.read_excel(OUTPUT, sheet_name='Active Roster')
    except IndexError:
        raise Exception('please make sure you have valid input parameters.\nUSAGE: python3 wics_voters.py <ATTENDANCE_DIR> <OUTPUT_FILENAME>')

# Main thread executed here
if __name__== '__main__':
    try:
        read_cmd_args()

        # Make sure directory exists & display error if not
        if not os.path.exists(ATTENDANCE_DIR):
            raise Exception(f"provided directory '{ATTENDANCE_DIR}' doesn't exist.")

        # Populate all 'unique' names from attendance directory
        for file in os.listdir(ATTENDANCE_DIR):
            # Define path to open & read the current excel sheet
            if file.endswith('.xlsx'):
                input_file_path = os.path.join(ATTENDANCE_DIR, file)
                get_emails_from(input_file_path)

        print(f'Total Members Found: {str(len(emails))}\n')

        # Set up voting result dataframe by providing all emails & its corresponding name/year as empty string by default
        VOTING_RESULTS['Email'] = emails
        VOTING_RESULTS['Name'] = [''] * len(emails)
        VOTING_RESULTS['Year'] = [''] * len(emails)

        # Populate attendance for all sheets
        for file in os.listdir(ATTENDANCE_DIR):
            if file.endswith('.xlsx'):
                input_file_path = os.path.join(ATTENDANCE_DIR, file)
                get_data_from(input_file_path)
                populate_attendance_for(input_file_path)

        # Sum of total #events that each students have attended
        event_columns = [col for col in VOTING_RESULTS.columns if col != 'Email' and col != 'Name' and col != 'Year']
        VOTING_RESULTS['Total #events attended'] = VOTING_RESULTS[event_columns].eq(1).sum(axis=1)
        
        # Sort result by total #events attended in descending order
        VOTING_RESULTS = VOTING_RESULTS.sort_values(by='Total #events attended', ascending=False)

        ### [Optional part] ###
        # Please feel free to remove this part or reference 'wics_voters_2324.py'
        # Re-arrange dataframe column display order by email, name, year, all events in chronological order (manual input) & total #events attended
        display_col_order = [
            'Email', 'Name', 'Year',
            # Enter event name in order - ex 'GBM#1: title', 'Workathon#1: title', ...
            'Total #events attended'
        ]
        VOTING_RESULTS = VOTING_RESULTS[display_col_order]
        ### [End of optional part] ###

        # Write result to output excel sheet
        writer = pd.ExcelWriter(OUTPUT, engine='xlsxwriter')
        VOTING_RESULTS.to_excel(writer, sheet_name='Active Roster', index=False)
        writer.close()

        print(f'\nFinal Dataframe:\n{VOTING_RESULTS}')
        print(f"\nResult has been saved to '{OUTPUT}' successfully!")
    except Exception as e:
        print(f'ERROR: {e}')
        sys.exit()