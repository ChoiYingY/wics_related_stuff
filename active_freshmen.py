import sys
import pandas as pd

'''
    Helper script to compile all active freshmen for running Fall Freshmen ECM.

    Created by Choi Ying Yau, 2023
'''

__copyright__  = 'Copyright 2023, Women in Computer Science(WiCS) @ SBU'


'''
    Helper function to find name & email of all active members who are freshmen
'''
def find_all_active_freshmen(attendance_sheet, min_num_events_attended):
    df = pd.read_excel(attendance_sheet, sheet_name='Active Roster')
    
    # filter out members that have attended less than N events & are not freshmen
    df = df[df['Total #events attended'] >= min_num_events_attended]
    df = df[df['Year'] == 'Freshman']

    # Only keep email & name columns (+ year col for temporary use) & rename the columns
    columns_to_keep = ['Email', 'Name', 'Year']
    df = df[columns_to_keep]

    # Set email column
    df['Email Address'] = df['Email']

    # Set first&last name column, handle case fo people who have two words in first name
    first_last_name = df['Name'].str.split()
    df['First Name'] = first_last_name.str[:-1].str.join(' ')
    df['Last Name'] = first_last_name.str[-1]

    # Drop unwanted columns & preserve Email Address,First Name,Last Name columns
    df.drop(columns=['Name', 'Email', 'Year'], inplace=True)

    # export dataframe as csv file
    df.to_csv('Active Members (Freshmen only).csv', index=False)

    
'''
    Main function to find all active members from given excel sheet & export to csv
'''
if __name__== '__main__':
    try:
        if len(sys.argv) != 3:
            raise Exception('Usage: python3 active_freshmen.py <attendance_sheet> <min_num_events_attended>')
        else:
            attendance_sheet = sys.argv[1].strip()
            min_num_events_attended = int(sys.argv[2].strip())

            # find email & name for all active freshmen & export to csv
            active_freshmen = find_all_active_freshmen(attendance_sheet, min_num_events_attended)

    except Exception as e:
        print(f'ERROR: {e}')
        sys.exit()