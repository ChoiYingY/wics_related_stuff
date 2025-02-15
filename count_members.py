import sys
import pandas as pd

'''
    Helper script to count all general body members + active members.

    Created by Choi Ying Yau, 2023
'''

__copyright__  = 'Copyright 2023, Women in Computer Science(WiCS) @ SBU'
year_order = ['Freshman', 'Sophomore', 'Junior', 'Senior', 'Graduate', 'Unknown']

'''
    Helper function to process dataframe: fill NaN w/ 0 & display numbers in Integer
'''
def get_year_count(df):
    year_counts = df['Year'].value_counts().reindex(year_order, fill_value=0)
    return year_counts.astype(int)

'''
    Helper function to count all general body members who have attended WiCS event(s)
'''
def count_general_body_members_from_each_year(attendance_sheet):
    df = pd.read_excel(attendance_sheet, sheet_name='Active Roster')
    
    # Count # members from each year: 'freshman', 'sophomore', 'junior', 'senior', 'graduate', etc.
    year_counts = get_year_count(df)

    # Print # members from each year
    print(f'Number of general body members from each year: {year_counts}')


'''
    Helper function to count all active members who have attended WiCS events at <min_num_events_attended> times
'''
def count_active_members_from_each_year(attendance_sheet, min_num_events_attended):
    df = pd.read_excel(attendance_sheet, sheet_name='Active Roster')

    # select rows for members who have attended events for at least N times
    df = df[df['Total #events attended'] >= min_num_events_attended]

    # Count #active members from each year: 'freshman', 'sophomore', 'junior', 'senior', etc.
    year_counts = get_year_count(df)

    # Print #active members from each year
    print(f'Number of active members from each year: {year_counts}')

'''
    Main function to count attendance for each event from given excel sheet
'''
if __name__== '__main__':
    try:
        if len(sys.argv) != 3:
            raise Exception('Usage: python3 count_members.py <attendance_sheet> <min_num_events_attended>')
        else:
            attendance_sheet = sys.argv[1].strip()
            min_num_events_attended = int(sys.argv[2].strip())

            count_general_body_members_from_each_year(attendance_sheet)
            count_active_members_from_each_year(attendance_sheet, min_num_events_attended)
    except Exception as e:
        print(f'ERROR: {e}')
        sys.exit()