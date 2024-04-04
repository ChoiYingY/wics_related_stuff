import pandas as pd
import sys

'''
    Helper script to count all event attendance.

    Created by Choi Ying Yau, 2023
'''

__copyright__  = 'Copyright 2023, Women in Computer Science(WiCS) @ SBU'


def count_event_attendance(attendance_sheet):
    # read all event columns from given attendance_sheet
    df = pd.read_excel(attendance_sheet, sheet_name='Active Roster')
    event_columns = [col for col in df.columns if col != 'Email' and col != 'Name' and col != 'Year' and col != 'Total #events attended']

    # list total attendance of each event
    for event_name in event_columns:
        print(f'Event: {event_name} - {df[event_name].sum()}')
    
'''
    Main function to count attendance for each event from given excel sheet
'''
if __name__== '__main__':
    try:
        if len(sys.argv) < 2:
            raise Exception('Usage: python3 event_attendance.py <attendance_sheet>')
        else:
            attendance_sheet = sys.argv[1].strip()
            count_event_attendance(attendance_sheet)
    except Exception as e:
        print(f'ERROR: {e}')
        sys.exit()