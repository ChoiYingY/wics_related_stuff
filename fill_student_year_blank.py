import sys
import pandas as pd

'''
    Helper script to use known student information from 'Member Attendance.xlsx' to fill up all blank values in 'Year' column.
    We use this when 'Year' is not provided in an event attendance spreadsheet.

    Created by Choi Ying Yau, 2024
'''

__copyright__  = 'Copyright 2024, Women in Computer Science(WiCS) @ SBU'

'''
    Helper function to fill all blanks in attended students' year in an event
'''
def fill_blank(event_df, member_attendance_df):
    # Sometimes the 'Email' col of event_df might be capitalized, convert it to lowercase to match member_attendance_df email 
    event_df['Email'] = event_df['Email'].str.lower()

    # We merge the data based on 'Email' col
    # Use 'left' merge to ensure all entries in target_df are retained and supplemented by matching entries from source_df
    updated_df = pd.merge(event_df, member_attendance_df[['Email', 'Year']], on='Email', how='left', suffixes=('', '_source'))

    # Update the 'Year' col in the target dataframe
    # We use data from source_df if the Year in target_df is NaN
    updated_df['Year'] = updated_df.apply(lambda row: row['Year_source'] if pd.isna(row['Year']) else row['Year'], axis=1)

    # Remove the temporary '_source' column used for the merge
    updated_df.drop(columns=['Year_source'], inplace=True)

    # Fill up the rest unknown year
    updated_df.fillna('Unknown', inplace=True)

    return updated_df

'''
    Main function to read member attendance excel sheet to fill up blank year in an event attendance spreadsheet.
'''
if __name__== '__main__':
    try:
        if len(sys.argv) != 3:
            raise Exception('Usage: python3 fill_student_year_blank.py <event_attendance_sheet> <member_attendance_sheet>')
        else:
            # read xlsx & convert them into dataframe
            event_attendance_sheet = sys.argv[1].strip()
            member_attendance_sheet = sys.argv[2].strip()

            event_df = pd.read_excel(event_attendance_sheet)
            member_attendance_df = pd.read_excel(member_attendance_sheet)

            # fill in the blank
            updated_event_df = fill_blank(event_df, member_attendance_df)

            # Get event name
            event_name = event_attendance_sheet.split('.')[0]

            # Display the updated DataFrame
            print(f"\nUpdated DataFrame for {event_name}:")
            print(updated_event_df)

            # Update the DataFrame stuff back to the given event Excel file
            updated_event_df.to_excel(event_attendance_sheet, index=False)
    except Exception as e:
        print(f'ERROR: {e}')
        sys.exit()