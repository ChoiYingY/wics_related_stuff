# About
Purpose of each Python file:
<ol>
    <li>search_subscribers.py is for reading & searching CSV for new subscribers who answered 'yes' when asked to subscribe WiCS newsletter</li>
    <li>basic_wics_voters.py is for keeping track of each WiCS member's attendance in every WiCS event w/ no particular sort order</li>
    <li>wics_voters_2324.py is for keeping track of each WiCS member's attendance in every WiCS event in 2023-2024 w/ event time order </li>
    <li>event_attendance.py is for keeping track of the total attendance of each WiCS event</li>
    <li>count_members.py is for counting all general body members & active members each year.</li>
    <li>active_members.py is for keeping track of all WiCS active members who have attended at least N events</li>
    <li>active_freshmen.py is for keeping track of all WiCS active members who are freshmen & have attended at least N events</li>
    <li>check_election_vote.py is for count all votes from eligible active members who don't violate voting rules.</li>
</ol>

<br>*Note: output of wics_voters.py will be the attendance_sheet used in other .py file, ex. event_attendance, count_members, active_members, etc.

# Usage menu
Usage: python3 search_subscribers.py <event_response_csv_directory>
<br>Usage: python3 basic_wics_voters.py <attendance_responses_directory> <output_attendance_sheet_filename>
<br>Usage: python3 wics_voters_2324.py <attendance_responses_directory> <output_attendance_sheet_filename>
<br>Usage: python3 event_attendance.py <attendance_sheet>
<br>Usage: python3 count_members.py <attendance_sheet> <min_num_events_attended>
<br>Usage: python3 active_members.py <attendance_sheet> <min_num_events_attended>
<br>Usage: python3 active_freshmen.py <attendance_sheet> <min_num_events_attended>
<br>Usage: python3 check_election_vote.py <voting_response_xlsx>

# Example
python3 search_subscribers.py csv_responses
<br>python3 basic_wics_voters.py xlsx_responses 'Member Attendance'
<br>python3 wics_voters_2324.py xlsx_responses 'Member Attendance'
<br>python3 event_attendance.py 'Member Attendance.xlsx'
<br>python3 count_members.py 'Member Attendance.xlsx' 2
<br>python3 active_members.py 'Member Attendance.xlsx' 2
<br>python3 active_freshmen.py 'Member Attendance.xlsx' 2
<br>python3 check_election_vote.py 'President Voting Form (Responses).xlsx'
