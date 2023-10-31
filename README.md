# About
Purpose of each python files:
<ol>
    <li>search_subscribers.py is for reading csv to find new subscribers who answered 'yes' when asked to subscribe WiCS newsletter</li>
    <li>wics_voters_2324.py is for keeping track of each WiCS member's attendance in every WiCS event in 2023-2024</li>
    <li>event_attendance.py is for keeping track of total attendance of each WiCS event</li>
    <li>count_members.py is for counting all general body members & active members in each year.</li>
    <li>active_members.py is for keeping track of all WiCS active members who have attended at least N events</li>
    <li>active_freshmen.py is for keeping track of all WiCS active members who are freshmen & have attended at least N events</li>
</ol>

# Usage menu
Usage: python3 search_subscribers.py <event_response_csv>
<br>Usage: python3 wics_voters_2324.py <attendance_responses_directory> <output_attendance_sheet_filename>
<br>Usage: python3 event_attendance.py <attendance_sheet>
<br>Usage: python3 count_members.py <attendance_sheet> <min_num_events_attended>
<br>Usage: python3 active_members.py <attendance_sheet> <min_num_events_attended>
<br>Usage: python3 active_freshmen.py <attendance_sheet> <min_num_events_attended>

# Example