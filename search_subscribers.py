import csv
import sys
import os

'''
    Helper functions
'''
# Read response & save a contact list of students who wants to subscribe WiCS newsletter
def addSubscribers(event_response_csv):
    new_subscriber_list = []
    with open(event_response_csv, newline='') as fp:
        csv_reader = csv.reader(fp)
        
        # Iterate through the rows in the CSV file
        for fields in csv_reader:
            email = fields[1].strip()
            last_name = fields[2].strip()
            first_name = fields[3].strip()
            will_subscribe = fields[6].strip()

            if will_subscribe.lower() == 'yes':
                subscriber = [email, first_name, last_name]
                new_subscriber_list.append(subscriber)

    return new_subscriber_list

# Save contact of all subscribers into a separate csv file
def saveSubscriberContact(subscriber_list, file_name):
    output_directory = 'subscribers'

    # Create the output directory if it doesn't exist
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    output_csv_file_path = os.path.join(output_directory, f'{file_name}_subscriber.csv')

    # Write contact to output CSV file
    with open(output_csv_file_path, mode='w', newline='') as csvfile:
        csv_writer = csv.writer(csvfile)

        csv_writer.writerow(['Email Address', 'First Name', 'Last Name'])
        for subscriber in subscriber_list:
            csv_writer.writerow(subscriber)

    print('All future subscriber contacts have been saved.')

'''
    Main function to save subscribers' contact based on response csv file
'''
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python3 search_subscribers.py <event_response_csv>')
        exit(1)
    else:
        event_response_csv = sys.argv[1].strip()
        file_path = event_response_csv.split('.')[0]
        file_name = file_path.split('/')[-1]
        subscriber_list = addSubscribers(event_response_csv)
        saveSubscriberContact(subscriber_list, file_name)