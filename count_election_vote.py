import sys
import pandas as pd
import matplotlib.pyplot as plt

'''
    Helper script to count vote & remove votes that violate our regulations:
        - not active member
        - cannot vote for ECM for over 3 person

   Created by Choi Ying Yau, 2024
'''

__copyright__  = 'Copyright 2024, Women in Computer Science(WiCS) @ SBU'


# set up global var: mapping for vote result in (voted_name, freq) pair
result = {}

'''
    Helper function to read through current voter's choice (who they voted for) & update result map
    - ECM can vote up to 3 people, so we have to split up their choices before updating result map
        -> ex. if voted for 'A, B, C', this will each add a vote to person A, B & C
'''
def process_vote(is_voting_ecm, choice):
    # split up choices & add 1 vote to each chosen person in ECM; ignore those vote >3 candidates
    if is_voting_ecm:
        choices = choice.split(',')
        if(len(choices) <= 3):
            for choice in choices:
                choice = choice.strip()
                result[choice] = 1 if (choice not in result) else result[choice]+1
    else:
        # else add 1 vote to the person
        result[choice] = 1 if (choice not in result) else result[choice]+1


'''
    Helper function to parse through all voting responses from xlsx sheet & store voting result
'''
def parse_response(is_voting_ecm, voting_df):
    num_voter = 0
    # only go thru responses of voters who are active members
    for voter in voting_df['Email Address']:
        if voter in active_df['Email Address'].values:
            num_voter += 1
            # get the 2nd column (voting choice) from current voter & add vote based on that
            choice = voting_df[voting_df['Email Address'] == voter].iloc[0,2]
            process_vote(is_voting_ecm, choice)
        else:
            continue
    
    print(f'{num_voter} people have voted in total.')


'''
    Helper function to plot result, given plot title, x & y (candidate names + num of vote they have)
'''
def plot(title, candidates, num_votes):
    # set up plot size & title
    plt.figure(figsize =(10, 5))
    plt.title(title)

    # set up horizontal bar graph w/ #votes shown
    plt.barh(candidates, num_votes)
    for index, value in enumerate(num_votes):
        plt.text(value, index, str(value))

    # display in tight layout
    plt.tight_layout()
    plt.show()


'''
    Run the following if the script has been called
    - For making sure vote from non-active members are not counted
'''
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python3 count_election_vote.py <voting_response_xlsx>')
        exit(1)
    else:
        # Read list of active member + voting responses
        active_df = pd.read_csv('./Active Members.csv')
        voting_form_name = sys.argv[1].strip()
        voting_df = pd.read_excel(voting_form_name)

        # Determining if we are voting for ecm or not
        is_voting_ecm = "Event Committee Member" in voting_form_name

        # Read through all responses to store voting result for current position
        parse_response(is_voting_ecm, voting_df)

        # After parsing all, sort vote result after going through all responses
        sorted_result = sorted(result.items(), key=lambda x: x[1], reverse=True)
        sorted_result, sorted_counts = zip(*sorted_result)

        # Print winning result
        if is_voting_ecm:
            print(f'Winning ECMs are {sorted_result[0]}, {sorted_result[1]}, {sorted_result[2]}')
        else:
            print(f'Winning candidate is {sorted_result[0]}')
        
        # Make a plot, given the title name, x & y
        plot(title=voting_form_name.split('.')[0], candidates=sorted_result, num_votes=sorted_counts)