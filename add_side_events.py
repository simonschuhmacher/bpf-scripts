import pandas as pd


"""
This script checks different sources of side event registrations and
makes sure that the correct side events are then available in a custom
guest filed.
As automatic updates to a user are not possible, the script shows a link
to edit the corresponding user manually in zkipster.
"""

def check_side_events(row: pd.Series, side_event_dfs: list[tuple[str, pd.DataFrame]]) -> pd.Series:
    side_events: list[str] = []

    for list_description, list_df in side_event_dfs:
        if row['EMAIL'] in list_df['EMAIL'].values:
            side_events.append(list_description)

    if row['CAS / ALUMNI NETWORKING LUNCH'] == 'I will participate':
        side_events.append('CAS Alumni Lunch | 12:15 | Room: Forum')

    row['SIDE EVENTS'] = '; '.join(side_events)

    return row

def add_side_events(invited_filename: str):
    invited_df = pd.read_excel(invited_filename)

    side_event_lists = {
        'Side Event: UN Peace Missions': 'Guided Tour Exhibition | 12:30 | Room: Check-in, Plaza',
        'KOFF Advisory board 27/11': 'KOFF Advisory Board Lunch | 12:00 | Room: Atelier',
    }

    confirmed_bpf_df = invited_df[~invited_df['GUESTLIST'].isin(list(side_event_lists.keys()))]
    confirmed_bpf_df = confirmed_bpf_df[confirmed_bpf_df['GUEST STATUS'] == 'Confirmed']
    confirmed_bpf_df['SIDE EVENTS'] = confirmed_bpf_df['SIDE EVENTS'].fillna('')

    side_event_dfs: list[tuple[str, pd.DataFrame]] = []
    for list_name in side_event_lists:
        list_description = side_event_lists[list_name]
        list_df = invited_df[invited_df['GUESTLIST'] == list_name]
        side_event_dfs.append((list_description, list_df))

    with_new_side_events = confirmed_bpf_df.apply(lambda row: check_side_events(row, side_event_dfs=side_event_dfs), axis=1)
    
    correct_side_events_df = with_new_side_events[confirmed_bpf_df['SIDE EVENTS'] == with_new_side_events['SIDE EVENTS']]
    wrong_side_events_df = with_new_side_events[confirmed_bpf_df['SIDE EVENTS'] != with_new_side_events['SIDE EVENTS']]

    print(f'{correct_side_events_df.shape[0]}/{confirmed_bpf_df.shape[0]} guests have the correct side events set.')
    print(f'{wrong_side_events_df.shape[0]}/{confirmed_bpf_df.shape[0]} guests have the wrong side events set.')

    print('')
    print('The following guests need updates for their side events:')
    print('--------------------------------------------------------')

    for index, row in wrong_side_events_df.sort_values(by='SIDE EVENTS').iterrows():        
        print(f'https://account.zkipster.com/Guests/Edit?InvitationId={row["ID"]} {row["FIRST NAME"]} {row["LAST NAME"]} {row["EMAIL"]} :: {row["SIDE EVENTS"]}')
        print(f'Previous seleciton: {confirmed_bpf_df.at[index, "SIDE EVENTS"]}')
        print('----')


add_side_events(invited_filename='lists/invited_2024-01-23--20-19.xlsx')
