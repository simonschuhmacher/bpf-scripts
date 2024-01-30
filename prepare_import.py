import pandas as pd
import re
from datetime import datetime


"""
Create a CSV which can be imported into the BPF platform.
It adds side events from different sources and creates a column
for each side event. If a user has registred for the corresponding
side event, the value of the column will be 'x', otherwise empty (nan).
The script also prints a list of all generated side event column
names. Those need to be added to the corresponding workshop as their
'Import Value' in the BPF website CMS.
"""

def create_unique_column(row: pd.Series, col: str, col_prefix: str) -> pd.Series:
    if not pd.isna(row[col]):
        value = re.sub(r'\W', '_', row[col])
        row[f'{col_prefix}{value}'] = 'x'

    return row

def prepare_import(invited_filename: str):
    invited_df = pd.read_excel(invited_filename)
    speakers_df = pd.read_excel('badges_in/speakers.xlsx')

    side_event_lists = [
        'KOFF Advisory board 27/11',
        'Side Event: UN Peace Missions',
        'Rebooking Health'
    ]

    confirmed_bpf_df = invited_df[~invited_df['GUESTLIST'].isin(side_event_lists)]
    confirmed_bpf_df = confirmed_bpf_df[confirmed_bpf_df['GUEST STATUS'] == 'Confirmed']

    speakers_df = speakers_df[~speakers_df['EMAIL'].str.lower().isin(invited_df['EMAIL'].str.lower())]

    confirmed_bpf_df = pd.concat([confirmed_bpf_df, speakers_df])

    # Add import value of forum to assign users to the correct forum
    confirmed_bpf_df['YEAR'] = 'BPF2024'

    columns_before_workshops = confirmed_bpf_df.columns
    
    confirmed_bpf_df = confirmed_bpf_df.apply(lambda row: create_unique_column(row=row, col='FOCUS ZONE I', col_prefix='FZ1_'), axis=1)
    confirmed_bpf_df = confirmed_bpf_df.apply(lambda row: create_unique_column(row=row, col='FOCUS ZONE II', col_prefix='FZ2_'), axis=1)

    workshop_columns = confirmed_bpf_df.columns.difference(columns_before_workshops)
    print('Workshop columns:')
    print('\n'.join(workshop_columns.to_list()))
    
    current_date = datetime.now().strftime("%Y-%m-%d--%H-%M")
    import_filename = f'lists/import_{current_date}.csv'

    export_columns = ['FIRST NAME', 'LAST NAME', 'EMAIL', 'ORGANIZATION', 'POSITION', 'YEAR'] + workshop_columns.to_list()
    export_df = confirmed_bpf_df[export_columns].sort_values(by=['LAST NAME', 'FIRST NAME'])

    export_df.to_csv(import_filename, index=False)

    print(f'Export sheet has been written to: {import_filename}')


prepare_import(invited_filename='lists/invited_2024-01-25--11-06.xlsx')
