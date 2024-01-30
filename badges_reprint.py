import pandas as pd


"""
Create CSV files to print badges with InDesign. This is a simpler version
than `badges.py`, as it imports data from a single excel file and does not
check for duplicates.
"""

def badges():
    guests_df = pd.read_excel('lists/reprint_2024-01-24--15-17.xlsx')

    export_dir = 'lists'

    side_event_lists = [
        'KOFF Advisory board 27/11',
        'Side Event: UN Peace Missions',
        'Rebooking Health',
        'Confirmed virtual'
    ]

    # Guests
    guests_df = guests_df[~guests_df['GUESTLIST'].isin(side_event_lists)]
    guests_df = guests_df[guests_df['GUEST STATUS'] == 'Confirmed']
    guests_df = guests_df[(guests_df['ATTENDANCE'].isna()) | (guests_df['ATTENDANCE'] == 'I will travel to Basel to attend in person.')]

    guests_export_df = guests_df[['FIRST NAME', 'LAST NAME', 'ORGANIZATION', 'POSITION']].sort_values(by=['LAST NAME', 'FIRST NAME'])

    # Export
    guests_export_filename = f'{export_dir}/badges-guests.csv'
    guests_export_df.to_csv(guests_export_filename, index=False, encoding='utf16')
    print(f'CSV for guest badges has been written to: {guests_export_filename}')


badges()
