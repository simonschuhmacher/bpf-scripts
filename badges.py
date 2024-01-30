import pandas as pd


"""
Create CSV files to print badges with InDesign. The script gets its data
from different sources, such as the zkipster guest list and additional
excel sheets for speakers, media people and the team.
The script makes also sure that no duplicates are present, i.e., that
speakers or team members are not in guest badges.
"""

def badges():
    guests_df = pd.read_excel('lists/invited_2024-01-22--15-17.xlsx')
    extraguests_df = pd.read_excel('badges_in/extra-guests.xlsx')
    speakers_df = pd.read_excel('badges_in/speakers.xlsx')
    media_df = pd.read_excel('badges_in/media.xlsx')
    team_df = pd.read_excel('badges_in/team.xlsx')

    export_dir = 'lists'

    side_event_lists = [
        'KOFF Advisory board 27/11',
        'Side Event: UN Peace Missions'
    ]

    # Guests
    guests_df = guests_df[~guests_df['GUESTLIST'].isin(side_event_lists)]
    guests_df = guests_df[guests_df['GUEST STATUS'] == 'Confirmed']
    guests_df = guests_df[(guests_df['ATTENDANCE'].isna()) | (guests_df['ATTENDANCE'] == 'I will travel to Basel to attend in person.')]
    guests_df = guests_df[~guests_df['EMAIL'].str.lower().isin(speakers_df['EMAIL'].str.lower())]
    guests_df = guests_df[~guests_df['EMAIL'].str.lower().isin(media_df['EMAIL'].str.lower())]
    guests_df = guests_df[~guests_df['EMAIL'].str.lower().isin(team_df['EMAIL'].str.lower())]

    extraguests_df = extraguests_df[(extraguests_df['EMAIL'].isna()) | (~extraguests_df['EMAIL'].fillna('').str.lower().isin(guests_df['EMAIL'].str.lower()))]
    guests_df = pd.concat([guests_df, extraguests_df])

    guests_export_df = guests_df[['FIRST NAME', 'LAST NAME', 'EMAIL', 'ORGANIZATION', 'POSITION']].sort_values(by=['LAST NAME', 'FIRST NAME'])

    # Speakers
    speakers_df = speakers_df[speakers_df['ONLINE'].isna()]
    speakers_export_df = speakers_df[['FIRST NAME', 'LAST NAME', 'EMAIL', 'ORGANIZATION', 'POSITION']].sort_values(by=['LAST NAME', 'FIRST NAME'])

    # Media
    media_export_df = media_df[['FIRST NAME', 'LAST NAME', 'EMAIL', 'ORGANIZATION', 'POSITION']].sort_values(by=['LAST NAME', 'FIRST NAME'])

    # Team
    team_df = team_df[~team_df['EMAIL'].str.lower().isin(speakers_df['EMAIL'].str.lower())]
    team_export_df = team_df[['FIRST NAME', 'LAST NAME', 'EMAIL', 'ORGANIZATION', 'POSITION']].sort_values(by=['LAST NAME', 'FIRST NAME'])
    

    # Export
    guests_export_filename = f'{export_dir}/badges-guests.csv'
    guests_export_df.to_csv(guests_export_filename, index=False, encoding='utf16')
    print(f'CSV for guest badges has been written to: {guests_export_filename}')

    speakers_export_filename = f'{export_dir}/badges-speakers.csv'
    speakers_export_df.to_csv(speakers_export_filename, index=False, encoding='utf16')
    print(f'CSV for speaker badges has been written to: {speakers_export_filename}')

    media_export_filename = f'{export_dir}/badges-media.csv'
    media_export_df.to_csv(media_export_filename, index=False, encoding='utf16')
    print(f'CSV for media badges has been written to: {media_export_filename}')

    team_export_filename = f'{export_dir}/badges-team.csv'
    team_export_df.to_csv(team_export_filename, index=False, encoding='utf16')
    print(f'CSV for team badges has been written to: {team_export_filename}')


badges()
