import pandas as pd
import re
from datetime import datetime


"""
Create a summary of registration numers (overall, on-site, virtual, workshps, etc.).
It also creates a excel file with sheets for workshop registrtions.
Probably the script that's most useful :)
"""

def print_edit_guest_link(row: pd.Series):     
    print(f'https://account.zkipster.com/Guests/Edit?InvitationId={row["ID"]} {row["FIRST NAME"]} {row["LAST NAME"]} {row["EMAIL"]}')
    print('----')

def workshop_registrations(invited_filename: str):
    invited_df = pd.read_excel(invited_filename)

    sheets: dict[str, pd.DataFrame] = {}

    side_event_lists = [
        'KOFF Advisory board 27/11',
        'Side Event: UN Peace Missions',
        'Rebooking Health'
    ]

    focus_zone_names = [
        'FOCUS ZONES REGISTRATION',
        'FOCUS ZONE I',
        'FOCUS ZONE II'
    ]

    confirmed_bpf_df = invited_df[~invited_df['GUESTLIST'].isin(side_event_lists)]
    confirmed_bpf_df = confirmed_bpf_df[confirmed_bpf_df['GUEST STATUS'] == 'Confirmed']

    sheets['All Confirmed'] = confirmed_bpf_df

    sheets['Invalid Focus Zones'] = confirmed_bpf_df[
        (confirmed_bpf_df['ATTENDANCE'] == 'I will attend virtually.') &
        (
            confirmed_bpf_df['FOCUS ZONE I'].fillna('').str.startswith('[on-site]') |
            confirmed_bpf_df['FOCUS ZONE II'].fillna('').str.startswith('[on-site]')
        )
    ]

    # sheets['CAS / ALUMNI NETWORKING LUNCH'] = invited_df[invited_df['CAS / ALUMNI NETWORKING LUNCH'] == 'I will participate']

    for side_event in side_event_lists:
        sheets[side_event] = invited_df[(invited_df['GUEST STATUS'] == 'Confirmed') & (invited_df['GUESTLIST'] == side_event)]

    for focus_zone_name in focus_zone_names:
        focus_zone_values = confirmed_bpf_df[focus_zone_name].dropna().unique()
        for focus_zone_value in focus_zone_values:
            sheets[f'{focus_zone_name}: {focus_zone_value}'] = confirmed_bpf_df[confirmed_bpf_df[focus_zone_name] == focus_zone_value]

    print('BPF Participants Summary:')
    print('------------------------')
    for sheet in sheets:
        sheet_df = sheets[sheet]
        onsite_df = sheet_df[sheet_df['ATTENDANCE'] == 'I will travel to Basel to attend in person.']
        virtual_df = sheet_df[sheet_df['ATTENDANCE'] == 'I will attend virtually.']
        unspecified_attendance = sheet_df.shape[0] - onsite_df.shape[0] - virtual_df.shape[0]
        
        print(sheet)
        print(f'on-site: {onsite_df.shape[0]}')
        print(f'virtual: {virtual_df.shape[0]}')
        print(f'unspecified: {unspecified_attendance}')
        print('-----')

    speakers_df = pd.read_excel('badges_in/speakers.xlsx')
    double_onsite_df = confirmed_bpf_df[
        (confirmed_bpf_df['ATTENDANCE'] == 'I will travel to Basel to attend in person.') &
        (confirmed_bpf_df['FOCUS ZONES REGISTRATION'].notna())
    ]
    speakers_in_guests_df = double_onsite_df[double_onsite_df['EMAIL'].str.lower().isin(speakers_df['EMAIL'].str.lower())]

    print(f'Double confirmed speakers in guestlist: {speakers_in_guests_df.shape[0]}')
    for _, row in speakers_in_guests_df.iterrows():
        print_edit_guest_link(row=row)

    current_date = datetime.now().strftime("%Y-%m-%d--%H-%M")
    summary_filename = f'lists/summary_{current_date}.xlsx'
    with pd.ExcelWriter(summary_filename) as writer:
        for sheet in sheets:
            sheet_name = re.sub(r'\W', '_', sheet)[:30]
            sheets[sheet].to_excel(writer, sheet_name=sheet_name, index=False)

    print(f'Summary has been written to: {summary_filename}')


workshop_registrations(invited_filename='lists/invited_2024-01-24--19-51.xlsx')
