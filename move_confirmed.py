import requests
import pandas as pd


"""
Move guests to different lists based on the attendance set.
Also helps to identify which guests have no attendance specified.
This script is capable to automatically move guests to a different
guestlist by performing network requests reverse engineered from
the zkipster web UI. In order for this functionality to work,
the correct headers, cookies and request data needs to be present.
This data can be extracted from a similar network request being made
in Google Chrome. Then, a plugin like "Copy as Python Requests" can
be used to get the required data which needs to be copied into this
script at the correct positions. Note that headers and cookies will
only be valid for a short time (duration of validity of the session
in the browser)
"""

def move_to_list(guest_ids: list[str], guestlist_id: str):
    url = 'https://account.zkipster.com/Guests/BatchMoveToGuestList'
    # Update this
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding": "gzip, deflate, br",
        "Accept-Language": "en-US,en;q=0.9,de;q=0.8",
        "Cache-Control": "max-age=0",
        "Connection": "keep-alive",
        "Content-Length": "379",
        "Host": "account.zkipster.com",
        "Origin": "https://account.zkipster.com",
        "Referer": "https://account.zkipster.com/Guests/Index?EventId=your-event-id",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "sec-ch-ua": "\"Not_A Brand\";v=\"8\", \"Chromium\";v=\"120\", \"Google Chrome\";v=\"120\"",
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": "\"Windows\"",
    }
    # Add cookies here as found when copying request with "Copy as Python Requests" extension
    cookies = {
    }
    # Update this
    data = {
        "__RequestVerificationToken": "your-token-here",
        "EventId": "your-event-id-here",
        "SelectedGuestListId": guestlist_id,
        "Filter": "",
        "SearchTerm": "",
        "GuestListId": "",
        "EventId": "your-event-id-here",
        "ExportAll": "false",
        "SourceAll": "false",
        "ChangeAll": "false",
        "RemoveAll": "false",
        "ResendToAll": "false",
        "GuestIds": guest_ids,
    }

    response = requests.post(url=url, headers=headers, cookies=cookies, data=data)
    print(f'Response of {url}: {response}')


def print_edit_guest_link(row: pd.Series):     
    print(f'https://account.zkipster.com/Guests/Edit?InvitationId={row["ID"]} {row["FIRST NAME"]} {row["LAST NAME"]} {row["EMAIL"]}')
    print('----')


def move_reminder(invited_filename: str, from_guestlist: str):
    invited_df = pd.read_excel(invited_filename)

    side_event_lists = [
        'KOFF Advisory board 27/11',
        'Side Event: UN Peace Missions',
        'Rebooking Health'
    ]

    confirmed_bpf_df = invited_df[invited_df['GUEST STATUS'] == 'Confirmed']
    confirmed_bpf_df = confirmed_bpf_df[~confirmed_bpf_df['GUESTLIST'].isin(side_event_lists)]
    confirmed_bpf_df = confirmed_bpf_df[confirmed_bpf_df['GUESTLIST'] == from_guestlist]

    confirmed_unspecified_df = confirmed_bpf_df[confirmed_bpf_df['ATTENDANCE'].isna()]
    confirmed_onsite_df = confirmed_bpf_df[confirmed_bpf_df['ATTENDANCE'] == 'I will travel to Basel to attend in person.']
    confirmed_virtual_df = confirmed_bpf_df[confirmed_bpf_df['ATTENDANCE'] == 'I will attend virtually.']

    print(f'{confirmed_unspecified_df.shape[0]}/{confirmed_bpf_df.shape[0]} guests have no attendance specified:')
    print('-----------------------------------------------------')
    for _, row in confirmed_unspecified_df.iterrows():
        print(print_edit_guest_link(row))

    # move_to_list(guest_ids=confirmed_onsite_df["ID"].to_list(), guestlist_id='')
    # move_to_list(guest_ids=confirmed_virtual_df["ID"].to_list(), guestlist_id='')


move_reminder(invited_filename='lists/invited_2024-01-23--20-19.xlsx', from_guestlist='Guests')
