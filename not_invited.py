import pandas as pd
from datetime import datetime


"""
Check from a excel list with names and email addresses which of the
email addresses have already been added (by comparing to a zkipster export).
It then creates a CSV with the people not yet in zkipster in the format that
it can directly be imported to zkipster.
"""

def check_invitations(invited_filename: str, compare_filename: str) -> pd.DataFrame:
    invited_df = pd.read_excel(invited_filename)
    compare_df = pd.read_excel(compare_filename)

    # Don't consider people in the list 'KOFF Advisory board 27/11', as those are unfortunately
    # duplicates, thus people invited to this list *need* to be also in another list if they want
    # to participate in the BPF.
    invited_df = invited_df[invited_df['GUESTLIST'] != 'KOFF Advisory board 27/11']

    invited_df['EMAIL'] = invited_df['EMAIL'].str.lower()
    compare_df['EMAIL'] = compare_df['EMAIL'].str.lower()

    not_invited = compare_df[~compare_df['EMAIL'].isin(invited_df['EMAIL'])]
    duplicates = compare_df[compare_df['EMAIL'].isin(invited_df['EMAIL'])]

    export_df = pd.read_csv('list-templates/CSVImportTemplate.csv')

    not_invited_export_df = export_df.copy()
    not_invited_export_df['First Name'] = not_invited['FIRST NAME'].str.strip()
    not_invited_export_df['Last Name'] = not_invited['LAST NAME'].str.strip()
    not_invited_export_df['Email'] = not_invited['EMAIL'].str.strip()
    #not_invited_export_df['Medienhaus'] = not_invited['COMPANY NAME'].str.strip()

    duplicates_export_df = export_df.copy()
    duplicates_export_df['First Name'] = duplicates['FIRST NAME'].str.strip()
    duplicates_export_df['Last Name'] = duplicates['LAST NAME'].str.strip()
    duplicates_export_df['Email'] = duplicates['EMAIL'].str.strip()
    #duplicates_export_df['Medienhaus'] = duplicates['COMPANY NAME'].str.strip()

    return not_invited_export_df, duplicates_export_df


not_invited_list, duplicates_list = check_invitations(
    invited_filename='lists/invited_2024-01-12--16-16.xlsx', 
    compare_filename='lists/to-invite.xlsx'
)

current_date = datetime.now().strftime("%Y-%m-%d--%H-%M")

print("The following guest were not yet invited:\n")
for _, row in not_invited_list.iterrows():
    print(f'{row["First Name"]} {row["Last Name"]} {row["Email"]}')

not_invited_csv_filename = f'lists/not-invited_{current_date}.csv'
not_invited_list.to_csv(not_invited_csv_filename, index=False)

print(f'\nNot invited guest were written to the following file: {not_invited_csv_filename}')

print("\n\nThe following guest are already in the list:\n")
for _, row in duplicates_list.iterrows():
    print(f'{row["First Name"]} {row["Last Name"]} {row["Email"]}')

duplicates_csv_filename = f'lists/duplicates_{current_date}.csv'
duplicates_list.to_csv(duplicates_csv_filename, index=False)

print(f'\nDuplicate guest were written to the following file: {duplicates_csv_filename}')
