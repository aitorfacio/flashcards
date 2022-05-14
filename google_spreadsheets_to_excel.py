import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient import discovery
from googleapiclient.errors import HttpError
import os
import pickle
import pandas as pd
import httplib2
import argparse

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID_input ='1ngQBQTJysleK5ATkWrh5E_mhELeXPo4DRRiZEaMq5Yw'
SAMPLE_RANGE_NAME = 'A1:AA1000'


def main(key=None, select_genders=False):
    discovery_url = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build(
        'sheets',
        'v4',
        http=httplib2.Http(),
        discoveryServiceUrl=discovery_url,
        developerKey=key)

    spreadsheet_id = SAMPLE_SPREADSHEET_ID_input#'1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
    range_name = SAMPLE_RANGE_NAME#Class Data!A2:E'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])


    if not values:
        print('No data found.')
    else:
        # create an excel file to save all the data from the google sheet, removing the first column
        df = pd.DataFrame(values[1:], columns=values[0])
        writer = pd.ExcelWriter('google_spreadsheets_to_excel.xlsx')
        if select_genders:
            maskulin_df = df[df['Alemán'].str.startswith('der ')]
            femenin_df = df[df['Alemán'].str.startswith('die ')]
            neutral_df = df[df['Alemán'].str.startswith('das ')]
            #select the values in df that are not in maskulin_df, femenin_df or neutral_df
            other_df = df[~df['Alemán'].isin(maskulin_df['Alemán']) & ~df['Alemán'].isin(femenin_df['Alemán'])
                          & ~df['Alemán'].isin(neutral_df['Alemán'])]
            # remove the column Etiquetas from all the dataframes
            maskulin_df = maskulin_df.drop(['Etiquetas'], axis=1)
            femenin_df = femenin_df.drop(['Etiquetas'], axis=1)
            neutral_df = neutral_df.drop(['Etiquetas'], axis=1)
            other_df = other_df.drop(['Etiquetas'], axis=1)

            # save them in separate sheets in the excel file if they are not empty
            if not maskulin_df.empty:
                maskulin_df.to_excel(writer, sheet_name='Maskulin', index=False)
            if not femenin_df.empty:
                femenin_df.to_excel(writer, sheet_name='Femenin', index=False)
            if not neutral_df.empty:
                neutral_df.to_excel(writer, sheet_name='Neutrum', index=False)
            if not other_df.empty:
                other_df.to_excel(writer, sheet_name='Other', index=False)

        else:
            df.to_excel(writer, 'All')
        writer.save()
        print('Data saved to google_spreadsheets_to_excel.xlsx')


if __name__ == '__main__':
    #define arguments
    parser = argparse.ArgumentParser(description='Get data from google sheets')
    parser.add_argument('--key', help='File containing Google API key', required=True)
    #add one argument to divide the data by gender (masculine, femenine, neuter and other)
    parser.add_argument('-g', help='Dividir las palabras por género.', action='store_true', default=False)

    args = parser.parse_args()
    api_file = args.key
    # read the api key from the file and pass it to the main method
    if os.path.isfile(api_file):
        with open(api_file, "r") as f:
            api_key = f.read()
            main(api_key, select_genders=args.g)
