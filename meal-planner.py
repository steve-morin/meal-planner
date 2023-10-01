import pandas as pd
import numpy as np
import os.path
from pathlib import Path
import datetime 
from datetime import timedelta, date
#from googleapiclient.discovery import build

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Scopes for Google API
SCOPES = ['https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/spreadsheets']

calendarId = '<INSERT ID>'

SPREADSHEET_ID = '<INSERT ID>'
RANGE = 'A1:G27'
DATE_LAST_RECIPE = ''
START_DAY = ''
NEXT_WEEK = ''
PREV_WEEK = ''

eligible_recipes = pd.DataFrame()

def get_credentials():

    my_local_file = os.path.join(os.path.dirname(__file__), 'calendar-quickstart.json')

    token_path = str(Path().resolve()) +'\\token.json'
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    # If credentials are valid then let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(my_local_file, SCOPES)
            creds = flow.run_local_server(port=0)
            # Save credentials for next run
            with open(token_path, 'w') as token:
                token.write(creds.to_json())
    return creds

def add_weekmenu_to_calendar(service, weekmenu_df, calendarId):
    for i, r in weekmenu_df.iterrows():
        event = {
        'summary': r.menu_item,
#        'description': r.description,
        'start': {
            'date': i.date().isoformat(),
            'timeZone': 'Canada/Mountain',
        },
        'end': {
            'date': i.date().isoformat(),
            'timeZone': 'Canada/Mountain',
        }
        #     'attendees': [
        #         {'email': 'email@example.com'},
        #     ],
        }
        print(event)
#        event = service.events().insert(calendarId=calendarId, body=event).execute()

def update_sheet(service, row_number, date, spreadsheetId):
    range = "F"  + str(row_number)
    values = [[date]]
    body = {'values' : values}
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=spreadsheetId
                                                    , range=range
                                                    , valueInputOption='USER_ENTERED'
                                                    , body=body).execute()

def choose_recipe(difficulty, idx, weekmenu_df, eligible_recipes):
#    choice_idx = np.random.choice(eligible_recipes.query("difficulty == '" + difficulty + "'" ).index.values)
    choice_idx = np.random.choice(eligible_recipes.query("menu_item == menu_item").index.values)
    weekmenu_df.loc[idx, 'menu_item'] = eligible_recipes.loc[choice_idx, 'menu_item']
    weekmenu_df.loc[idx, 'url'] = eligible_recipes.loc[choice_idx, 'url']
    eligible_recipes.drop(choice_idx, inplace=True)
    return choice_idx    

def get_recipes(service, creds):
    # get meal options from Google spreadsheet
    try:
        # Call the Sheets API
        sheet = service.spreadsheets()
        recipes_result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE).execute()
        recipes = recipes_result.get('values', [])
        if not recipes:
            print('No data found.')
            return

        recipes_df = pd.DataFrame.from_records(recipes[1:], columns=recipes[0])
        recipes_df.last_date_on_menu = pd.to_datetime(recipes_df.last_date_on_menu, dayfirst=True)
        recipes_df.set_index('row_number', inplace=True)
        eligible_recipes = recipes_df[ (recipes_df.last_date_on_menu < datetime.datetime.strptime(PREV_WEEK,'%Y-%m-%dT%H:%M:%SZ')) | (np.isnat(recipes_df.last_date_on_menu)) ]
        return recipes_df, eligible_recipes
    
    except HttpError as err:
        print(err)

def generate_weekmenu(service, events_df, eligible_recipes):
    weekmenu_df = events_df.copy()

    for i, r in events_df.iterrows():
        row_number = choose_recipe('difficult', i, weekmenu_df, eligible_recipes)
        update_sheet(service, row_number, i.strftime('%Y-%m-%d'), SPREADSHEET_ID)
    return weekmenu_df

def create_events_df():

    dates = list(pd.period_range(START_DAY, NEXT_WEEK, freq='D').values)
    date_arr = []
    for d in dates:
        date_arr.append(d.to_timestamp())
    dates_np = np.array(date_arr)

    events_df = pd.DataFrame()
    events_df['date'] = dates_np.tolist()
    events_df.reset_index(inplace=True)
    events_df['weekday'] = events_df.date.apply(lambda x: x.strftime('%A'))
    events_df.set_index('date', inplace=True)
    events_df.sort_index(inplace=True)

    return events_df

def get_date_last_event(service, calendarId):
    events_result = service.events().list(calendarId=calendarId
                                        , singleEvents=True
                                        , orderBy='startTime').execute()
    if (len(events_result['items'])>0):
        tmp = events_result.get('items', [])[-1]['start']
        date_last_event = events_result.get('items', [])[-1]['start']['date'][:10]
        date_last_event = datetime.datetime.strptime(date_last_event, '%Y-%m-%d').date()
    else:
        date_last_event = date.today()
    return date_last_event

def format_date(date):
    date_time = datetime.datetime.combine(date, datetime.datetime.min.time())
    date_time_utc = date_time.isoformat() + 'Z'
    return date_time_utc

def main():
    global DATE_LAST_RECIPE, START_DAY, NEXT_WEEK, PREV_WEEK 

    # Get credentials for Google Sheets and Calendar APIs
    creds = None
    creds = get_credentials()

    try:
        service_cal = build('calendar', 'v3', credentials=creds)
        service_sht = build('sheets', 'v4', credentials=creds)

        last_event_date = get_date_last_event(service_cal, calendarId)
        # Defining dates
        DATE_LAST_RECIPE = last_event_date
        START_DAY = DATE_LAST_RECIPE + timedelta(days=1)
        NEXT_WEEK = START_DAY + timedelta(days=6)
        PREV_WEEK = START_DAY + timedelta(days=-7)
        START_DAY = format_date(START_DAY)
        NEXT_WEEK = format_date(NEXT_WEEK)
        PREV_WEEK = format_date(PREV_WEEK)

        recipes_df, eligible_recipes = get_recipes(service_sht, creds)

        events = create_events_df()
        week_menu = generate_weekmenu(service_sht, events, eligible_recipes)    
        print(week_menu)
        add_weekmenu_to_calendar(service_sht, week_menu, calendarId)

    except HttpError as error:
        print('An error occurred: %s' % error)

    print('Process Complete')

if __name__ == '__main__':

    main()
