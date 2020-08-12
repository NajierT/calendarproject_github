from datetime import timedelta
from calendar_automation import get_calendar_service
import datefinder
from sheets_setup import sheets_service
from list_calendars import calendar_list

print("*Note: All events are set to a duration of 2 hours*")
print("*Any event on the spreadsheet missing a date or time will be skipped over*")
print("*Running the script more than once over the same values will create duplicate events*")
spreadsheet_id = input('\nThe spreadsheet ID can be found in the url of the Google Sheet after the "spreadsheets/d/"\n(Example: https://docs.google.com/spreadsheets/d/(SPREADSHEET ID HERE)/edit?usp=sharing)\nEnter the spreadsheet ID: ')
calendar_id = calendar_list()


def sheet_values(range1):
    service = sheets_service()
    # The A1 notation of the values to retrieve.
    ranges = range1.split(", ")
    value_render_option = "FORMATTED_VALUE"
    request = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, ranges=ranges, valueRenderOption=value_render_option)
    response = request.execute()
    return response["valueRanges"][0]["values"]


date = sheet_values(range1=input("Which cells contain the dates?(Ex: For cell A8 to cell A21 put A8:A21): "))
time = sheet_values(range1=input("Which cells contain the times?(Ex: For cell D8 to cell D21 put D8:D21): "))
opponent = sheet_values(range1=input("Which cells contain the opponents?(Ex: For cell B8 to cell B21 put B8:B21): "))
sport = input("Which sport?(Ex: Women's Soccer): ")


def main():
    for d, t, o in zip(date, time, opponent):
        # Separates jv/varsity times if applicable
        strtime = str(t).strip("[']")
        jv_var = strtime.split("/")
        if strtime.find("/") != -1:
            j = jv_var[0]
            v = jv_var[1]
        else:
            j = t
            v = t
        service = get_calendar_service()
        matches1 = list(datefinder.find_dates(str(d).strip("[']")+" "+str(j).strip("[']")))
        matches2 = list(datefinder.find_dates(str(d).strip("[']")+" "+str(v).strip("[']")))

        # Skips over any rows that don't have a date/time
        if matches1 == []:
            continue
        start = matches1[0].isoformat()
        end = (matches2[0] + timedelta(hours=2)).isoformat()

        # Puts "vs" opponent in the title of the event if it doesn't already have an @
        if str(o).find("@") != -1:
            home_away = " "
        else:
            home_away = " vs "

        event_result = service.events().insert(calendarId=calendar_id,
            body={
               "summary": sport + home_away + str(o).strip("[']"),
               "start": {"dateTime": start, "timeZone": 'America/New_York'},
               "end": {"dateTime": end, "timeZone": 'America/New_York'},
            }
        ).execute()

        print("created event")
        print("id: ", event_result['id'])
        print("summary: ", event_result['summary'])
        print("starts at: ", event_result['start']['dateTime'])
        print("ends at: ", event_result['end']['dateTime'])
    print("\nAll events added!")


main()
