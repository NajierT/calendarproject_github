from calendar_automation import get_calendar_service


def calendar_list():
    service = get_calendar_service()
    print('\nHere are the names of your calendars:\n')
    calendars_result = service.calendarList().list().execute()
    calendars = calendars_result.get('items')
    if not calendars:
        print('No calendars found.')
    for calendar in calendars:
        summary = calendar['summary']
        primary = "(Primary)" if calendar.get('primary') else ""
        print("%s\t%s" % (summary, primary))
    calendar_name = input("\nWhat is the name of the calendar you would like to use?: ")
    # Gets the calendar id given the calendar name:
    return str(next(item for item in calendars if item['summary'] == calendar_name)['id'])
