import win32com.client
import pandas as pd
from datetime import datetime, timedelta

def get_outlook_meetings(start_date, end_date):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        calendar = outlook.GetDefaultFolder(9)  # 9 indicates the calendar folder

        calendar_items = calendar.Items
        calendar_items.IncludeRecurrences = True  # Include recurring meetings
        calendar_items.Sort("[Start]")

        start_date = datetime.combine(start_date, datetime.min.time())
        end_date = datetime.combine(end_date, datetime.max.time())

        restriction = "[Start] >= '{}' AND [End] <= '{}'".format(
            start_date.strftime("%m/%d/%Y %H:%M %p"),
            end_date.strftime("%m/%d/%Y %H:%M %p")
        )

        meetings = calendar_items.Restrict(restriction)

        event_list = []

        for meeting in meetings:
            meeting_start = meeting.Start
            meeting_end = meeting.End

            # Convert to naive datetime for consistent comparison
            if meeting_start.tzinfo is not None:
                meeting_start = meeting_start.replace(tzinfo=None)
            if meeting_end.tzinfo is not None:
                meeting_end = meeting_end.replace(tzinfo=None)

            event_list.append({
                'Subject': meeting.Subject,
                'Start': meeting_start,
                'End': meeting_end,
                'Duration': (meeting_end - meeting_start).total_seconds() / 3600,  # duration in hours
                'Project': 'Unassigned'  # Default value
            })

        return pd.DataFrame(event_list)
    except Exception as e:
        print(f"An error occurred: {e}")
        return pd.DataFrame()

if __name__ == "__main__":
    start_date = datetime(2023, 8, 1)
    end_date = datetime(2023, 8, 7, 23, 59, 59)
    df = get_outlook_meetings(start_date, end_date)
    print(df)
