import win32com.client
import pandas as pd
from datetime import datetime, timedelta

def get_outlook_meetings():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        calendar = outlook.GetDefaultFolder(9)  # 9 indicates the calendar folder

        start_of_week = datetime.now().replace(tzinfo=None) - timedelta(days=datetime.now().weekday())
        end_of_week = start_of_week + timedelta(days=6)

        print(f"Start of week: {start_of_week}")
        print(f"End of week: {end_of_week}")

        restriction = "[Start] >= '{}' AND [End] <= '{}'".format(
            start_of_week.strftime("%m/%d/%Y %H:%M %p"),
            end_of_week.strftime("%m/%d/%Y %H:%M %p")
        )

        print(f"Restriction: {restriction}")

        meetings = calendar.Items.Restrict(restriction)
        meetings.Sort("[Start]")

        total_items = calendar.Items.Count
        print(f"Total items in calendar: {total_items}")

        event_list = []
        for meeting in meetings:
            meeting_start = meeting.Start
            meeting_end = meeting.End

            # Convert to naive datetime for consistent comparison
            if meeting_start.tzinfo is not None:
                meeting_start = meeting_start.replace(tzinfo=None)
            if meeting_end.tzinfo is not None:
                meeting_end = meeting_end.replace(tzinfo=None)

            if meeting_start >= start_of_week and meeting_end <= end_of_week:
                event_list.append({
                    'Subject': meeting.Subject,
                    'Start': meeting_start,
                    'End': meeting_end,
                    'Duration': (meeting_end - meeting_start).total_seconds() / 3600  # duration in hours
                })

        print(f"Number of meetings found: {len(event_list)}")
        return pd.DataFrame(event_list)
    except Exception as e:
        print(f"An error occurred: {e}")
        return pd.DataFrame()

# Fetch meetings from Outlook
df = get_outlook_meetings()

# Print the DataFrame
print(df)
